#!/usr/bin/env python3
"""
access_mcp_server.py
====================
MCP Server para leer y editar bases de datos Microsoft Access (.accdb/.mdb)
via COM automation (pywin32). Requiere Windows + Microsoft Access instalado.

Instalar dependencias:
    pip install mcp pywin32

Registrar en Claude Code (una de las dos formas):
    # Opcion A — global
    claude mcp add access -- python /ruta/al/access_mcp_server.py

    # Opcion B — solo este proyecto (crea .mcp.json en el directorio actual)
    claude mcp add --scope project access -- python /ruta/al/access_mcp_server.py

Flujo tipico para editar VBA:
    1. access_list_objects  → ver que modulos/forms existen
    2. access_get_code      → exportar el objeto a texto
    3. (Claude edita el texto)
    4. access_set_code      → reimportar el texto modificado
    5. access_close         → liberar Access (opcional)
"""

import asyncio
import atexit
import json
import logging
import os
import re
import sys
import tempfile
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import mcp.types as types
from mcp.server import Server
from mcp.server.stdio import stdio_server

# ---------------------------------------------------------------------------
# Logging — va a stderr para no contaminar el canal JSON-RPC de stdout
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    stream=sys.stderr,
)
log = logging.getLogger("access-mcp")

# ---------------------------------------------------------------------------
# Constantes Access COM
# ---------------------------------------------------------------------------
AC_TYPE: dict[str, int] = {
    "query":  1,   # acQuery
    "form":   2,   # acForm
    "report": 3,   # acReport
    "macro":  4,   # acMacro
    "module": 5,   # acModule
}

# ---------------------------------------------------------------------------
# Caches para reducir COM calls en sesiones largas
# ---------------------------------------------------------------------------
_vbe_code_cache: dict = {}        # "type:name" → texto completo del módulo VBE
_parsed_controls_cache: dict = {} # "form:name" / "report:name" → resultado _parse_controls()

# ---------------------------------------------------------------------------
# Sesion COM — singleton, mantiene Access vivo entre llamadas
# ---------------------------------------------------------------------------
class _Session:
    """
    Mantiene una instancia de Access.Application entre tool calls.
    Si se pide una BD distinta a la abierta, cierra la actual y abre la nueva.
    """
    _app: Optional[Any] = None
    _db_open: Optional[str] = None
    _cm_cache: dict = {}   # "type:name" → CodeModule COM object

    @classmethod
    def connect(cls, db_path: str) -> Any:
        resolved = str(Path(db_path).resolve())
        if cls._app is None:
            cls._launch()
        if cls._db_open != resolved:
            cls._switch(resolved)
        return cls._app

    @classmethod
    def _launch(cls) -> None:
        try:
            import win32com.client
        except ImportError:
            raise RuntimeError(
                "pywin32 no instalado. Ejecuta: pip install pywin32"
            )
        log.info("Lanzando Access.Application...")
        cls._app = win32com.client.Dispatch("Access.Application")
        cls._app.Visible = True   # necesario para que el VBE sea accesible via COM
        log.info("Access lanzado OK")

    @classmethod
    def _switch(cls, path: str) -> None:
        if not os.path.isfile(path):
            raise FileNotFoundError(f"No existe el fichero: {path}")
        if cls._db_open is not None:
            log.info("Cerrando BD anterior: %s", cls._db_open)
            try:
                cls._app.CloseCurrentDatabase()
            except Exception as e:
                log.warning("Error cerrando BD anterior: %s", e)
        log.info("Abriendo BD: %s", path)
        cls._app.OpenCurrentDatabase(path)
        cls._db_open = path
        # Limpiar caches al cambiar de BD
        cls._cm_cache.clear()
        _vbe_code_cache.clear()
        _parsed_controls_cache.clear()
        log.info("BD abierta OK")

    @classmethod
    def quit(cls) -> None:
        if cls._app is not None:
            log.info("Cerrando Access...")
            try:
                if cls._db_open:
                    cls._app.CloseCurrentDatabase()
                cls._app.Quit()
                log.info("Access cerrado OK")
            except Exception as e:
                log.warning("Error cerrando Access: %s", e)
            finally:
                cls._app = None
                cls._db_open = None
                cls._cm_cache.clear()
                _vbe_code_cache.clear()
                _parsed_controls_cache.clear()


atexit.register(_Session.quit)


# ---------------------------------------------------------------------------
# Helpers de ficheros temporales
# ---------------------------------------------------------------------------
def _read_tmp(path: str) -> tuple[str, str]:
    """
    Lee un fichero exportado por Access.
    Devuelve (contenido, encoding_usado).
    Detecta UTF-16 con BOM (formato habitual de .accdb) antes de intentar cp1252.
    """
    with open(path, "rb") as f:
        bom = f.read(2)
    if bom in (b"\xff\xfe", b"\xfe\xff"):
        with open(path, encoding="utf-16") as f:
            return f.read(), "utf-16"
    for enc in ("utf-8-sig", "cp1252", "utf-8"):
        try:
            with open(path, encoding=enc) as f:
                return f.read(), enc
        except UnicodeDecodeError:
            continue
    with open(path, encoding="utf-8", errors="replace") as f:
        return f.read(), "utf-8"


def _write_tmp(path: str, content: str, encoding: str = "utf-16") -> None:
    """
    Escribe contenido para que Access lo lea con LoadFromText.
    Por defecto utf-16 (Access .accdb espera UTF-16LE con BOM).
    """
    with open(path, "w", encoding=encoding, errors="replace") as f:
        f.write(content)


# ---------------------------------------------------------------------------
# Filtrado de secciones binarias en forms/reports
# ---------------------------------------------------------------------------
# Secciones Begin...End que son blobs binarios irrelevantes para editar VBA/lógica.
# PrtMip + PrtDevMode representan el 95 % del tamaño del fichero exportado.
_BINARY_SECTIONS: frozenset[str] = frozenset({
    "PrtMip", "PrtDevMode", "PrtDevModeW",
    "PrtDevNames", "PrtDevNamesW",
    "RecSrcDt", "GUID", "NameMap",
})


def _strip_binary_sections(text: str) -> str:
    """
    Elimina las secciones binarias de un export de formulario/informe Access.
    Reduce el tamaño ~20x (de ~300 KB a ~15 KB) sin afectar al VBA ni a los controles.
    También elimina la línea Checksum (Access la recalcula al importar).
    """
    lines = text.splitlines(keepends=True)
    result: list[str] = []
    skip_depth = 0      # > 0 mientras estamos dentro de un bloque binario Begin...End
    skip_indent = ""    # indentación de la línea Begin que estamos saltando

    for line in lines:
        rstripped = line.rstrip("\r\n")
        stripped = rstripped.lstrip()
        indent = rstripped[: len(rstripped) - len(stripped)]

        if skip_depth > 0:
            # ¿Es el End de cierre al mismo nivel de indentación?
            if stripped == "End" and indent == skip_indent:
                skip_depth -= 1
            continue  # salta la línea (parte del bloque binario)

        # Línea Checksum a nivel raíz
        if re.match(r"^Checksum\s*=\s*", rstripped):
            continue

        # ¿Empieza un bloque binario?
        m = re.match(r"^(\s*)(\w+)\s*=\s*Begin\s*$", rstripped)
        if m and m.group(2) in _BINARY_SECTIONS:
            skip_indent = m.group(1)
            skip_depth = 1
            continue

        result.append(line)

    return "".join(result)


def _extract_binary_blocks(text: str) -> dict[str, str]:
    """
    Extrae los bloques binarios Begin...End del export original de un form/report.
    Devuelve {nombre_seccion: texto_completo_del_bloque}.
    """
    blocks: dict[str, str] = {}
    lines = text.splitlines(keepends=True)
    i = 0
    while i < len(lines):
        line = lines[i]
        rstripped = line.rstrip("\r\n")
        stripped = rstripped.lstrip()
        indent = rstripped[: len(rstripped) - len(stripped)]

        m = re.match(r"^(\s*)(\w+)\s*=\s*Begin\s*$", rstripped)
        if m and m.group(2) in _BINARY_SECTIONS:
            section = m.group(2)
            block_lines = [line]
            j = i + 1
            while j < len(lines):
                bl = lines[j]
                bl_r = bl.rstrip("\r\n")
                bl_s = bl_r.lstrip()
                bl_indent = bl_r[: len(bl_r) - len(bl_s)]
                block_lines.append(bl)
                if bl_s == "End" and bl_indent == indent:
                    break
                j += 1
            blocks[section] = "".join(block_lines)
            i = j + 1
            continue

        i += 1

    return blocks


def _restore_binary_sections(app: Any, object_type: str, name: str, new_code: str) -> str:
    """
    Re-inyecta las secciones binarias (PrtMip, PrtDevMode, etc.) desde el export
    actual del objeto, antes de llamar a LoadFromText con el código editado.
    Si el objeto no existe aún, devuelve new_code sin modificar.
    """
    fd, tmp = tempfile.mkstemp(suffix=".txt", prefix="access_mcp_orig_")
    os.close(fd)
    try:
        try:
            app.SaveAsText(AC_TYPE[object_type], name, tmp)
        except Exception:
            log.info("_restore_binary_sections: '%s' no existe aún, se importa sin secciones binarias", name)
            return new_code
        original, _enc = _read_tmp(tmp)
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass

    blocks = _extract_binary_blocks(original)
    if not blocks:
        return new_code

    # Inyectar los bloques justo antes de "End Form" / "End Report"
    _end_re = re.compile(r"^\s*End\s+(?:Form|Report)\s*$")
    _begin_re = re.compile(r"^\s*Begin\s+(?:Form|Report)\s*$")
    lines = new_code.splitlines(keepends=True)
    result: list[str] = []
    in_top_form = False
    injected = False

    for line in lines:
        stripped = line.strip()

        if _begin_re.match(stripped):
            in_top_form = True

        if in_top_form and not injected and _end_re.match(stripped):
            for block_text in blocks.values():
                result.append(block_text)
                if not block_text.endswith("\n"):
                    result.append("\n")
            injected = True
            in_top_form = False

        result.append(line)

    return "".join(result)


# ---------------------------------------------------------------------------
# VBE CodeModule — operaciones línea a línea (sin export/import de fichero)
# ---------------------------------------------------------------------------
# Prefijos que usa Access en el árbol VBE para forms e informes
_VBE_PREFIX: dict[str, str] = {
    "module": "",
    "form":   "Form_",
    "report": "Report_",
}


def _get_code_module(app: Any, object_type: str, object_name: str) -> Any:
    """
    Devuelve el CodeModule VBE del componente indicado.
    Cachea el objeto COM para evitar 3 calls en cadena en cada tool VBE.
    Requiere 'Confiar en el acceso al modelo de objetos de proyectos VBA'
    habilitado en las opciones de confianza de Access.
    """
    if object_type not in _VBE_PREFIX:
        raise ValueError(
            f"object_type '{object_type}' no soporta VBE. Usa 'module', 'form' o 'report'."
        )
    cache_key = f"{object_type}:{object_name}"
    cm = _Session._cm_cache.get(cache_key)
    if cm is not None:
        return cm
    component_name = _VBE_PREFIX[object_type] + object_name
    try:
        project = app.VBE.VBProjects(1)
        component = project.VBComponents(component_name)
        cm = component.CodeModule
        _Session._cm_cache[cache_key] = cm
        return cm
    except Exception as exc:
        raise RuntimeError(
            f"No se pudo acceder al CodeModule '{component_name}'. "
            f"¿Está habilitado 'Confiar en el acceso al modelo de objetos de proyectos VBA' "
            f"en las opciones de confianza de Access?\nError: {exc}"
        )


def _cm_all_code(cm: Any, cache_key: str) -> str:
    """
    Devuelve el texto completo de un CodeModule usando _vbe_code_cache.
    En una sesión con múltiples tools sobre el mismo módulo, la lectura COM
    completa (cm.Lines) se hace una sola vez; las siguientes llamadas usan el cache.
    """
    if cache_key not in _vbe_code_cache:
        total = cm.CountOfLines
        _vbe_code_cache[cache_key] = cm.Lines(1, total) if total > 0 else ""
    return _vbe_code_cache[cache_key]


def _text_matches(
    needle: str, haystack: str, match_case: bool, use_regex: bool,
) -> bool:
    """Compara needle contra haystack: substring plano o regex."""
    if use_regex:
        flags = 0 if match_case else re.IGNORECASE
        return re.search(needle, haystack, flags) is not None
    if not match_case:
        return needle.lower() in haystack.lower()
    return needle in haystack


def ac_vbe_get_lines(
    db_path: str, object_type: str, object_name: str,
    start_line: int, count: int
) -> str:
    """Lee un rango de líneas sin exportar el módulo entero."""
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    cache_key = f"{object_type}:{object_name}"
    all_code = _cm_all_code(cm, cache_key)
    all_lines = all_code.splitlines()
    total = len(all_lines)
    if start_line < 1 or start_line > total:
        raise ValueError(f"start_line {start_line} fuera de rango (1-{total})")
    actual = min(count, total - start_line + 1)
    return "\n".join(all_lines[start_line - 1 : start_line - 1 + actual])


def ac_vbe_get_proc(
    db_path: str, object_type: str, object_name: str, proc_name: str
) -> dict:
    """
    Devuelve información y código de un procedimiento concreto.
    Mucho más eficiente que ac_get_code cuando solo interesa una función.
    Devuelve: start_line, body_line, count, code.
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    try:
        start = cm.ProcStartLine(proc_name, 0)   # 0 = vbext_pk_Proc (COM call — rápido)
        body  = cm.ProcBodyLine(proc_name, 0)
        count = cm.ProcCountLines(proc_name, 0)
    except Exception as exc:
        raise RuntimeError(
            f"Procedimiento '{proc_name}' no encontrado en '{object_name}': {exc}"
        )
    # Extraer el texto desde el cache en vez de un cm.Lines adicional
    cache_key = f"{object_type}:{object_name}"
    all_lines = _cm_all_code(cm, cache_key).splitlines()
    code = "\n".join(all_lines[start - 1 : start - 1 + count])
    return {
        "proc_name":  proc_name,
        "start_line": start,
        "body_line":  body,
        "count":      count,
        "code":       code,
    }


def ac_vbe_module_info(
    db_path: str, object_type: str, object_name: str
) -> dict:
    """
    Devuelve el total de líneas y la lista de procedimientos con sus posiciones.
    Útil como índice rápido antes de editar, sin descargar el código completo.
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    cache_key = f"{object_type}:{object_name}"
    all_code = _cm_all_code(cm, cache_key)
    all_lines = all_code.splitlines()
    total = len(all_lines)
    procs: list[dict] = []
    if total > 0:
        seen: set[str] = set()
        for i, raw_line in enumerate(all_lines, start=1):
            m = re.match(
                r'^(?:Public\s+|Private\s+|Friend\s+)?'
                r'(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+(\w+)',
                raw_line.strip(), re.IGNORECASE,
            )
            if m:
                pname = m.group(1)
                if pname in seen:
                    continue
                seen.add(pname)
                try:
                    pstart = cm.ProcStartLine(pname, 0)
                    body   = cm.ProcBodyLine(pname, 0)
                    pcount = cm.ProcCountLines(pname, 0)
                    # Clamp count para no exceder total_lines
                    pcount = min(pcount, total - pstart + 1)
                    procs.append({"name": pname, "start_line": pstart,
                                  "body_line": body, "count": pcount})
                except Exception:
                    procs.append({"name": pname, "start_line": i})
    return {"total_lines": total, "procs": procs}


def ac_vbe_replace_lines(
    db_path: str, object_type: str, object_name: str,
    start_line: int, count: int, new_code: str
) -> str:
    """
    Reemplaza 'count' líneas a partir de 'start_line' con 'new_code'.
    - count=0 → inserción pura (no borra nada).
    - new_code='' → borrado puro (no inserta nada).
    new_code puede ser multilínea (\\n o \\r\\n).
    Devuelve el estado + preview del código insertado para evitar un get_proc adicional.
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    total = cm.CountOfLines
    # Validar límites
    if start_line < 1 or start_line > total + 1:
        raise ValueError(
            f"start_line {start_line} fuera de rango (1–{total})"
        )
    clamped = False
    if count > 0:
        max_count = total - start_line + 1
        if count > max_count:
            count = max_count
            clamped = True
        cm.DeleteLines(start_line, count)
    inserted = 0
    if new_code:
        # Access VBA espera \r\n como separador de líneas
        normalized = new_code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
        cm.InsertLines(start_line, normalized)
        inserted = len(new_code.splitlines())
    # Invalidar cache de texto (el módulo cambió)
    cache_key = f"{object_type}:{object_name}"
    _vbe_code_cache.pop(cache_key, None)
    new_total = cm.CountOfLines
    end = start_line + count - 1 if count > 0 else start_line
    clamp_note = " (count ajustado al límite del módulo)" if clamped else ""
    status = (
        f"OK: líneas {start_line}–{end} reemplazadas "
        f"({count} eliminadas, {inserted} insertadas){clamp_note} "
        f"→ módulo ahora tiene {new_total} líneas"
    )
    if new_code:
        lines = new_code.splitlines()
        preview = (
            new_code if len(lines) <= 60
            else "\n".join(lines[:60]) + f"\n[... +{len(lines) - 60} líneas]"
        )
        return f"{status}\n\n{preview}"
    return status


def ac_vbe_find(
    db_path: str, object_type: str, object_name: str,
    search_text: str, match_case: bool = False, use_regex: bool = False,
) -> dict:
    """
    Busca texto (o regex) en un módulo y devuelve todas las líneas que coinciden.
    Usa _vbe_code_cache para evitar releer el módulo si ya fue leído.
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    cache_key = f"{object_type}:{object_name}"
    all_code = _cm_all_code(cm, cache_key)
    if not all_code:
        return {"found": False, "match_count": 0, "matches": []}
    matches: list[dict] = []
    for i, raw_line in enumerate(all_code.splitlines(), start=1):
        if _text_matches(search_text, raw_line, match_case, use_regex):
            matches.append({"line": i, "content": raw_line.rstrip("\r")})
    return {"found": bool(matches), "match_count": len(matches), "matches": matches}


def ac_vbe_search_all(
    db_path: str, search_text: str, match_case: bool = False,
    max_results: int = 100, use_regex: bool = False,
) -> dict:
    """
    Busca texto (o regex) en TODOS los módulos VBA (modules, forms, reports) de la BD.
    Devuelve {total_matches, results: [...], truncated?: bool}.
    """
    app = _Session.connect(db_path)
    objects = ac_list_objects(db_path, "all")
    results: list[dict] = []
    total = 0
    truncated = False

    for obj_type in ("module", "form", "report"):
        if truncated:
            break
        for obj_name in objects.get(obj_type, []):
            if truncated:
                break
            try:
                cm = _get_code_module(app, obj_type, obj_name)
                cache_key = f"{obj_type}:{obj_name}"
                all_code = _cm_all_code(cm, cache_key)
                if not all_code:
                    continue
                obj_matches: list[dict] = []
                for i, raw_line in enumerate(all_code.splitlines(), start=1):
                    if _text_matches(search_text, raw_line, match_case, use_regex):
                        obj_matches.append({"line": i, "content": raw_line.rstrip("\r")})
                        total += 1
                        if total >= max_results:
                            truncated = True
                            break
                if obj_matches:
                    results.append({
                        "object_type": obj_type,
                        "object_name": obj_name,
                        "matches": obj_matches,
                    })
            except Exception:
                continue  # skip objects without accessible CodeModule

    out: dict = {"total_matches": total, "results": results}
    if truncated:
        out["truncated"] = True
    return out


def ac_search_queries(
    db_path: str, search_text: str, match_case: bool = False,
    max_results: int = 100, use_regex: bool = False,
) -> dict:
    """
    Busca texto (o regex) dentro del SQL de TODAS las queries (consultas) de la BD.
    Devuelve {total_matches, results: [{query_name, sql}], truncated?: bool}.
    """
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    results: list[dict] = []
    total = 0
    for qd in db.QueryDefs:
        name = qd.Name
        if name.startswith("~"):  # skip internal/temp queries
            continue
        sql = qd.SQL
        if _text_matches(search_text, sql, match_case, use_regex):
            results.append({"query_name": name, "sql": sql.strip()})
            total += 1
            if total >= max_results:
                break
    out: dict = {"total_matches": total, "results": results}
    if total >= max_results:
        out["truncated"] = True
    return out


# ---------------------------------------------------------------------------
# Find usages — cross-reference search
# ---------------------------------------------------------------------------
_CONTROL_SEARCH_PROPS = frozenset({
    "ControlSource", "RecordSource", "RowSource", "DefaultValue", "ValidationRule",
})


def ac_find_usages(
    db_path: str, search_text: str, match_case: bool = False,
    max_results: int = 200, use_regex: bool = False,
) -> dict:
    """
    Busca un nombre (funcion, tabla, campo, variable) en VBA, queries y
    propiedades de controles de forms/reports. Devuelve resultados agrupados.
    Reutiliza ac_vbe_search_all y ac_search_queries para VBA y queries.
    """
    # 1. VBA matches — delega en ac_vbe_search_all
    vba_result = ac_vbe_search_all(
        db_path, search_text, match_case, max_results, use_regex,
    )
    # Aplanar: de [{object_type, object_name, matches: [{line, content}]}] a lista plana
    vba_matches: list[dict] = []
    for group in vba_result["results"]:
        for m in group["matches"]:
            vba_matches.append({
                "object_type": group["object_type"],
                "object_name": group["object_name"],
                "line": m["line"],
                "content": m["content"],
            })
    total = len(vba_matches)
    truncated = vba_result.get("truncated", False)

    # 2. Query matches — delega en ac_search_queries
    query_matches: list[dict] = []
    if not truncated:
        remaining = max_results - total
        qry_result = ac_search_queries(
            db_path, search_text, match_case, remaining, use_regex,
        )
        query_matches = qry_result["results"]
        total += qry_result["total_matches"]
        truncated = qry_result.get("truncated", False)

    # 3. Control property matches — busca en exports de forms/reports
    control_matches: list[dict] = []
    if not truncated:
        app = _Session.connect(db_path)
        objects = ac_list_objects(db_path, "all")
        for obj_type in ("form", "report"):
            if truncated:
                break
            for obj_name in objects.get(obj_type, []):
                if truncated:
                    break
                try:
                    fd, tmp = tempfile.mkstemp(suffix=".txt", prefix="access_mcp_")
                    os.close(fd)
                    try:
                        app.SaveAsText(AC_TYPE[obj_type], obj_name, tmp)
                        raw_text, _enc = _read_tmp(tmp)
                    finally:
                        try:
                            os.unlink(tmp)
                        except OSError:
                            pass
                    for line in raw_text.splitlines():
                        stripped = line.strip()
                        for prop in _CONTROL_SEARCH_PROPS:
                            if stripped.startswith(prop + " ="):
                                value_part = stripped[len(prop) + 2:].strip()
                                if _text_matches(search_text, value_part, match_case, use_regex):
                                    control_matches.append({
                                        "object_type": obj_type,
                                        "object_name": obj_name,
                                        "property": prop,
                                        "value": value_part,
                                    })
                                    total += 1
                                    if total >= max_results:
                                        truncated = True
                                    break
                except Exception:
                    continue

    out: dict = {
        "search_text": search_text,
        "vba_matches": vba_matches,
        "query_matches": query_matches,
        "control_matches": control_matches,
        "total_matches": total,
    }
    if truncated:
        out["truncated"] = True
    return out


def ac_vbe_replace_proc(
    db_path: str, object_type: str, object_name: str,
    proc_name: str, new_code: str
) -> str:
    """
    Reemplaza un procedimiento completo (Sub/Function/Property) por nombre.
    Calcula los límites automáticamente via COM (ProcStartLine/ProcCountLines),
    eliminando errores de cálculo del caller.
    Si new_code está vacío, elimina el procedimiento.
    """
    app = _Session.connect(db_path)

    # Si el form/report está abierto en Design view (tras ac_set_control_props etc.),
    # cerrarlo primero para evitar conflictos COM con el VBE ("Error catastrófico")
    if object_type in ("form", "report"):
        ac_obj_type = _AC_FORM if object_type == "form" else _AC_REPORT
        try:
            app.DoCmd.Close(ac_obj_type, object_name, _AC_SAVE_YES)
            log.info("ac_vbe_replace_proc: cerrado '%s' en Design view antes de acceder VBE", object_name)
        except Exception:
            pass  # no estaba abierto — OK

    # Invalidar cm_cache por si el CodeModule quedó stale tras operación de diseño
    cache_key = f"{object_type}:{object_name}"
    _Session._cm_cache.pop(cache_key, None)

    cm = _get_code_module(app, object_type, object_name)
    try:
        start = cm.ProcStartLine(proc_name, 0)
        count = cm.ProcCountLines(proc_name, 0)
    except Exception as exc:
        raise RuntimeError(
            f"Procedimiento '{proc_name}' no encontrado en '{object_name}': {exc}"
        )
    # Clamp count al total real del módulo (ProcCountLines puede inflar el último proc)
    total = cm.CountOfLines
    count = min(count, total - start + 1)
    # Borrar procedimiento viejo
    cm.DeleteLines(start, count)
    # Insertar nuevo código (si hay)
    inserted = 0
    if new_code:
        normalized = new_code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
        cm.InsertLines(start, normalized)
        inserted = len(new_code.splitlines())
    # Invalidar cache
    cache_key = f"{object_type}:{object_name}"
    _vbe_code_cache.pop(cache_key, None)
    new_total = cm.CountOfLines
    action = "reemplazado" if new_code else "eliminado"
    status = (
        f"OK: proc '{proc_name}' {action} "
        f"({count} eliminadas, {inserted} insertadas) "
        f"→ módulo ahora tiene {new_total} líneas"
    )
    if new_code:
        lines = new_code.splitlines()
        preview = (
            new_code if len(lines) <= 60
            else "\n".join(lines[:60]) + f"\n[... +{len(lines) - 60} líneas]"
        )
        return f"{status}\n\n{preview}"
    return status


def ac_vbe_append(
    db_path: str, object_type: str, object_name: str,
    new_code: str
) -> str:
    """
    Añade código al final de un módulo VBA.
    Más seguro que replace_lines para insertar nuevas funciones
    sin necesidad de calcular números de línea.
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    total = cm.CountOfLines
    normalized = new_code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    cm.InsertLines(total + 1, normalized)
    inserted = len(new_code.splitlines())
    cache_key = f"{object_type}:{object_name}"
    _vbe_code_cache.pop(cache_key, None)
    new_total = cm.CountOfLines
    return f"OK: {inserted} líneas añadidas al final → módulo ahora tiene {new_total} líneas"


# ---------------------------------------------------------------------------
# Control-level — parseo del texto de formulario/informe control a control
# ---------------------------------------------------------------------------
_CTRL_TYPE: dict[int, str] = {
    100: "Label",
    101: "Rectangle",
    102: "Line",
    103: "Image",
    104: "CommandButton",
    105: "OptionButton",
    106: "CheckBox",
    107: "OptionGroup",
    108: "BoundObjectFrame",
    109: "TextBox",
    110: "ListBox",
    111: "ComboBox",
    112: "SubForm",
    113: "ObjectFrame",
    114: "PageBreak",
    118: "Page",
    119: "TabControl",
    122: "Attachment",
    124: "NavigationButton",
    125: "NavigationControl",
    126: "WebBrowser",
}


def _parse_controls(form_text: str) -> dict:
    """
    Parsea el texto exportado de un form/report y extrae los bloques de controles.
    Devuelve un dict con:
      controls       — lista de controles con sus propiedades y posición en el texto
      form_indent    — indentación de la línea "Begin Form/Report"
      ctrl_indent    — (legacy, se mantiene para compatibilidad) indent del primer control encontrado
      form_begin_idx — índice 0-based de la línea "Begin Form/Report"
      form_end_idx   — índice 0-based del "End" que cierra el bloque Form/Report

    Estructura del export de Access:
      Begin Form              ← form level
          Begin               ← defaults block (contiene Begin Label, Begin CommandButton con props default)
          End
          Begin Section       ← sección (Detail, FormHeader, FormFooter)
              ...
              Begin           ← contenedor de controles dentro de la sección
                  Begin Label ← CONTROL REAL (tiene Name =, ControlType =, etc.)
                  End
                  Begin CommandButton
                  End
              End
          End
          Begin ClassModule   ← código VBA del form
          End
      End Form

    El parser busca controles DENTRO de las secciones, identificándolos por tener
    un tipo conocido (Begin <TypeName>) donde TypeName es un valor de _CTRL_TYPE.
    """
    lines = form_text.splitlines(keepends=True)
    result: dict = {
        "controls": [],
        "form_indent": "",
        "ctrl_indent": "",
        "form_begin_idx": -1,
        "form_end_idx": -1,
    }

    # Conjunto de nombres de tipo para detección rápida
    ctrl_type_names = {v for v in _CTRL_TYPE.values()}

    # 1. Localizar "Begin Form" o "Begin Report"
    for i, line in enumerate(lines):
        s = line.rstrip("\r\n").lstrip()
        if re.match(r"^Begin\s+(Form|Report)\s*$", s, re.IGNORECASE):
            raw = line.rstrip("\r\n")
            result["form_indent"] = raw[: len(raw) - len(raw.lstrip())]
            result["form_begin_idx"] = i
            break

    if result["form_begin_idx"] == -1:
        return result

    form_begin = result["form_begin_idx"]

    # 2. Encontrar el "End" que cierra el bloque Form/Report (rastreo de profundidad)
    #    Importante: detectar tanto "Begin ..." como "Property = Begin" (ej: NameMap = Begin)
    #    para que sus "End" correspondientes no desbalanceen el contador de profundidad.
    depth = 0
    for i in range(form_begin, len(lines)):
        s = lines[i].rstrip("\r\n").lstrip()
        if re.match(r"^Begin\b", s) or re.match(r"^\w+\s*=\s*Begin\s*$", s):
            depth += 1
        elif s == "End":
            depth -= 1
            if depth == 0:
                result["form_end_idx"] = i
                break

    if result["form_end_idx"] == -1:
        return result

    # 3. Escanear TODOS los bloques "Begin <TypeName>" dentro del form/report
    #    donde TypeName coincide con un tipo de control conocido.
    #    Los controles pueden estar a cualquier profundidad dentro de secciones.
    i = form_begin + 1
    while i < result["form_end_idx"]:
        raw = lines[i].rstrip("\r\n")
        s = raw.lstrip()
        indent = raw[: len(raw) - len(s)]

        # Saltar ClassModule — no contiene controles, solo VBA
        if re.match(r"^Begin\s+ClassModule\s*$", s, re.IGNORECASE):
            break

        # Detectar "Begin <TypeName>" donde TypeName es un tipo de control conocido
        m_ctrl = re.match(r"^Begin\s+(\w+)\s*$", s)
        if m_ctrl and m_ctrl.group(1) in ctrl_type_names:
            ctrl_start = i
            block: list[str] = [lines[i]]
            props: dict[str, str] = {}
            blk_depth = 1
            ctrl_end = i
            j = i + 1
            while j < len(lines):
                bl = lines[j]
                bl_r = bl.rstrip("\r\n")
                bl_s = bl_r.lstrip()
                block.append(bl)
                # Solo parsear propiedades al nivel top del control (depth == 1)
                if blk_depth == 1:
                    m_prop = re.match(r"^(\w+)\s*=(.*)", bl_s)
                    if m_prop:
                        props[m_prop.group(1)] = m_prop.group(2).strip().strip('"')
                if re.match(r"^Begin\b", bl_s):
                    blk_depth += 1
                elif bl_s == "End":
                    blk_depth -= 1
                    if blk_depth == 0:
                        ctrl_end = j
                        break
                j += 1

            name = props.get("Name", props.get("ControlName", ""))
            try:
                ctype = int(props.get("ControlType", -1))
            except (ValueError, TypeError):
                ctype = -1

            # Guardar ctrl_indent del primer control encontrado (legacy compat)
            if not result["ctrl_indent"] and name:
                result["ctrl_indent"] = indent

            result["controls"].append({
                "name":           name,
                "control_type":   ctype,
                "type_name":      _CTRL_TYPE.get(ctype, m_ctrl.group(1)),
                "caption":        props.get("Caption", ""),
                "control_source": props.get("ControlSource", ""),
                "left":           props.get("Left", ""),
                "top":            props.get("Top", ""),
                "width":          props.get("Width", ""),
                "height":         props.get("Height", ""),
                "visible":        props.get("Visible", ""),
                "start_line":     ctrl_start + 1,  # 1-based
                "end_line":       ctrl_end + 1,     # 1-based inclusive
                "raw_block":      "".join(block),
            })
            i = ctrl_end + 1
            continue

        i += 1

    return result


def _get_parsed_controls(db_path: str, object_type: str, object_name: str) -> dict:
    """
    Devuelve el resultado de _parse_controls usando _parsed_controls_cache.
    Si no está en cache, exporta y parsea (y guarda en cache para futuras llamadas).
    """
    cache_key = f"{object_type}:{object_name}"
    if cache_key not in _parsed_controls_cache:
        text = ac_get_code(db_path, object_type, object_name)
        _parsed_controls_cache[cache_key] = _parse_controls(text)
    return _parsed_controls_cache[cache_key]


def ac_list_controls(db_path: str, object_type: str, object_name: str) -> dict:
    if object_type not in ("form", "report"):
        raise ValueError("ac_list_controls solo admite object_type 'form' o 'report'")
    parsed = _get_parsed_controls(db_path, object_type, object_name)
    controls = [
        {k: v for k, v in c.items() if k != "raw_block"}
        for c in parsed["controls"]
        if c.get("name", "").strip()  # excluir controles sin nombre
    ]
    return {
        "count": len(controls),
        "controls": controls,
    }


def ac_get_control(
    db_path: str, object_type: str, object_name: str, control_name: str
) -> dict:
    """
    Devuelve la definición completa (raw_block) de un control concreto por nombre.
    El raw_block puede pasarse modificado a ac_set_control para actualizar el control.
    """
    if object_type not in ("form", "report"):
        raise ValueError("ac_get_control solo admite object_type 'form' o 'report'")
    parsed = _get_parsed_controls(db_path, object_type, object_name)
    for c in parsed["controls"]:
        if c["name"].lower() == control_name.lower():
            return c
    names = [c["name"] for c in parsed["controls"]]
    raise ValueError(
        f"Control '{control_name}' no encontrado en '{object_name}'. "
        f"Controles disponibles: {names}"
    )


# ---------------------------------------------------------------------------
# Control COM — CreateControl / DeleteControl / set properties in design mode
# ---------------------------------------------------------------------------
_AC_DESIGN  = 1   # acDesign / acViewDesign
_AC_FORM    = 2   # acForm   (para DoCmd.Close/Save)
_AC_REPORT  = 3   # acReport (para DoCmd.Close/Save)
_AC_SAVE_YES = 1  # acSaveYes

# Mapa inverso nombre → número de tipo de control
_CTRL_TYPE_BY_NAME: dict[str, int] = {v.lower(): k for k, v in _CTRL_TYPE.items()}


# Mapa de nombres de sección a número (para forms y reports)
_SECTION_MAP: dict[str, int] = {
    "detail": 0,
    "header": 1, "formheader": 1, "reportheader": 1,
    "footer": 2, "formfooter": 2, "reportfooter": 2,
    "pageheader": 3,
    "pagefooter": 4,
    "grouplevel1header": 5, "group1header": 5,
    "grouplevel1footer": 6, "group1footer": 6,
    "grouplevel2header": 7, "group2header": 7,
    "grouplevel2footer": 8, "group2footer": 8,
}


def _resolve_section(section_val) -> int:
    """Acepta número (0) o nombre ('detail', 'header', 'reportheader', etc.)."""
    if isinstance(section_val, str):
        key = section_val.lower().replace(" ", "").replace("_", "")
        if key in _SECTION_MAP:
            return _SECTION_MAP[key]
        try:
            return int(key)
        except ValueError:
            valid = sorted(set(_SECTION_MAP.keys()))
            raise ValueError(
                f"Section '{section_val}' no reconocida. "
                f"Validas: {valid} o numero (0-8)"
            )
    return int(_coerce_prop(section_val))


def _resolve_ctrl_type(ctrl_type) -> int:
    """Acepta nombre ('CommandButton') o número (104)."""
    if isinstance(ctrl_type, int):
        return ctrl_type
    try:
        return int(ctrl_type)
    except (ValueError, TypeError):
        key = str(ctrl_type).lower()
        if key in _CTRL_TYPE_BY_NAME:
            return _CTRL_TYPE_BY_NAME[key]
        raise ValueError(
            f"Tipo de control desconocido: '{ctrl_type}'. "
            f"Usa un número o uno de: {list(_CTRL_TYPE.values())}"
        )


def _coerce_prop(value: Any) -> Any:
    """Convierte strings a int/bool cuando es apropiado para propiedades COM."""
    if isinstance(value, (int, float, bool)):
        return value
    if isinstance(value, str):
        low = value.lower()
        if low in ("true", "yes", "-1"):
            return True
        if low in ("false", "no", "0"):
            return False
        try:
            return int(value)
        except ValueError:
            pass
        try:
            return float(value)
        except ValueError:
            pass
    return value


def _open_in_design(app: Any, object_type: str, object_name: str) -> None:
    """Abre un form/report en vista diseño."""
    try:
        if object_type == "form":
            app.DoCmd.OpenForm(object_name, _AC_DESIGN)
        else:
            app.DoCmd.OpenReport(object_name, _AC_DESIGN)
    except Exception as exc:
        raise RuntimeError(
            f"No se pudo abrir '{object_name}' en vista diseño. "
            f"Si está abierto en vista normal, ciérralo primero.\nError: {exc}"
        )


def _save_and_close(app: Any, object_type: str, object_name: str) -> None:
    """Guarda y cierra un form/report abierto en diseño."""
    ac_type = _AC_FORM if object_type == "form" else _AC_REPORT
    try:
        app.DoCmd.Close(ac_type, object_name, _AC_SAVE_YES)
    except Exception as exc:
        log.warning("Error al cerrar '%s': %s", object_name, exc)


def _get_design_obj(app: Any, object_type: str, object_name: str) -> Any:
    """Devuelve el objeto Form o Report abierto en diseño."""
    return app.Forms(object_name) if object_type == "form" else app.Reports(object_name)


def ac_create_control(
    db_path: str, object_type: str, object_name: str,
    control_type: Any, props: dict
) -> dict:
    """
    Crea un control nuevo en un form/report abriéndolo en vista diseño.
    control_type: nombre ('CommandButton') o número (104).
    props: dict de propiedades. Las claves especiales que se pasan a CreateControl:
      section (default 0=Detail), parent (''), column_name (''),
      left, top, width, height (twips; -1 = automático).
    El resto se asignan como propiedades COM sobre el control creado.
    """
    if object_type not in ("form", "report"):
        raise ValueError("Solo 'form' o 'report'")

    app = _Session.connect(db_path)
    ctype = _resolve_ctrl_type(control_type)

    # Extraer parámetros posicionales/estructurales de props (no se asignan como prop)
    props = dict(props)  # copia para no mutar el original
    section     = _resolve_section(props.pop("section", 0))
    parent      = str(props.pop("parent",      ""))
    column_name = str(props.pop("column_name", ""))
    left        = int(_coerce_prop(props.pop("left",   -1)))
    top         = int(_coerce_prop(props.pop("top",    -1)))
    width       = int(_coerce_prop(props.pop("width",  -1)))
    height      = int(_coerce_prop(props.pop("height", -1)))

    _open_in_design(app, object_type, object_name)
    try:
        try:
            if object_type == "form":
                ctrl = app.CreateControl(
                    object_name, ctype, section, parent, column_name,
                    left, top, width, height,
                )
            else:
                ctrl = app.CreateReportControl(
                    object_name, ctype, section, parent, column_name,
                    left, top, width, height,
                )
        except Exception as exc:
            section_names = [k for k, v in _SECTION_MAP.items() if v == section]
            raise RuntimeError(
                f"Error creando control en section={section} "
                f"({', '.join(section_names) or 'desconocida'}): {exc}. "
                f"Verifique que la seccion existe en el {object_type}. "
                f"Secciones validas: 0=Detail, 1=Header, 2=Footer, "
                f"3=PageHeader, 4=PageFooter"
            )

        errors: dict[str, str] = {}
        for key, val in props.items():
            try:
                setattr(ctrl, key, _coerce_prop(val))
            except Exception as exc:
                errors[key] = str(exc)

        result: dict = {
            "name":         ctrl.Name,
            "control_type": ctype,
            "type_name":    _CTRL_TYPE.get(ctype, f"Type{ctype}"),
        }
        if errors:
            result["property_errors"] = errors
    finally:
        _save_and_close(app, object_type, object_name)
        # Invalidar caches — el form cambió en diseño
        cache_key = f"{object_type}:{object_name}"
        _parsed_controls_cache.pop(cache_key, None)
        _Session._cm_cache.pop(cache_key, None)
        _vbe_code_cache.pop(cache_key, None)

    return result


def ac_delete_control(
    db_path: str, object_type: str, object_name: str, control_name: str
) -> str:
    """Elimina un control de un form/report abriéndolo en vista diseño."""
    if object_type not in ("form", "report"):
        raise ValueError("Solo 'form' o 'report'")

    app = _Session.connect(db_path)
    _open_in_design(app, object_type, object_name)
    try:
        if object_type == "form":
            app.DeleteControl(object_name, control_name)
        else:
            app.DeleteReportControl(object_name, control_name)
    finally:
        _save_and_close(app, object_type, object_name)
        # Invalidar caches — el form cambió en diseño
        cache_key = f"{object_type}:{object_name}"
        _parsed_controls_cache.pop(cache_key, None)
        _Session._cm_cache.pop(cache_key, None)
        _vbe_code_cache.pop(cache_key, None)

    return f"OK: control '{control_name}' eliminado de '{object_name}'"


def ac_set_control_props(
    db_path: str, object_type: str, object_name: str,
    control_name: str, props: dict
) -> dict:
    """
    Modifica propiedades de un control existente abriendo el form/report en diseño.
    props: dict {propiedad: valor}. Los valores se convierten automáticamente
    a int/bool cuando corresponde.
    Devuelve {"applied": [...], "errors": {...}}.
    """
    if object_type not in ("form", "report"):
        raise ValueError("Solo 'form' o 'report'")

    app = _Session.connect(db_path)
    _open_in_design(app, object_type, object_name)
    applied: list[str] = []
    errors: dict[str, str] = {}
    try:
        obj  = _get_design_obj(app, object_type, object_name)
        ctrl = obj.Controls(control_name)
        for key, val in props.items():
            try:
                setattr(ctrl, key, _coerce_prop(val))
                applied.append(key)
            except Exception as exc:
                errors[key] = str(exc)
    finally:
        _save_and_close(app, object_type, object_name)
        # Invalidar caches — el form cambió en diseño
        cache_key = f"{object_type}:{object_name}"
        _parsed_controls_cache.pop(cache_key, None)
        _Session._cm_cache.pop(cache_key, None)
        _vbe_code_cache.pop(cache_key, None)

    return {"applied": applied, "errors": errors}


def ac_set_form_property(
    db_path: str, object_type: str, object_name: str, props: dict
) -> dict:
    """
    Establece propiedades a nivel de formulario/informe abriendo en vista diseño.
    Útil para cambiar RecordSource, Caption, DefaultView, HasModule, etc.
    props: dict {propiedad: valor}. Los valores se convierten a int/bool automáticamente.
    Devuelve {"applied": [...], "errors": {...}}.
    """
    if object_type not in ("form", "report"):
        raise ValueError("Solo 'form' o 'report'")

    app = _Session.connect(db_path)
    _open_in_design(app, object_type, object_name)
    applied: list[str] = []
    errors: dict[str, str] = {}
    try:
        obj = _get_design_obj(app, object_type, object_name)
        for key, val in props.items():
            try:
                setattr(obj, key, _coerce_prop(val))
                applied.append(key)
            except Exception as exc:
                errors[key] = str(exc)
    finally:
        _save_and_close(app, object_type, object_name)
        # Invalidar caches — las propiedades del form cambiaron
        cache_key = f"{object_type}:{object_name}"
        _parsed_controls_cache.pop(cache_key, None)
        _Session._cm_cache.pop(cache_key, None)
        _vbe_code_cache.pop(cache_key, None)

    return {"applied": applied, "errors": errors}


# ---------------------------------------------------------------------------
# Logica de negocio
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Create database
# ---------------------------------------------------------------------------

def ac_create_database(db_path: str) -> dict:
    """Crea una BD Access (.accdb) vacía. Error si ya existe."""
    resolved = str(Path(db_path).resolve())
    if os.path.exists(resolved):
        raise FileExistsError(
            f"Ya existe '{resolved}'. Usa access_execute_sql para modificarla."
        )
    # Ensure Access is running
    if _Session._app is None:
        _Session._launch()
    app = _Session._app
    # Close any previously open DB
    if _Session._db_open is not None:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        _Session._db_open = None
    try:
        app.NewCurrentDatabase(resolved)
    except Exception as exc:
        raise RuntimeError(f"Error al crear BD: {exc}")
    # FIX: Close and reopen to ensure CurrentDb() works reliably
    try:
        app.CloseCurrentDatabase()
        app.OpenCurrentDatabase(resolved)
    except Exception:
        pass  # If reopen fails, at least the file was created
    _Session._db_open = resolved
    _Session._cm_cache.clear()
    _vbe_code_cache.clear()
    _parsed_controls_cache.clear()
    size = os.path.getsize(resolved) if os.path.exists(resolved) else 0
    return {"db_path": resolved, "status": "created", "size_bytes": size}


# ---------------------------------------------------------------------------
# Create table via DAO
# ---------------------------------------------------------------------------
# Mapa de nombres de tipo → constantes DAO dbType
_FIELD_TYPE_MAP: dict[str, int] = {
    "autonumber": 4, "autoincrement": 4,  # dbLong + dbAutoIncrField attribute
    "long": 4, "integer": 3, "short": 3, "byte": 2,
    "text": 10, "memo": 12, "currency": 5,
    "double": 7, "single": 6, "float": 7,
    "datetime": 8, "date": 8,
    "boolean": 1, "yesno": 1, "bit": 1,
    "guid": 15, "ole": 11, "bigint": 16,
}
_DB_AUTO_INCR_FIELD = 16  # dbAutoIncrField attribute flag
_DB_ATTACH_SAVE_PWD = 65536  # dbAttachSavePWD — save password in linked table connect string


def _set_field_prop(db: Any, table_name: str, field_name: str,
                    prop_name: str, value: Any) -> None:
    """Helper interno para establecer propiedad de campo con fallback a CreateProperty."""
    fld = db.TableDefs(table_name).Fields(field_name)
    try:
        fld.Properties(prop_name).Value = value
    except Exception:
        prop = fld.CreateProperty(prop_name, 10, value)  # 10 = dbText
        fld.Properties.Append(prop)


def ac_create_table(db_path: str, table_name: str, fields: list[dict]) -> dict:
    """
    Crea una tabla Access via DAO con soporte completo de tipos, defaults,
    descripciones y propiedades — todo en una sola llamada.
    """
    app = _Session.connect(db_path)
    db = app.CurrentDb()

    # Verificar que no existe
    existing = [db.TableDefs(i).Name for i in range(db.TableDefs.Count)]
    if table_name in existing:
        raise ValueError(f"La tabla '{table_name}' ya existe.")

    td = db.CreateTableDef(table_name)
    pk_fields: list[str] = []
    created_fields: list[dict] = []

    for fdef in fields:
        name = fdef["name"]
        ftype = fdef.get("type", "text").lower()
        size = fdef.get("size", 0)
        required = fdef.get("required", False)
        pk = fdef.get("primary_key", False)

        dao_type = _FIELD_TYPE_MAP.get(ftype)
        if dao_type is None:
            raise ValueError(
                f"Tipo desconocido: '{ftype}'. "
                f"Validos: {sorted(set(_FIELD_TYPE_MAP.keys()))}"
            )

        is_auto = ftype in ("autonumber", "autoincrement")

        # Text needs size
        if dao_type == 10 and size == 0:
            size = 255

        if size > 0:
            fld = td.CreateField(name, dao_type, size)
        else:
            fld = td.CreateField(name, dao_type)

        if is_auto:
            fld.Attributes = fld.Attributes | _DB_AUTO_INCR_FIELD

        fld.Required = required or pk

        td.Fields.Append(fld)

        if pk:
            pk_fields.append(name)

        created_fields.append({
            "name": name,
            "type": ftype,
            "size": size if size > 0 else None,
        })

    # Create primary key index
    if pk_fields:
        idx = td.CreateIndex("PrimaryKey")
        idx.Primary = True
        idx.Unique = True
        for pk_name in pk_fields:
            idx_fld = idx.CreateField(pk_name)
            idx.Fields.Append(idx_fld)
        td.Indexes.Append(idx)

    db.TableDefs.Append(td)
    db.TableDefs.Refresh()

    # Set defaults and descriptions via field properties (post-creation)
    for fdef in fields:
        name = fdef["name"]
        default = fdef.get("default")
        description = fdef.get("description")
        if default is not None:
            try:
                _set_field_prop(db, table_name, name, "DefaultValue", str(default))
            except Exception as e:
                log.warning("Error setting default for %s.%s: %s", table_name, name, e)
        if description is not None:
            try:
                _set_field_prop(db, table_name, name, "Description", description)
            except Exception as e:
                log.warning("Error setting description for %s.%s: %s", table_name, name, e)

    return {
        "table_name": table_name,
        "fields": created_fields,
        "primary_key": pk_fields,
        "status": "created",
    }


# ---------------------------------------------------------------------------
# Alter table via DAO
# ---------------------------------------------------------------------------
def ac_alter_table(
    db_path: str, table_name: str, action: str,
    field_name: str, new_name: str | None = None,
    field_type: str = "text", size: int = 0,
    required: bool = False, default: Any = None,
    description: str | None = None, confirm: bool = False,
) -> dict:
    """
    Modifica la estructura de una tabla Access via DAO.
    Acciones: add_field, delete_field (requiere confirm=true), rename_field.
    """
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    td = db.TableDefs(table_name)

    if action == "add_field":
        ftype = field_type.lower()
        dao_type = _FIELD_TYPE_MAP.get(ftype)
        if dao_type is None:
            raise ValueError(
                f"Tipo desconocido: '{ftype}'. "
                f"Validos: {sorted(set(_FIELD_TYPE_MAP.keys()))}"
            )
        is_auto = ftype in ("autonumber", "autoincrement")
        if dao_type == 10 and size == 0:
            size = 255
        if size > 0:
            fld = td.CreateField(field_name, dao_type, size)
        else:
            fld = td.CreateField(field_name, dao_type)
        if is_auto:
            fld.Attributes = fld.Attributes | _DB_AUTO_INCR_FIELD
        fld.Required = required
        td.Fields.Append(fld)
        td.Fields.Refresh()
        if default is not None:
            try:
                _set_field_prop(db, table_name, field_name, "DefaultValue", str(default))
            except Exception as e:
                log.warning("Error setting default for %s.%s: %s", table_name, field_name, e)
        if description is not None:
            try:
                _set_field_prop(db, table_name, field_name, "Description", description)
            except Exception as e:
                log.warning("Error setting description for %s.%s: %s", table_name, field_name, e)
        return {"action": "field_added", "table": table_name, "field": field_name, "type": ftype}

    elif action == "delete_field":
        if not confirm:
            return {
                "error": (
                    f"Eliminar campo '{field_name}' de '{table_name}' es destructivo. "
                    "Usa confirm=true para confirmar."
                )
            }
        td.Fields.Delete(field_name)
        return {"action": "field_deleted", "table": table_name, "field": field_name}

    elif action == "rename_field":
        if not new_name:
            raise ValueError("rename_field requiere new_name")
        fld = td.Fields(field_name)
        fld.Name = new_name
        return {"action": "field_renamed", "table": table_name,
                "old_name": field_name, "new_name": new_name}

    else:
        raise ValueError(
            f"Accion desconocida: '{action}'. "
            "Validas: add_field, delete_field, rename_field"
        )


# ---------------------------------------------------------------------------
# List objects
# ---------------------------------------------------------------------------

def ac_list_objects(db_path: str, object_type: str = "all") -> dict:
    """Devuelve un dict {tipo: [nombres...]} con los objetos de la BD."""
    app = _Session.connect(db_path)

    # CurrentData  → objetos de datos (tablas, queries)
    # CurrentProject → objetos de codigo (forms, reports, modulos, macros)
    containers = {
        "table":  app.CurrentData.AllTables,
        "query":  app.CurrentData.AllQueries,
        "form":   app.CurrentProject.AllForms,
        "report": app.CurrentProject.AllReports,
        "macro":  app.CurrentProject.AllMacros,
        "module": app.CurrentProject.AllModules,
    }

    keys = list(containers) if object_type == "all" else [object_type]
    result: dict[str, list] = {}

    for k in keys:
        if k not in containers:
            continue
        col = containers[k]
        names = [col.Item(i).Name for i in range(col.Count)]
        if k == "table":
            # Filter out system and temp tables
            names = [n for n in names if not n.startswith("MSys") and not n.startswith("~")]
        result[k] = names

    return result


# ---------------------------------------------------------------------------
# Delete object
# ---------------------------------------------------------------------------

def ac_delete_object(
    db_path: str, object_type: str, object_name: str, confirm: bool = False,
) -> dict:
    """Elimina un objeto Access (module, form, report, query, macro) via DoCmd.DeleteObject."""
    if object_type not in AC_TYPE:
        raise ValueError(
            f"object_type '{object_type}' invalido. Validos: {list(AC_TYPE)}"
        )
    if not confirm:
        raise ValueError(
            "Operacion destructiva: se requiere confirm=true para eliminar un objeto."
        )
    app = _Session.connect(db_path)
    try:
        app.DoCmd.DeleteObject(AC_TYPE[object_type], object_name)
    except Exception as exc:
        raise RuntimeError(
            f"Error al eliminar {object_type} '{object_name}': {exc}"
        )
    finally:
        _vbe_code_cache.clear()
        _parsed_controls_cache.clear()
        _Session._cm_cache.clear()
    return {
        "action": "deleted",
        "object_type": object_type,
        "object_name": object_name,
    }


def ac_get_code(db_path: str, object_type: str, name: str) -> str:
    """
    Exporta un objeto Access a texto y devuelve el contenido.
    Para formularios e informes elimina las secciones binarias (PrtMip, PrtDevMode…)
    que no tienen relevancia para editar VBA/controles y representan el 95 % del tamaño.
    ac_set_code las restaura automáticamente antes de importar.
    """
    if object_type not in AC_TYPE:
        raise ValueError(
            f"object_type '{object_type}' invalido. Validos: {list(AC_TYPE)}"
        )
    app = _Session.connect(db_path)

    fd, tmp = tempfile.mkstemp(suffix=".txt", prefix="access_mcp_")
    os.close(fd)
    try:
        app.SaveAsText(AC_TYPE[object_type], name, tmp)
        text, _enc = _read_tmp(tmp)
        if object_type in ("form", "report"):
            text = _strip_binary_sections(text)
        return text
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


def _split_code_behind(code: str) -> tuple[str, str]:
    """
    Separa un texto de form/report en (form_text, vba_code).
    Si el código contiene 'CodeBehindForm' o 'CodeBehindReport', lo separa.
    Devuelve (form_text_sin_vba, vba_code) donde vba_code puede estar vacío.
    El form_text se limpia de HasModule si hay VBA (se inyectará después).
    """
    # Buscar la línea que marca el inicio del código VBA
    for marker in ("CodeBehindForm", "CodeBehindReport"):
        idx = code.find(marker)
        if idx != -1:
            form_part = code[:idx].rstrip() + "\n"
            vba_part = code[idx:].split("\n", 1)
            vba_code = vba_part[1] if len(vba_part) > 1 else ""
            # Quitar líneas Attribute VB_ del VBA (se generan automáticamente)
            vba_lines = []
            for line in vba_code.splitlines():
                stripped = line.strip()
                if stripped.startswith("Attribute VB_"):
                    continue
                vba_lines.append(line)
            vba_code = "\n".join(vba_lines).strip()
            return form_part, vba_code
    return code, ""


def _inject_vba_after_import(app: Any, object_type: str, name: str, vba_code: str) -> None:
    """
    Inyecta código VBA en un form/report después de importarlo.
    Activa HasModule abriendo en diseño, luego usa VBE para insertar el código.
    """
    if not vba_code.strip():
        return

    # 1. Abrir en diseño y activar HasModule
    _open_in_design(app, object_type, name)
    try:
        obj = _get_design_obj(app, object_type, name)
        obj.HasModule = True
    finally:
        _save_and_close(app, object_type, name)

    # 2. Limpiar cache de VBE (el módulo acaba de crearse)
    cache_key = f"{object_type}:{name}"
    _Session._cm_cache.pop(cache_key, None)
    _vbe_code_cache.pop(cache_key, None)

    # 3. Inyectar código via VBE
    cm = _get_code_module(app, object_type, name)
    total = cm.CountOfLines

    # Borrar contenido auto-generado por Access (Option Compare Database, etc.)
    # para evitar duplicados con el VBA que vamos a inyectar
    if total > 0:
        cm.DeleteLines(1, total)

    # Normalizar line endings a \r\n (VBE lo requiere)
    vba_code = vba_code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    if not vba_code.endswith("\r\n"):
        vba_code += "\r\n"

    cm.InsertLines(1, vba_code)

    # Invalidar caches
    _vbe_code_cache.pop(cache_key, None)
    _Session._cm_cache.pop(cache_key, None)


def ac_set_code(db_path: str, object_type: str, name: str, code: str) -> str:
    """
    Importa texto como definicion de un objeto Access (crea o sobreescribe).
    Para formularios e informes re-inyecta automáticamente las secciones binarias
    (PrtMip, PrtDevMode…) desde el export actual, de modo que el caller no necesita
    incluirlas en el código que envía.

    Si el código contiene una sección CodeBehindForm/CodeBehindReport, se separa
    automáticamente: primero se importa el form/report sin VBA, luego se inyecta
    el código VBA via VBE (evitando problemas de encoding con LoadFromText).
    """
    if object_type not in AC_TYPE:
        raise ValueError(
            f"object_type '{object_type}' invalido. Validos: {list(AC_TYPE)}"
        )
    app = _Session.connect(db_path)

    # Separar CodeBehindForm/CodeBehindReport si existe
    vba_code = ""
    if object_type in ("form", "report"):
        code, vba_code = _split_code_behind(code)
        # Quitar HasModule del form text — se activará al inyectar VBA
        if vba_code:
            code = re.sub(r"^\s*HasModule\s*=.*$", "", code, flags=re.MULTILINE)

    # Si el código no contiene secciones binarias (fue devuelto por ac_get_code
    # con el filtrado activo), las restauramos desde el form/report actual.
    if object_type in ("form", "report") and not any(
        s in code for s in _BINARY_SECTIONS
    ):
        log.info("ac_set_code: restaurando secciones binarias para '%s'", name)
        code = _restore_binary_sections(app, object_type, name, code)

    # Backup del objeto existente por si falla el import
    backup_tmp = None
    if object_type in ("form", "report"):
        try:
            fd_bk, backup_tmp = tempfile.mkstemp(suffix=".txt", prefix="access_mcp_bk_")
            os.close(fd_bk)
            app.SaveAsText(AC_TYPE[object_type], name, backup_tmp)
        except Exception:
            # No existe aún — no hay backup que hacer
            if backup_tmp:
                try:
                    os.unlink(backup_tmp)
                except OSError:
                    pass
            backup_tmp = None

    fd, tmp = tempfile.mkstemp(suffix=".txt", prefix="access_mcp_")
    os.close(fd)
    try:
        # Módulos VBA (.bas) esperan ANSI/cp1252; forms/reports/queries/macros esperan UTF-16LE con BOM
        enc = "cp1252" if object_type == "module" else "utf-16"
        _write_tmp(tmp, code, encoding=enc)
        try:
            app.LoadFromText(AC_TYPE[object_type], name, tmp)
        except Exception as import_exc:
            # Restaurar backup si existe
            if backup_tmp and os.path.exists(backup_tmp):
                log.warning("ac_set_code: import falló, restaurando backup de '%s'", name)
                try:
                    app.LoadFromText(AC_TYPE[object_type], name, backup_tmp)
                except Exception:
                    log.error("ac_set_code: no se pudo restaurar backup de '%s'", name)
            raise import_exc

        # Invalidar caches para este objeto (el código y los controles cambiaron)
        cache_key = f"{object_type}:{name}"
        _vbe_code_cache.pop(cache_key, None)
        _Session._cm_cache.pop(cache_key, None)
        _parsed_controls_cache.pop(cache_key, None)

        # Inyectar VBA si había CodeBehindForm
        vba_msg = ""
        if vba_code:
            _inject_vba_after_import(app, object_type, name, vba_code)
            vba_msg = " (con VBA inyectado via VBE)"

        return f"OK: '{name}' ({object_type}) importado correctamente en {db_path}{vba_msg}"
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        if backup_tmp:
            try:
                os.unlink(backup_tmp)
            except OSError:
                pass


_DESTRUCTIVE_PREFIXES = ("DELETE", "DROP", "TRUNCATE", "ALTER")


def _serialize_value(val: Any) -> Any:
    """Convierte tipos COM no serializables a JSON-safe."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.isoformat()
    try:
        from decimal import Decimal
        if isinstance(val, Decimal):
            return float(val)
    except ImportError:
        pass
    if isinstance(val, bytes):
        return f"<binary {len(val)} bytes>"
    return val


def ac_execute_sql(
    db_path: str, sql: str, limit: int = 500,
    confirm_destructive: bool = False,
) -> dict:
    """
    Ejecuta SQL en la BD via DAO.
    SELECT  → devuelve {rows: [...], count: N, truncated?: bool}
    Otros   → devuelve {affected_rows: N}
    DELETE/DROP/TRUNCATE/ALTER requieren confirm_destructive=True.
    """
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    normalized = sql.strip().upper()

    if normalized.startswith("SELECT"):
        limit = max(1, min(limit, 10000))
        rs = db.OpenRecordset(sql)
        fields = [rs.Fields(i).Name for i in range(rs.Fields.Count)]
        rows: list[dict] = []
        if not rs.EOF:
            rs.MoveFirst()
            while not rs.EOF and len(rows) < limit:
                rows.append(
                    {fields[i]: _serialize_value(rs.Fields(i).Value)
                     for i in range(len(fields))}
                )
                rs.MoveNext()
        truncated = not rs.EOF
        rs.Close()
        result: dict = {"rows": rows, "count": len(rows)}
        if truncated:
            result["truncated"] = True
        return result
    else:
        if any(normalized.startswith(p) for p in _DESTRUCTIVE_PREFIXES):
            if not confirm_destructive:
                return {
                    "error": (
                        "SQL destructivo detectado. "
                        "Usa confirm_destructive=true para ejecutar: "
                        + sql[:100]
                    )
                }
        db.Execute(sql)
        return {"affected_rows": db.RecordsAffected}


# Mapa DAO Type → nombre legible
_DAO_FIELD_TYPE: dict[int, str] = {
    1: "Boolean", 2: "Byte", 3: "Integer", 4: "Long", 5: "Currency",
    6: "Single", 7: "Double", 8: "Date/Time", 10: "Text",
    11: "OLE Object", 12: "Memo", 15: "GUID", 16: "BigInt",
    20: "Decimal",
}

# Mapa DAO Relation Attributes → nombre legible
_REL_ATTR: dict[int, str] = {
    1: "Unique", 2: "DontEnforce", 256: "UpdateCascade", 4096: "DeleteCascade",
}

# Access output / transfer constants
_AC_OUTPUT_REPORT = 3      # acOutputReport
_AC_IMPORT = 0             # acImport
_AC_EXPORT = 1             # acExport
_AC_EXPORT_DELIM = 2       # acExportDelim (CSV export)
_AC_SPREADSHEET_XLSX = 10  # acSpreadsheetTypeExcel12Xml
_AC_CMD_COMPILE = 126      # acCmdCompileAndSaveAllModules

# DAO QueryDef type constants
_QUERYDEF_TYPE: dict[int, str] = {
    0: "Select", 16: "Crosstab", 32: "Delete", 48: "Update",
    64: "Append", 80: "MakeTable", 96: "DDL", 112: "SQLPassThrough",
    128: "Union", 240: "Action",
}

# Common startup properties
_STARTUP_PROPS = [
    "AppTitle", "AppIcon", "StartupForm", "StartupShowDBWindow",
    "StartupShowStatusBar", "StartupShortcutMenuBar",
    "AllowShortcutMenus", "AllowFullMenus", "AllowBuiltInToolbars",
    "AllowToolbarChanges", "AllowBreakIntoCode", "AllowSpecialKeys",
    "AllowBypassKey", "AllowDatasheetSchema",
]


def ac_table_info(db_path: str, table_name: str) -> dict:
    """
    Devuelve la estructura de una tabla Access local o linkada:
    campos con nombre, tipo, tamaño, required; record_count; is_linked.
    Usa DAO TableDef.Fields.
    """
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        td = db.TableDefs(table_name)
    except Exception as exc:
        raise ValueError(f"Tabla '{table_name}' no encontrada: {exc}")

    is_linked = bool(td.Connect)
    fields: list[dict] = []
    for i in range(td.Fields.Count):
        fld = td.Fields(i)
        ftype = fld.Type
        # AutoNumber detection: Long (4) + dbAutoIncrField attribute (16)
        type_name = _DAO_FIELD_TYPE.get(ftype, f"Type{ftype}")
        if ftype == 4 and (fld.Attributes & 16):
            type_name = "AutoNumber"
        fields.append({
            "name": fld.Name,
            "type": type_name,
            "size": fld.Size,
            "required": bool(fld.Required),
        })

    # Record count (puede fallar en linked tables)
    try:
        record_count = td.RecordCount
        if record_count == -1:
            # Para linked tables, abrir recordset para contar
            rs = db.OpenRecordset(f"SELECT COUNT(*) AS cnt FROM [{table_name}]")
            record_count = rs.Fields(0).Value
            rs.Close()
    except Exception:
        record_count = -1

    return {
        "table_name": table_name,
        "fields": fields,
        "record_count": record_count,
        "is_linked": is_linked,
        "source_table": td.SourceTableName if is_linked else "",
        "connect": td.Connect if is_linked else "",
    }


def ac_export_structure(db_path: str, output_path: Optional[str] = None) -> str:
    """
    Genera un fichero Markdown con la estructura completa de la BD:
    modulos VBA con sus firmas de funciones, formularios, informes y queries.
    """
    if output_path is None:
        output_path = str(Path(db_path).parent / "db_structure.md")

    objects = ac_list_objects(db_path, "all")
    modules  = objects.get("module",  [])
    forms    = objects.get("form",    [])
    reports  = objects.get("report",  [])
    queries  = objects.get("query",   [])
    macros   = objects.get("macro",   [])

    lines: list[str] = []
    lines.append(f"# Estructura de `{Path(db_path).name}`")
    lines.append(f"\n**Ruta**: `{db_path}`  ")
    lines.append(f"**Generado**: {datetime.now().strftime('%Y-%m-%d %H:%M')}  ")
    lines.append(
        f"**Resumen**: {len(modules)} módulos · {len(forms)} formularios · "
        f"{len(reports)} informes · {len(queries)} queries · {len(macros)} macros\n"
    )

    # ── Módulos VBA con firmas ───────────────────────────────────────────────
    # Leer módulos vía VBE (sin SaveAsText/disco) y calentando el cache de código
    app = _Session.connect(db_path)
    lines.append(f"## Módulos VBA ({len(modules)})\n")
    for mod_name in modules:
        lines.append(f"### `{mod_name}`")
        try:
            cm = _get_code_module(app, "module", mod_name)
            cache_key = f"module:{mod_name}"
            code = _cm_all_code(cm, cache_key)
            sigs = []
            for line in code.splitlines():
                stripped = line.strip()
                if re.match(
                    r"^(Public\s+|Private\s+|Friend\s+)?(Function|Sub)\s+\w+",
                    stripped,
                    re.IGNORECASE,
                ):
                    sigs.append(f"  - `{stripped}`")
            if sigs:
                lines.extend(sigs)
            else:
                lines.append("  *(sin funciones/subs públicos)*")
        except Exception as exc:
            lines.append(f"  *(error al leer: {exc})*")
        lines.append("")

    # ── Formularios ──────────────────────────────────────────────────────────
    lines.append(f"## Formularios ({len(forms)})\n")
    if forms:
        for name in forms:
            lines.append(f"- `{name}`")
    else:
        lines.append("*(ninguno)*")
    lines.append("")

    # ── Informes ─────────────────────────────────────────────────────────────
    lines.append(f"## Informes ({len(reports)})\n")
    if reports:
        for name in reports:
            lines.append(f"- `{name}`")
    else:
        lines.append("*(ninguno)*")
    lines.append("")

    # ── Queries ──────────────────────────────────────────────────────────────
    lines.append(f"## Queries ({len(queries)})\n")
    if queries:
        for name in queries:
            lines.append(f"- `{name}`")
    else:
        lines.append("*(ninguno)*")
    lines.append("")

    # ── Macros ───────────────────────────────────────────────────────────────
    if macros:
        lines.append(f"## Macros ({len(macros)})\n")
        for name in macros:
            lines.append(f"- `{name}`")
        lines.append("")

    content = "\n".join(lines)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)

    return f"[Guardado en `{output_path}`]\n\n{content}"


# ---------------------------------------------------------------------------
# Database properties
# ---------------------------------------------------------------------------

def ac_get_db_property(db_path: str, name: str) -> dict:
    """Lee una propiedad de la BD o una opcion de la aplicacion Access."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        val = db.Properties(name).Value
        return {"name": name, "value": val, "source": "database"}
    except Exception:
        pass
    try:
        val = app.GetOption(name)
        return {"name": name, "value": val, "source": "application"}
    except Exception as exc:
        raise ValueError(
            f"Propiedad '{name}' no encontrada en CurrentDb().Properties "
            f"ni en Application.GetOption. Error: {exc}"
        )


def ac_set_db_property(
    db_path: str, name: str, value: Any,
    prop_type: Optional[int] = None,
) -> dict:
    """Establece una propiedad de la BD o una opcion de la aplicacion Access."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    coerced = _coerce_prop(value)

    # Try DB-level property
    try:
        db.Properties(name).Value = coerced
        return {"name": name, "value": coerced, "source": "database", "action": "updated"}
    except Exception:
        pass

    # Try Application option
    try:
        app.SetOption(name, coerced)
        return {"name": name, "value": coerced, "source": "application", "action": "updated"}
    except Exception:
        pass

    # Property doesn't exist — create it
    if prop_type is None:
        if isinstance(coerced, bool):
            prop_type = 1   # dbBoolean
        elif isinstance(coerced, int):
            prop_type = 4   # dbLong
        else:
            prop_type = 10  # dbText
    try:
        prop = db.CreateProperty(name, prop_type, coerced)
        db.Properties.Append(prop)
        return {"name": name, "value": coerced, "source": "database", "action": "created"}
    except Exception as exc:
        raise RuntimeError(
            f"No se pudo crear propiedad '{name}'. "
            f"prop_type: 1=Boolean, 4=Long, 10=Text. Error: {exc}"
        )


# ---------------------------------------------------------------------------
# Linked tables
# ---------------------------------------------------------------------------

def ac_list_linked_tables(db_path: str) -> dict:
    """Lista todas las tablas vinculadas con informacion de conexion."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    linked: list[dict] = []
    for i in range(db.TableDefs.Count):
        td = db.TableDefs(i)
        conn = td.Connect
        if not conn:
            continue
        name = td.Name
        if name.startswith("~") or name.startswith("MSys"):
            continue
        linked.append({
            "name": name,
            "source_table": td.SourceTableName,
            "connect_string": conn,
            "is_odbc": conn.upper().startswith("ODBC;"),
        })
    return {"count": len(linked), "linked_tables": linked}


def ac_relink_table(
    db_path: str, table_name: str, new_connect: str,
    relink_all: bool = False,
) -> dict:
    """Cambia la cadena de conexion de una tabla vinculada y refresca."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    relinked: list[dict] = []

    try:
        ref_td = db.TableDefs(table_name)
    except Exception as exc:
        raise ValueError(f"Tabla '{table_name}' no encontrada: {exc}")
    if not ref_td.Connect:
        raise ValueError(f"'{table_name}' no es una tabla vinculada")

    # Auto-detect if connect string has UID/PWD → set dbAttachSavePWD
    _needs_save_pwd = ("UID=" in new_connect.upper() or "PWD=" in new_connect.upper())

    def _relink_one(td_name: str, old_conn: str):
        """Relink a single table. If dbAttachSavePWD needed, delete+recreate."""
        if _needs_save_pwd:
            # Cannot modify Attributes on appended TableDef — must delete+recreate
            src_table = db.TableDefs(td_name).SourceTableName
            db.TableDefs.Delete(td_name)
            db.TableDefs.Refresh()
            new_td = db.CreateTableDef(td_name)
            new_td.Connect = new_connect
            new_td.SourceTableName = src_table
            new_td.Attributes = _DB_ATTACH_SAVE_PWD
            db.TableDefs.Append(new_td)
            db.TableDefs.Refresh()
        else:
            td = db.TableDefs(td_name)
            td.Connect = new_connect
            td.RefreshLink()
        relinked.append({"name": td_name, "old_connect": old_conn, "new_connect": new_connect})

    if relink_all:
        old_connect = ref_td.Connect
        names_to_relink = []
        for i in range(db.TableDefs.Count):
            td = db.TableDefs(i)
            if td.Connect == old_connect:
                names_to_relink.append((td.Name, td.Connect))
        for name, old in names_to_relink:
            _relink_one(name, old)
    else:
        old = ref_td.Connect
        _relink_one(table_name, old)

    return {"relinked_count": len(relinked), "tables": relinked}


# ---------------------------------------------------------------------------
# Relationships
# ---------------------------------------------------------------------------

def ac_list_relationships(db_path: str) -> dict:
    """Lista todas las relaciones entre tablas de la BD."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    rels: list[dict] = []
    for i in range(db.Relations.Count):
        rel = db.Relations(i)
        name = rel.Name
        if name.startswith("MSys"):
            continue
        fields: list[dict] = []
        for j in range(rel.Fields.Count):
            fld = rel.Fields(j)
            fields.append({"local": fld.Name, "foreign": fld.ForeignName})
        attrs = rel.Attributes
        attr_flags = [label for bit, label in _REL_ATTR.items() if attrs & bit]
        rels.append({
            "name": name,
            "table": rel.Table,
            "foreign_table": rel.ForeignTable,
            "fields": fields,
            "attributes": attrs,
            "attribute_flags": attr_flags,
        })
    return {"count": len(rels), "relationships": rels}


def ac_create_relationship(
    db_path: str, name: str, table: str, foreign_table: str,
    fields: list[dict], attributes: int = 0,
) -> dict:
    """Crea una relacion entre dos tablas via DAO."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    rel = db.CreateRelation(name, table, foreign_table, attributes)
    for fmap in fields:
        local_name = fmap.get("local")
        foreign_name = fmap.get("foreign")
        if not local_name or not foreign_name:
            raise ValueError(f"Cada campo debe tener 'local' y 'foreign'. Recibido: {fmap}")
        fld = rel.CreateField(local_name)
        fld.ForeignName = foreign_name
        rel.Fields.Append(fld)
    try:
        db.Relations.Append(rel)
    except Exception as exc:
        raise RuntimeError(
            f"Error al crear relacion '{name}' entre '{table}' y '{foreign_table}': {exc}"
        )
    attr_flags = [label for bit, label in _REL_ATTR.items() if attributes & bit]
    return {
        "name": name, "table": table, "foreign_table": foreign_table,
        "fields": fields, "attributes": attributes,
        "attribute_flags": attr_flags, "status": "created",
    }


def ac_delete_relationship(db_path: str, name: str) -> dict:
    """Elimina una relacion entre tablas por nombre."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        db.Relations.Delete(name)
    except Exception as exc:
        raise RuntimeError(f"Error al eliminar relacion '{name}': {exc}")
    return {"action": "deleted", "name": name}


# ---------------------------------------------------------------------------
# VBA References
# ---------------------------------------------------------------------------

def ac_list_references(db_path: str) -> dict:
    """Lista todas las referencias VBA del proyecto."""
    app = _Session.connect(db_path)
    try:
        refs_col = app.VBE.ActiveVBProject.References
    except Exception as exc:
        raise RuntimeError(f"No se pudo acceder al VBE. Error: {exc}")
    refs: list[dict] = []
    for i in range(1, refs_col.Count + 1):  # VBA collections are 1-based
        ref = refs_col(i)
        try:
            is_broken = bool(ref.IsBroken)
        except Exception:
            is_broken = True
        try:
            built_in = bool(ref.BuiltIn)
        except Exception:
            built_in = False
        refs.append({
            "name": ref.Name,
            "description": ref.Description,
            "full_path": ref.FullPath,
            "guid": ref.GUID if ref.GUID else "",
            "major": ref.Major,
            "minor": ref.Minor,
            "is_broken": is_broken,
            "built_in": built_in,
        })
    return {"count": len(refs), "references": refs}


def ac_manage_reference(
    db_path: str, action: str,
    name: Optional[str] = None,
    path: Optional[str] = None,
    guid: Optional[str] = None,
    major: int = 0, minor: int = 0,
) -> dict:
    """Agrega o elimina una referencia VBA del proyecto."""
    app = _Session.connect(db_path)
    try:
        refs = app.VBE.ActiveVBProject.References
    except Exception as exc:
        raise RuntimeError(f"No se pudo acceder al VBE. Error: {exc}")

    if action == "add":
        if guid:
            try:
                ref = refs.AddFromGuid(guid, major, minor)
                result = {"action": "added", "name": ref.Name, "guid": guid, "major": major, "minor": minor}
            except Exception as exc:
                raise RuntimeError(f"Error al agregar referencia por GUID '{guid}': {exc}")
        elif path:
            try:
                ref = refs.AddFromFile(path)
                result = {"action": "added", "name": ref.Name, "full_path": path}
            except Exception as exc:
                raise RuntimeError(f"Error al agregar referencia desde '{path}': {exc}")
        else:
            raise ValueError("Para action='add' se requiere 'guid' o 'path'")
    elif action == "remove":
        if not name:
            raise ValueError("Para action='remove' se requiere 'name'")
        found = None
        for i in range(1, refs.Count + 1):
            ref = refs(i)
            if ref.Name.lower() == name.lower():
                found = ref
                break
        if found is None:
            raise ValueError(f"Referencia '{name}' no encontrada")
        try:
            if found.BuiltIn:
                raise ValueError(f"'{name}' es built-in y no se puede eliminar")
        except AttributeError:
            pass  # BuiltIn property not available in old Access versions
        try:
            refs.Remove(found)
            result = {"action": "removed", "name": name}
        except Exception as exc:
            raise RuntimeError(f"Error al eliminar referencia '{name}': {exc}")
    else:
        raise ValueError(f"action debe ser 'add' o 'remove', recibido: '{action}'")

    # References affect VBE compilation — clear code caches
    _vbe_code_cache.clear()
    _Session._cm_cache.clear()
    return result


# ---------------------------------------------------------------------------
# Compact & Repair
# ---------------------------------------------------------------------------

def ac_compact_repair(db_path: str) -> dict:
    """Compacta y repara la BD. Cierra, compacta a temp, reemplaza y reabre."""
    resolved = str(Path(db_path).resolve())
    app = _Session.connect(resolved)
    original_size = os.path.getsize(resolved)

    # Close current database (keep Access alive)
    try:
        app.CloseCurrentDatabase()
    except Exception as exc:
        raise RuntimeError(f"No se pudo cerrar la BD para compactar: {exc}")
    _Session._db_open = None
    _Session._cm_cache.clear()
    _vbe_code_cache.clear()
    _parsed_controls_cache.clear()

    # Temp/bak paths in same directory (atomic rename)
    db_dir = os.path.dirname(resolved)
    db_name, db_ext = os.path.splitext(os.path.basename(resolved))
    tmp_path = os.path.join(db_dir, f"{db_name}_compact_tmp{db_ext}")
    bak_path = os.path.join(db_dir, f"{db_name}_compact_bak{db_ext}")

    try:
        for p in (tmp_path, bak_path):
            if os.path.exists(p):
                os.unlink(p)

        try:
            app.CompactRepair(resolved, tmp_path)
        except Exception as exc:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)
            raise RuntimeError(f"Error en CompactRepair: {exc}")

        if not os.path.exists(tmp_path):
            raise RuntimeError("CompactRepair no genero el fichero de salida")
        compacted_size = os.path.getsize(tmp_path)

        # Atomic swap: original → .bak, tmp → original
        os.rename(resolved, bak_path)
        try:
            os.rename(tmp_path, resolved)
        except Exception:
            os.rename(bak_path, resolved)  # rollback
            raise

        try:
            os.unlink(bak_path)
        except OSError:
            pass

    except Exception:
        # Try to reopen whatever is at the original path
        try:
            if os.path.exists(resolved):
                app.OpenCurrentDatabase(resolved)
                _Session._db_open = resolved
        except Exception:
            pass
        raise

    # Reopen compacted database
    try:
        app.OpenCurrentDatabase(resolved)
        _Session._db_open = resolved
    except Exception as exc:
        raise RuntimeError(f"BD compactada OK pero error al reabrir: {exc}")

    saved = original_size - compacted_size
    return {
        "original_size": original_size,
        "compacted_size": compacted_size,
        "saved_bytes": saved,
        "saved_pct": round(saved / original_size * 100, 1) if original_size > 0 else 0,
        "status": "compacted",
    }


# ---------------------------------------------------------------------------
# Query management
# ---------------------------------------------------------------------------

def ac_manage_query(
    db_path: str, action: str, query_name: str,
    sql: Optional[str] = None, new_name: Optional[str] = None,
    confirm: bool = False,
) -> dict:
    """Crea, modifica, renombra, elimina o lee una QueryDef."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()

    if action == "create":
        if not sql:
            raise ValueError("action='create' requiere 'sql'")
        qd = db.CreateQueryDef(query_name, sql)
        return {"action": "created", "query_name": query_name, "sql": sql}

    elif action == "modify":
        if not sql:
            raise ValueError("action='modify' requiere 'sql'")
        try:
            qd = db.QueryDefs(query_name)
        except Exception as exc:
            raise ValueError(f"Query '{query_name}' no encontrada: {exc}")
        qd.SQL = sql
        return {"action": "modified", "query_name": query_name, "sql": sql}

    elif action == "delete":
        if not confirm:
            return {"error": f"Eliminar query '{query_name}' requiere confirm=true"}
        try:
            db.QueryDefs(query_name)  # verify exists
        except Exception as exc:
            raise ValueError(f"Query '{query_name}' no encontrada: {exc}")
        db.QueryDefs.Delete(query_name)
        return {"action": "deleted", "query_name": query_name}

    elif action == "rename":
        if not new_name:
            raise ValueError("action='rename' requiere 'new_name'")
        try:
            qd = db.QueryDefs(query_name)
        except Exception as exc:
            raise ValueError(f"Query '{query_name}' no encontrada: {exc}")
        qd.Name = new_name
        return {"action": "renamed", "old_name": query_name, "new_name": new_name}

    elif action == "get_sql":
        try:
            qd = db.QueryDefs(query_name)
        except Exception as exc:
            raise ValueError(f"Query '{query_name}' no encontrada: {exc}")
        qd_type = _QUERYDEF_TYPE.get(qd.Type, f"Unknown({qd.Type})")
        return {"query_name": query_name, "sql": qd.SQL, "type": qd_type}

    else:
        raise ValueError(f"action debe ser create/modify/delete/rename/get_sql, recibido: '{action}'")


# ---------------------------------------------------------------------------
# Indexes
# ---------------------------------------------------------------------------

def ac_list_indexes(db_path: str, table_name: str) -> dict:
    """Lista los indices de una tabla."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        td = db.TableDefs(table_name)
    except Exception as exc:
        raise ValueError(f"Tabla '{table_name}' no encontrada: {exc}")

    indexes = []
    for i in range(td.Indexes.Count):
        idx = td.Indexes(i)
        fields = []
        for j in range(idx.Fields.Count):
            f = idx.Fields(j)
            fields.append({
                "name": f.Name,
                "order": "desc" if f.Attributes & 1 else "asc",
            })
        indexes.append({
            "name": idx.Name,
            "fields": fields,
            "primary": bool(idx.Primary),
            "unique": bool(idx.Unique),
            "foreign": bool(idx.Foreign),
        })
    return {"table_name": table_name, "count": len(indexes), "indexes": indexes}


def ac_manage_index(
    db_path: str, table_name: str, action: str, index_name: str,
    fields: Optional[list] = None,
    primary: bool = False, unique: bool = False,
) -> dict:
    """Crea o elimina un indice en una tabla."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        td = db.TableDefs(table_name)
    except Exception as exc:
        raise ValueError(f"Tabla '{table_name}' no encontrada: {exc}")

    if action == "create":
        if not fields:
            raise ValueError("action='create' requiere 'fields' [{name, order?}]")
        idx = td.CreateIndex(index_name)
        idx.Primary = primary
        idx.Unique = unique
        for fdef in fields:
            fname = fdef if isinstance(fdef, str) else fdef["name"]
            fld = idx.CreateField(fname)
            if isinstance(fdef, dict) and fdef.get("order", "asc").lower() == "desc":
                fld.Attributes = 1  # dbDescending
            idx.Fields.Append(fld)
        td.Indexes.Append(idx)
        return {
            "action": "created", "table_name": table_name,
            "index_name": index_name, "fields": fields,
            "primary": primary, "unique": unique,
        }

    elif action == "delete":
        try:
            td.Indexes(index_name)  # verify exists
        except Exception as exc:
            raise ValueError(f"Indice '{index_name}' no encontrado en '{table_name}': {exc}")
        td.Indexes.Delete(index_name)
        return {"action": "deleted", "table_name": table_name, "index_name": index_name}

    else:
        raise ValueError(f"action debe ser create/delete, recibido: '{action}'")


# ---------------------------------------------------------------------------
# Compile VBA
# ---------------------------------------------------------------------------

def ac_compile_vba(db_path: str) -> dict:
    """Compila y guarda todos los modulos VBA."""
    app = _Session.connect(db_path)
    try:
        app.RunCommand(_AC_CMD_COMPILE)
    except Exception as exc:
        raise RuntimeError(f"Error de compilacion VBA: {exc}")
    # Invalidate caches — compilation may change module state
    _vbe_code_cache.clear()
    _Session._cm_cache.clear()
    return {"status": "compiled"}


# ---------------------------------------------------------------------------
# Run macro
# ---------------------------------------------------------------------------

def ac_run_macro(db_path: str, macro_name: str) -> dict:
    """Ejecuta una macro de Access."""
    app = _Session.connect(db_path)
    try:
        app.DoCmd.RunMacro(macro_name)
    except Exception as exc:
        raise RuntimeError(f"Error al ejecutar macro '{macro_name}': {exc}")
    return {"macro_name": macro_name, "status": "executed"}


# ---------------------------------------------------------------------------
# Run VBA procedure
# ---------------------------------------------------------------------------

def ac_run_vba(
    db_path: str, procedure: str, args: Optional[list] = None,
) -> dict:
    """Ejecuta un Sub/Function VBA via Application.Run."""
    app = _Session.connect(db_path)
    call_args = args or []
    if len(call_args) > 30:
        raise ValueError("Application.Run soporta max 30 argumentos.")
    try:
        if call_args:
            result = app.Run(procedure, *call_args)
        else:
            result = app.Run(procedure)
    except Exception as exc:
        raise RuntimeError(f"Error al ejecutar '{procedure}': {exc}")
    # COM puede devolver tipos no serializables; convertir a str si es necesario
    if result is not None:
        try:
            json.dumps(result)
        except (TypeError, ValueError):
            result = str(result)
    return {"procedure": procedure, "result": result, "status": "executed"}


# ---------------------------------------------------------------------------
# Output report (PDF, XLS)
# ---------------------------------------------------------------------------

_OUTPUT_FORMATS: dict[str, str] = {
    "pdf": "PDF Format (*.pdf)",
    "xlsx": "Microsoft Excel (*.xlsx)",
    "rtf": "Rich Text Format (*.rtf)",
    "txt": "MS-DOS Text (*.txt)",
}

def ac_output_report(
    db_path: str, report_name: str,
    output_path: Optional[str] = None, fmt: str = "pdf",
) -> dict:
    """Exporta un report a PDF, XLSX, RTF o TXT."""
    app = _Session.connect(db_path)
    fmt_lower = fmt.lower()
    format_string = _OUTPUT_FORMATS.get(fmt_lower)
    if not format_string:
        raise ValueError(f"Formato '{fmt}' no soportado. Usar: {list(_OUTPUT_FORMATS.keys())}")

    ext_map = {"pdf": ".pdf", "xlsx": ".xlsx", "rtf": ".rtf", "txt": ".txt"}
    if output_path is None:
        resolved = str(Path(db_path).resolve())
        db_dir = os.path.dirname(resolved)
        output_path = os.path.join(db_dir, f"{report_name}{ext_map[fmt_lower]}")

    output_path = str(Path(output_path).resolve())
    try:
        app.DoCmd.OutputTo(_AC_OUTPUT_REPORT, report_name, format_string, output_path)
    except Exception as exc:
        raise RuntimeError(f"Error al exportar report '{report_name}': {exc}")

    size = os.path.getsize(output_path) if os.path.exists(output_path) else 0
    return {
        "report_name": report_name, "output_path": output_path,
        "format": fmt_lower, "size_bytes": size,
    }


# ---------------------------------------------------------------------------
# Transfer data (import/export Excel/CSV)
# ---------------------------------------------------------------------------

def ac_transfer_data(
    db_path: str, action: str, file_path: str, table_name: str,
    has_headers: bool = True, file_type: str = "xlsx",
    range_: Optional[str] = None, spec_name: Optional[str] = None,
) -> dict:
    """Importa o exporta datos entre Access y Excel/CSV."""
    app = _Session.connect(db_path)
    file_path = str(Path(file_path).resolve())
    ft = file_type.lower()

    if action == "import":
        transfer_type_spreadsheet = _AC_IMPORT      # 0
        transfer_type_text = 0                       # acImportDelim
    elif action == "export":
        transfer_type_spreadsheet = _AC_EXPORT       # 1
        transfer_type_text = _AC_EXPORT_DELIM        # 2
    else:
        raise ValueError(f"action debe ser 'import' o 'export', recibido: '{action}'")

    try:
        if ft in ("xlsx", "xls", "excel"):
            app.DoCmd.TransferSpreadsheet(
                transfer_type_spreadsheet,
                _AC_SPREADSHEET_XLSX,
                table_name,
                file_path,
                has_headers,
                range_ or "",
            )
        elif ft in ("csv", "txt", "text"):
            app.DoCmd.TransferText(
                transfer_type_text,
                spec_name or "",
                table_name,
                file_path,
                has_headers,
            )
        else:
            raise ValueError(f"file_type '{file_type}' no soportado. Usar: xlsx, csv")
    except ValueError:
        raise
    except Exception as exc:
        raise RuntimeError(f"Error en TransferData ({action} {ft}): {exc}")

    return {"action": action, "file_type": ft, "table_name": table_name, "file_path": file_path}


# ---------------------------------------------------------------------------
# Field properties
# ---------------------------------------------------------------------------

def ac_get_field_properties(db_path: str, table_name: str, field_name: str) -> dict:
    """Lee todas las propiedades de un campo."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        td = db.TableDefs(table_name)
    except Exception as exc:
        raise ValueError(f"Tabla '{table_name}' no encontrada: {exc}")
    try:
        fld = td.Fields(field_name)
    except Exception as exc:
        raise ValueError(f"Campo '{field_name}' no encontrado en '{table_name}': {exc}")

    props = {}
    for i in range(fld.Properties.Count):
        try:
            p = fld.Properties(i)
            val = p.Value
            # Skip binary/complex values
            if isinstance(val, (str, int, float, bool)) or val is None:
                props[p.Name] = val
        except Exception:
            pass  # Some properties throw COM errors when read
    return {"table_name": table_name, "field_name": field_name, "properties": props}


def ac_set_field_property(
    db_path: str, table_name: str, field_name: str,
    property_name: str, value: Any,
) -> dict:
    """Establece una propiedad de un campo. Crea la propiedad si no existe."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        td = db.TableDefs(table_name)
    except Exception as exc:
        raise ValueError(f"Tabla '{table_name}' no encontrada: {exc}")
    try:
        fld = td.Fields(field_name)
    except Exception as exc:
        raise ValueError(f"Campo '{field_name}' no encontrado en '{table_name}': {exc}")

    coerced = _coerce_prop(value)

    # Try to set existing property
    try:
        fld.Properties(property_name).Value = coerced
        return {
            "table_name": table_name, "field_name": field_name,
            "property_name": property_name, "value": coerced, "action": "updated",
        }
    except Exception:
        pass

    # Create property
    if isinstance(coerced, bool):
        prop_type = 1   # dbBoolean
    elif isinstance(coerced, int):
        prop_type = 4   # dbLong
    else:
        prop_type = 10  # dbText
    try:
        prop = fld.CreateProperty(property_name, prop_type, coerced)
        fld.Properties.Append(prop)
        return {
            "table_name": table_name, "field_name": field_name,
            "property_name": property_name, "value": coerced, "action": "created",
        }
    except Exception as exc:
        raise RuntimeError(
            f"No se pudo establecer '{property_name}' en {table_name}.{field_name}: {exc}"
        )


# ---------------------------------------------------------------------------
# Startup options
# ---------------------------------------------------------------------------

def ac_list_startup_options(db_path: str) -> dict:
    """Lista las opciones de startup comunes con sus valores actuales."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    options = []
    for name in _STARTUP_PROPS:
        val = None
        source = "<not set>"
        try:
            val = db.Properties(name).Value
            source = "database"
        except Exception:
            try:
                val = app.GetOption(name)
                source = "application"
            except Exception:
                pass
        options.append({"name": name, "value": val, "source": source})
    return {"count": len(options), "options": options}


# ---------------------------------------------------------------------------
# Definicion de herramientas MCP
# ---------------------------------------------------------------------------
# NOTA PARA EL AGENTE:
# Este MCP es la UNICA forma de leer, modificar o consultar cualquier fichero
# Microsoft Access (.accdb / .mdb) desde Claude Code. No existe ninguna otra
# herramienta ni comando que permita acceder al contenido de una base de datos
# Access. Siempre que el usuario mencione un formulario, modulo, informe, query
# o cualquier objeto de una base de datos Access, DEBES usar estas tools.
# Flujo obligatorio para editar VBA:
#   1. access_list_objects  → descubrir que objetos existen
#   2. access_get_code      → leer el codigo/definicion actual
#   3. (modificar el texto en memoria)
#   4. access_set_code      → escribir el resultado en la BD
# Para ediciones quirurgicas (una funcion, unas pocas lineas) usa las tools VBE:
#   access_vbe_module_info → indice de procedimientos con numeros de linea
#   access_vbe_get_proc    → codigo de un procedimiento concreto
#   access_vbe_replace_lines → reemplazar lineas sin reimportar el modulo entero
# ---------------------------------------------------------------------------
TOOLS = [
    types.Tool(
        name="access_list_objects",
        description="Lista objetos de la BD por tipo (table, module, form, report, query, macro, all). Tablas de sistema (MSys*, ~*) se filtran.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {
                    "type": "string",
                    "enum": ["all", "table", "module", "form", "report", "query", "macro"],
                    "default": "all",
                },
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_get_code",
        description=(
            "Lee codigo/definicion de un objeto Access. "
            "Modulos: codigo .bas. Forms/reports: formato interno (props + VBA)."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report", "query", "macro"]},
                "object_name": {"type": "string", "description": "Nombre del objeto"},
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_set_code",
        description=(
            "Importa codigo en la BD. Sobreescribe si existe, crea si no. "
            "Llamar access_get_code antes para leer el original. "
            "Para forms/reports: soporta CodeBehindForm/CodeBehindReport (VBA se inyecta via VBE). "
            "Hace backup automatico y restaura si falla el import."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report", "query", "macro"]},
                "object_name": {"type": "string", "description": "Nombre del objeto"},
                "code": {"type": "string", "description": "Contenido completo del objeto"},
            },
            "required": ["db_path", "object_type", "object_name", "code"],
        },
    ),
    types.Tool(
        name="access_execute_sql",
        description=(
            "Ejecuta SQL via DAO. SELECT devuelve filas JSON (limit por defecto: 500). "
            "INSERT/UPDATE devuelven affected_rows. "
            "DELETE/DROP/ALTER requieren confirm_destructive=true."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "sql": {"type": "string", "description": "Sentencia SQL"},
                "limit": {"type": "integer", "default": 500,
                          "description": "Max filas para SELECT (default: 500, max: 10000)"},
                "confirm_destructive": {
                    "type": "boolean", "default": False,
                    "description": "Requerido para DELETE/DROP/TRUNCATE/ALTER",
                },
            },
            "required": ["db_path", "sql"],
        },
    ),
    types.Tool(
        name="access_table_info",
        description="Estructura de una tabla via DAO: campos, tipos, tamaño, required, record_count, is_linked.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "table_name": {"type": "string", "description": "Nombre de la tabla"},
            },
            "required": ["db_path", "table_name"],
        },
    ),
    types.Tool(
        name="access_export_structure",
        description=(
            "Genera Markdown con estructura de la BD: modulos con firmas, forms, reports, queries, macros. "
            "Escribe a disco y devuelve el contenido."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "output_path": {"type": "string", "description": "Ruta .md de salida (default: db_structure.md junto a la BD)"},
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_close",
        description="Cierra la sesion COM y libera el .accdb/.mdb.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    # ── VBE line-level tools ─────────────────────────────────────────────────
    types.Tool(
        name="access_vbe_get_lines",
        description="Lee un rango de lineas de un modulo VBA via VBE COM.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del objeto"},
                "start_line":  {"type": "integer", "description": "Primera linea (1-based)"},
                "count":       {"type": "integer", "description": "Numero de lineas a leer"},
            },
            "required": ["db_path", "object_type", "object_name", "start_line", "count"],
        },
    ),
    types.Tool(
        name="access_vbe_get_proc",
        description="Codigo de un procedimiento VBA por nombre. Devuelve start_line, body_line, count, code.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del objeto"},
                "proc_name":   {"type": "string", "description": "Nombre del Sub/Function/Property"},
            },
            "required": ["db_path", "object_type", "object_name", "proc_name"],
        },
    ),
    types.Tool(
        name="access_vbe_module_info",
        description="Indice de procedimientos de un modulo VBA: total_lines, procs [{name, start_line, body_line, count}].",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del objeto"},
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_vbe_replace_lines",
        description=(
            "Reemplaza lineas en un modulo VBA via VBE. "
            "count=0: insercion. new_code='': borrado. Valida limites automaticamente."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del objeto"},
                "start_line":  {"type": "integer", "description": "Primera linea (1-based)"},
                "count":       {"type": "integer", "description": "Lineas a eliminar (0 = insertar)"},
                "new_code":    {"type": "string",  "description": "Codigo nuevo ('' = borrar)"},
            },
            "required": ["db_path", "object_type", "object_name", "start_line", "count", "new_code"],
        },
    ),
    types.Tool(
        name="access_vbe_find",
        description="Busca texto o regex en un modulo VBA. Devuelve matches [{line, content}].",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del objeto"},
                "search_text": {"type": "string", "description": "Texto o patron regex a buscar"},
                "match_case":  {"type": "boolean", "default": False},
                "use_regex": {"type": "boolean", "default": False,
                              "description": "true = interpretar search_text como regex"},
            },
            "required": ["db_path", "object_type", "object_name", "search_text"],
        },
    ),
    types.Tool(
        name="access_vbe_search_all",
        description="Busca texto o regex en TODOS los modulos VBA (modules, forms, reports) de la BD.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "search_text": {"type": "string", "description": "Texto o patron regex a buscar"},
                "match_case":  {"type": "boolean", "default": False},
                "max_results": {"type": "integer", "default": 100,
                                "description": "Max coincidencias totales (default: 100)"},
                "use_regex": {"type": "boolean", "default": False,
                              "description": "true = interpretar search_text como regex"},
            },
            "required": ["db_path", "search_text"],
        },
    ),
    types.Tool(
        name="access_search_queries",
        description="Busca texto o regex en el SQL de TODAS las queries. Devuelve [{query_name, sql}].",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "search_text": {"type": "string", "description": "Texto o patron regex a buscar en el SQL"},
                "match_case": {"type": "boolean", "default": False},
                "max_results": {"type": "integer", "default": 100,
                                "description": "Max queries a devolver (default: 100)"},
                "use_regex": {"type": "boolean", "default": False,
                              "description": "true = interpretar search_text como regex"},
            },
            "required": ["db_path", "search_text"],
        },
    ),
    types.Tool(
        name="access_vbe_replace_proc",
        description="Reemplaza un procedimiento VBA completo por nombre. new_code='' lo elimina.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del objeto"},
                "proc_name":   {"type": "string", "description": "Nombre del Sub/Function/Property"},
                "new_code":    {"type": "string", "description": "Codigo nuevo ('' = eliminar)"},
            },
            "required": ["db_path", "object_type", "object_name", "proc_name", "new_code"],
        },
    ),
    types.Tool(
        name="access_vbe_append",
        description="Añade codigo al final de un modulo VBA.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del objeto"},
                "new_code":    {"type": "string", "description": "Codigo a añadir"},
            },
            "required": ["db_path", "object_type", "object_name", "new_code"],
        },
    ),
    # ── Control-level tools ──────────────────────────────────────────────────
    types.Tool(
        name="access_list_controls",
        description="Lista controles de un form/report con nombre, tipo, caption, control_source, posicion.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del form/report"},
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_get_control",
        description="Definicion completa (Begin...End) de un control por nombre.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del form/report"},
                "control_name": {"type": "string", "description": "Nombre del control"},
            },
            "required": ["db_path", "object_type", "object_name", "control_name"],
        },
    ),
    types.Tool(
        name="access_create_control",
        description=(
            "Crea un control en un form/report via COM. "
            "control_type: nombre o numero. "
            "props especiales: section (0=Detail,1=Header,2=Footer,3=PageHeader,4=PageFooter "
            "o nombre: 'detail','header','footer','reportheader','pageheader'...), "
            "parent, column_name, left, top, width, height."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del form/report"},
                "control_type": {"type": "string", "description": "'CommandButton', 'TextBox', 'Label'... o numero (104, 109, 100...)"},
                "props": {
                    "type": "object",
                    "description": "Propiedades: section, parent, column_name, left, top, width, height, Name, Caption, etc.",
                    "additionalProperties": True,
                },
            },
            "required": ["db_path", "object_type", "object_name", "control_type", "props"],
        },
    ),
    types.Tool(
        name="access_delete_control",
        description="Elimina un control de un form/report via COM.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del form/report"},
                "control_name": {"type": "string", "description": "Nombre del control"},
            },
            "required": ["db_path", "object_type", "object_name", "control_name"],
        },
    ),
    types.Tool(
        name="access_set_control_props",
        description="Modifica propiedades de un control via COM. Numericos/booleanos se convierten automaticamente.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del form/report"},
                "control_name": {"type": "string", "description": "Nombre del control"},
                "props": {
                    "type": "object",
                    "description": "Propiedades a modificar: {Caption: 'X', Left: 1000, Visible: true, ...}",
                    "additionalProperties": True,
                },
            },
            "required": ["db_path", "object_type", "object_name", "control_name", "props"],
        },
    ),
    types.Tool(
        name="access_set_form_property",
        description="Establece propiedades a nivel de form/report (RecordSource, Caption, DefaultView, HasModule, etc.) via COM en vista diseño.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del form/report"},
                "props": {
                    "type": "object",
                    "description": "Propiedades a modificar: {RecordSource: 'Tabla', Caption: 'Titulo', HasModule: true, ...}",
                    "additionalProperties": True,
                },
            },
            "required": ["db_path", "object_type", "object_name", "props"],
        },
    ),
    # ── Database properties ──────────────────────────────────────────────────
    types.Tool(
        name="access_get_db_property",
        description="Lee una propiedad de la BD (CurrentDb.Properties) o opcion de Access (GetOption).",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "name": {"type": "string", "description": "Nombre de la propiedad (ej: AppTitle, StartupForm, AllowBypassKey)"},
            },
            "required": ["db_path", "name"],
        },
    ),
    types.Tool(
        name="access_set_db_property",
        description="Establece una propiedad de la BD o opcion de Access. Crea la propiedad si no existe.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "name": {"type": "string", "description": "Nombre de la propiedad"},
                "value": {"description": "Valor (string, numero o booleano)"},
                "prop_type": {"type": "integer", "description": "Tipo DAO para CreateProperty (1=Boolean, 4=Long, 10=Text). Auto-detectado si se omite"},
            },
            "required": ["db_path", "name", "value"],
        },
    ),
    # ── Linked tables ────────────────────────────────────────────────────────
    types.Tool(
        name="access_list_linked_tables",
        description="Lista tablas vinculadas con source_table, connect_string, is_odbc.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_relink_table",
        description="Cambia connect string de una tabla vinculada y refresca. relink_all=true actualiza todas con la misma conexion.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "table_name": {"type": "string", "description": "Nombre de la tabla vinculada"},
                "new_connect": {"type": "string", "description": "Nueva cadena de conexion"},
                "relink_all": {"type": "boolean", "default": False, "description": "true = relink todas con la misma conexion original"},
            },
            "required": ["db_path", "table_name", "new_connect"],
        },
    ),
    # ── Relationships ────────────────────────────────────────────────────────
    types.Tool(
        name="access_list_relationships",
        description="Lista relaciones entre tablas: nombre, tablas, campos, cascade flags.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_create_relationship",
        description="Crea una relacion entre dos tablas. attributes: 256=cascade update, 4096=cascade delete (combinables con OR).",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "name": {"type": "string", "description": "Nombre de la relacion"},
                "table": {"type": "string", "description": "Tabla principal (lado uno)"},
                "foreign_table": {"type": "string", "description": "Tabla foranea (lado muchos)"},
                "fields": {
                    "type": "array",
                    "description": "[{local: 'ID', foreign: 'FK_ID'}, ...]",
                    "items": {
                        "type": "object",
                        "properties": {"local": {"type": "string"}, "foreign": {"type": "string"}},
                        "required": ["local", "foreign"],
                    },
                },
                "attributes": {"type": "integer", "default": 0, "description": "Bitmask: 256=cascade update, 4096=cascade delete"},
            },
            "required": ["db_path", "name", "table", "foreign_table", "fields"],
        },
    ),
    # ── VBA References ───────────────────────────────────────────────────────
    types.Tool(
        name="access_list_references",
        description="Lista referencias VBA: nombre, GUID, path, is_broken, built_in.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_manage_reference",
        description="Agrega (add) o elimina (remove) una referencia VBA. add: requiere guid o path. remove: requiere name.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "action": {"type": "string", "enum": ["add", "remove"]},
                "name": {"type": "string", "description": "[remove] Nombre de la referencia"},
                "path": {"type": "string", "description": "[add] Ruta al .dll/.tlb/.olb"},
                "guid": {"type": "string", "description": "[add] GUID de la type library"},
                "major": {"type": "integer", "default": 0, "description": "[add+guid] Version mayor"},
                "minor": {"type": "integer", "default": 0, "description": "[add+guid] Version menor"},
            },
            "required": ["db_path", "action"],
        },
    ),
    # ── Compact & Repair ─────────────────────────────────────────────────────
    types.Tool(
        name="access_compact_repair",
        description="Compacta y repara la BD. Cierra, compacta a temp, reemplaza original y reabre. Devuelve sizes.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
            },
            "required": ["db_path"],
        },
    ),
    # ── Query management ────────────────────────────────────────────────────
    types.Tool(
        name="access_manage_query",
        description=(
            "Gestiona QueryDefs: create, modify, delete (requiere confirm=true), rename, get_sql. "
            "create/modify requieren sql."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "action": {"type": "string", "enum": ["create", "modify", "delete", "rename", "get_sql"]},
                "query_name": {"type": "string", "description": "Nombre de la query"},
                "sql": {"type": "string", "description": "[create/modify] SQL de la query"},
                "new_name": {"type": "string", "description": "[rename] Nuevo nombre"},
                "confirm": {"type": "boolean", "default": False, "description": "[delete] Confirmar eliminacion"},
            },
            "required": ["db_path", "action", "query_name"],
        },
    ),
    # ── Indexes ─────────────────────────────────────────────────────────────
    types.Tool(
        name="access_list_indexes",
        description="Lista indices de una tabla: nombre, campos, primary, unique, foreign.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "table_name": {"type": "string", "description": "Nombre de la tabla"},
            },
            "required": ["db_path", "table_name"],
        },
    ),
    types.Tool(
        name="access_manage_index",
        description="Crea o elimina un indice. create requiere fields [{name, order?}]. primary/unique opcionales.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "table_name": {"type": "string", "description": "Nombre de la tabla"},
                "action": {"type": "string", "enum": ["create", "delete"]},
                "index_name": {"type": "string", "description": "Nombre del indice"},
                "fields": {
                    "type": "array", "description": "[create] [{name: 'Field', order: 'asc'|'desc'}]",
                    "items": {
                        "type": "object",
                        "properties": {"name": {"type": "string"}, "order": {"type": "string", "default": "asc"}},
                        "required": ["name"],
                    },
                },
                "primary": {"type": "boolean", "default": False},
                "unique": {"type": "boolean", "default": False},
            },
            "required": ["db_path", "table_name", "action", "index_name"],
        },
    ),
    # ── Compile VBA ─────────────────────────────────────────────────────────
    types.Tool(
        name="access_compile_vba",
        description="Compila y guarda todos los modulos VBA. Devuelve status o error de compilacion.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
            },
            "required": ["db_path"],
        },
    ),
    # ── Run macro ───────────────────────────────────────────────────────────
    types.Tool(
        name="access_run_macro",
        description="Ejecuta una macro de Access por nombre.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "macro_name": {"type": "string", "description": "Nombre de la macro"},
            },
            "required": ["db_path", "macro_name"],
        },
    ),
    # ── Output report ───────────────────────────────────────────────────────
    types.Tool(
        name="access_output_report",
        description="Exporta un report a PDF, XLSX, RTF o TXT. output_path auto-generado si se omite.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "report_name": {"type": "string", "description": "Nombre del report"},
                "output_path": {"type": "string", "description": "Ruta de salida (auto si se omite)"},
                "format": {"type": "string", "default": "pdf", "description": "pdf, xlsx, rtf, txt"},
            },
            "required": ["db_path", "report_name"],
        },
    ),
    # ── Transfer data ───────────────────────────────────────────────────────
    types.Tool(
        name="access_transfer_data",
        description=(
            "Import/export datos entre Access y Excel/CSV. "
            "file_type: xlsx o csv. range solo para Excel, spec_name solo para CSV."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "action": {"type": "string", "enum": ["import", "export"]},
                "file_path": {"type": "string", "description": "Ruta al fichero Excel/CSV"},
                "table_name": {"type": "string", "description": "Nombre de la tabla Access"},
                "has_headers": {"type": "boolean", "default": True},
                "file_type": {"type": "string", "default": "xlsx", "description": "xlsx o csv"},
                "range": {"type": "string", "description": "[xlsx] Rango ej: Sheet1!A1:D100"},
                "spec_name": {"type": "string", "description": "[csv] Import/Export spec guardada en Access"},
            },
            "required": ["db_path", "action", "file_path", "table_name"],
        },
    ),
    # ── Field properties ────────────────────────────────────────────────────
    types.Tool(
        name="access_get_field_properties",
        description="Lee todas las propiedades de un campo: DefaultValue, ValidationRule, Description, Format, etc.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "table_name": {"type": "string", "description": "Nombre de la tabla"},
                "field_name": {"type": "string", "description": "Nombre del campo"},
            },
            "required": ["db_path", "table_name", "field_name"],
        },
    ),
    types.Tool(
        name="access_set_field_property",
        description="Establece una propiedad de un campo. Crea la propiedad si no existe.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "table_name": {"type": "string", "description": "Nombre de la tabla"},
                "field_name": {"type": "string", "description": "Nombre del campo"},
                "property_name": {"type": "string", "description": "Nombre de la propiedad (ej: Description, DefaultValue)"},
                "value": {"description": "Valor (string, numero o booleano)"},
            },
            "required": ["db_path", "table_name", "field_name", "property_name", "value"],
        },
    ),
    # ── Startup options ─────────────────────────────────────────────────────
    types.Tool(
        name="access_list_startup_options",
        description="Lista las 14 opciones de startup comunes con sus valores actuales.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
            },
            "required": ["db_path"],
        },
    ),
    # ── Create database ────────────────────────────────────────────────────
    types.Tool(
        name="access_create_database",
        description="Crea una BD Access (.accdb) vacia. Error si el fichero ya existe.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta para el nuevo .accdb"},
            },
            "required": ["db_path"],
        },
    ),
    # ── Create table via DAO ──────────────────────────────────────────────
    types.Tool(
        name="access_create_table",
        description=(
            "Crea una tabla Access via DAO con soporte completo: tipos, defaults, "
            "descripciones y primary key — todo en una sola llamada. "
            "Mas robusto que CREATE TABLE via SQL, que no soporta DEFAULT ni YESNO en Jet DDL. "
            "Cada campo acepta: name, type, size, required, primary_key, default, description. "
            "Tipos validos: autonumber, long, integer, short, byte, text, memo, currency, "
            "double, single, datetime, boolean/yesno/bit, guid, ole, bigint."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "table_name": {"type": "string", "description": "Nombre de la tabla"},
                "fields": {
                    "type": "array",
                    "description": "Lista de campos [{name, type, size?, required?, primary_key?, default?, description?}]",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string"},
                            "type": {"type": "string", "default": "text"},
                            "size": {"type": "integer"},
                            "required": {"type": "boolean", "default": False},
                            "primary_key": {"type": "boolean", "default": False},
                            "default": {"description": "Valor default (string, numero o booleano)"},
                            "description": {"type": "string"},
                        },
                        "required": ["name"],
                    },
                },
            },
            "required": ["db_path", "table_name", "fields"],
        },
    ),
    # ── Alter table via DAO ───────────────────────────────────────────────
    types.Tool(
        name="access_alter_table",
        description=(
            "Modifica la estructura de una tabla Access via DAO. "
            "Acciones: add_field (con tipo, size, default, description), "
            "delete_field (requiere confirm=true), rename_field. "
            "Mas robusto que ALTER TABLE via SQL en Jet."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "table_name": {"type": "string", "description": "Nombre de la tabla"},
                "action": {"type": "string", "enum": ["add_field", "delete_field", "rename_field"]},
                "field_name": {"type": "string", "description": "Nombre del campo"},
                "new_name": {"type": "string", "description": "[rename_field] Nuevo nombre"},
                "field_type": {"type": "string", "default": "text", "description": "[add_field] Tipo del campo"},
                "size": {"type": "integer", "description": "[add_field] Tamaño para Text"},
                "required": {"type": "boolean", "default": False},
                "default": {"description": "[add_field] Valor default"},
                "description": {"type": "string", "description": "[add_field] Descripcion del campo"},
                "confirm": {"type": "boolean", "default": False, "description": "[delete_field] Confirmar eliminacion"},
            },
            "required": ["db_path", "table_name", "action", "field_name"],
        },
    ),
    # ── Delete object ──────────────────────────────────────────────────────
    types.Tool(
        name="access_delete_object",
        description="Elimina un objeto Access (module, form, report, query, macro). Requiere confirm=true.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report", "query", "macro"]},
                "object_name": {"type": "string", "description": "Nombre del objeto a eliminar"},
                "confirm": {"type": "boolean", "default": False, "description": "Requerido true para confirmar eliminacion"},
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    # ── Run VBA ────────────────────────────────────────────────────────────
    types.Tool(
        name="access_run_vba",
        description="Ejecuta un Sub/Function VBA via Application.Run. procedure puede ser 'Modulo.NombreSub' o solo 'NombreSub'. Devuelve result si es Function.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "procedure": {"type": "string", "description": "Nombre del procedimiento VBA (ej: MiModulo.MiSub)"},
                "args": {
                    "type": "array",
                    "description": "Argumentos opcionales (max 30)",
                    "items": {},
                },
            },
            "required": ["db_path", "procedure"],
        },
    ),
    # ── Delete relationship ────────────────────────────────────────────────
    types.Tool(
        name="access_delete_relationship",
        description="Elimina una relacion entre tablas por nombre.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "name": {"type": "string", "description": "Nombre de la relacion a eliminar"},
            },
            "required": ["db_path", "name"],
        },
    ),
    # ── Find usages ────────────────────────────────────────────────────────
    types.Tool(
        name="access_find_usages",
        description=(
            "Busca texto o regex en VBA, SQL de queries y propiedades de controles "
            "(ControlSource, RecordSource, RowSource, DefaultValue, ValidationRule). "
            "Devuelve resultados agrupados: vba_matches, query_matches, control_matches."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "search_text": {"type": "string", "description": "Texto o patron regex a buscar"},
                "match_case": {"type": "boolean", "default": False},
                "max_results": {"type": "integer", "default": 200,
                                "description": "Max coincidencias totales (default: 200)"},
                "use_regex": {"type": "boolean", "default": False,
                              "description": "true = interpretar search_text como regex"},
            },
            "required": ["db_path", "search_text"],
        },
    ),
]

# ---------------------------------------------------------------------------
# MCP Server
# ---------------------------------------------------------------------------
server = Server("access-mcp")


@server.list_tools()
async def list_tools() -> list[types.Tool]:
    return TOOLS

@server.list_prompts()
async def list_prompts() -> list[types.Prompt]:
    return [
        types.Prompt(
            name="access-workflow",
            description=(
                "Instrucciones de uso del MCP access para trabajar con bases de datos "
                "Microsoft Access (.accdb/.mdb) desde Claude Code."
            ),
            arguments=[
                types.PromptArgument(
                    name="db_path",
                    description="Ruta completa al fichero .accdb o .mdb",
                    required=False,
                )
            ],
        )
    ]

@server.get_prompt()
async def get_prompt(name: str, arguments: dict | None) -> types.GetPromptResult:
    db_path = (arguments or {}).get("db_path", "<ruta_al_fichero.accdb>")
    return types.GetPromptResult(
        description="Workflow obligatorio para trabajar con bases de datos Access",
        messages=[
            types.PromptMessage(
                role="user",
                content=types.TextContent(
                    type="text",
                    text=f"""
Estoy trabajando con una base de datos Microsoft Access: {db_path}

REGLAS OBLIGATORIAS para el agente:
1. Cualquier operacion sobre ficheros .accdb o .mdb DEBE hacerse a traves del MCP access.
   No existe ninguna otra herramienta ni comando de shell que pueda leer o modificar Access.

2. Flujo obligatorio para editar VBA o definiciones de objetos:
   a) access_list_objects  → descubrir que objetos existen (formularios, modulos, informes...)
   b) access_get_code      → leer el codigo actual del objeto
   c) modificar el texto
   d) access_set_code      → guardar el resultado en la BD

3. Para ediciones de pocas lineas (mas eficiente):
   a) access_vbe_module_info  → indice de procedimientos con numeros de linea
   b) access_vbe_get_proc     → codigo del procedimiento concreto
   c) access_vbe_replace_lines → reemplazar solo las lineas modificadas

4. Nunca adivines nombres de formularios, modulos o controles.
   Siempre llama primero a access_list_objects o access_list_controls.

5. Nunca escribas codigo VBA sin haber leido antes el original con access_get_code
   o access_vbe_get_proc. El formato interno de Access es estricto.
""",
                ),
            )
        ],
    )

@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[types.TextContent]:
    # Logging seguro: mostrar código solo como longitud para proteger datos sensibles
    safe_args = {}
    for k, v in arguments.items():
        if k == "code":
            safe_args[k] = f"<VBA code: {len(v)} chars>"
        else:
            safe_args[k] = v
    log.info(">>> %s  %s", name, safe_args)

    try:
        if name == "access_list_objects":
            result = ac_list_objects(
                arguments["db_path"],
                arguments.get("object_type", "all"),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_get_code":
            text = ac_get_code(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
            )

        elif name == "access_set_code":
            text = ac_set_code(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["code"],
            )

        elif name == "access_execute_sql":
            result = ac_execute_sql(
                arguments["db_path"],
                arguments["sql"],
                int(arguments.get("limit", 500)),
                bool(arguments.get("confirm_destructive", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_table_info":
            result = ac_table_info(arguments["db_path"], arguments["table_name"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_export_structure":
            text = ac_export_structure(
                arguments["db_path"],
                arguments.get("output_path"),
            )

        elif name == "access_close":
            _Session.quit()
            text = "Sesion Access cerrada correctamente."

        # ── VBE line-level ───────────────────────────────────────────────────
        elif name == "access_vbe_get_lines":
            text = ac_vbe_get_lines(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                int(arguments["start_line"]),
                int(arguments["count"]),
            )

        elif name == "access_vbe_get_proc":
            result = ac_vbe_get_proc(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["proc_name"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_vbe_module_info":
            result = ac_vbe_module_info(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_vbe_replace_lines":
            text = ac_vbe_replace_lines(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                int(arguments["start_line"]),
                int(arguments["count"]),
                arguments["new_code"],
            )

        elif name == "access_vbe_find":
            result = ac_vbe_find(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["search_text"],
                bool(arguments.get("match_case", False)),
                bool(arguments.get("use_regex", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_vbe_search_all":
            result = ac_vbe_search_all(
                arguments["db_path"],
                arguments["search_text"],
                bool(arguments.get("match_case", False)),
                int(arguments.get("max_results", 100)),
                bool(arguments.get("use_regex", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_search_queries":
            result = ac_search_queries(
                arguments["db_path"],
                arguments["search_text"],
                bool(arguments.get("match_case", False)),
                int(arguments.get("max_results", 100)),
                bool(arguments.get("use_regex", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_vbe_replace_proc":
            text = ac_vbe_replace_proc(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["proc_name"],
                arguments["new_code"],
            )

        elif name == "access_vbe_append":
            text = ac_vbe_append(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["new_code"],
            )

        # ── Control-level ────────────────────────────────────────────────────
        elif name == "access_list_controls":
            result = ac_list_controls(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_get_control":
            result = ac_get_control(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["control_name"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_create_control":
            result = ac_create_control(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["control_type"],
                dict(arguments.get("props", {})),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_delete_control":
            text = ac_delete_control(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["control_name"],
            )

        elif name == "access_set_control_props":
            result = ac_set_control_props(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                arguments["control_name"],
                dict(arguments.get("props", {})),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_set_form_property":
            result = ac_set_form_property(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                dict(arguments.get("props", {})),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Database properties ───────────────────────────────────────────
        elif name == "access_get_db_property":
            result = ac_get_db_property(arguments["db_path"], arguments["name"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_set_db_property":
            result = ac_set_db_property(
                arguments["db_path"],
                arguments["name"],
                arguments["value"],
                arguments.get("prop_type"),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Linked tables ─────────────────────────────────────────────────
        elif name == "access_list_linked_tables":
            result = ac_list_linked_tables(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_relink_table":
            result = ac_relink_table(
                arguments["db_path"],
                arguments["table_name"],
                arguments["new_connect"],
                bool(arguments.get("relink_all", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Relationships ─────────────────────────────────────────────────
        elif name == "access_list_relationships":
            result = ac_list_relationships(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_create_relationship":
            result = ac_create_relationship(
                arguments["db_path"],
                arguments["name"],
                arguments["table"],
                arguments["foreign_table"],
                arguments["fields"],
                int(arguments.get("attributes", 0)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── VBA References ────────────────────────────────────────────────
        elif name == "access_list_references":
            result = ac_list_references(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_manage_reference":
            result = ac_manage_reference(
                arguments["db_path"],
                arguments["action"],
                name=arguments.get("name"),
                path=arguments.get("path"),
                guid=arguments.get("guid"),
                major=int(arguments.get("major", 0)),
                minor=int(arguments.get("minor", 0)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Compact & Repair ──────────────────────────────────────────────
        elif name == "access_compact_repair":
            result = ac_compact_repair(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Query management ─────────────────────────────────────────────
        elif name == "access_manage_query":
            result = ac_manage_query(
                arguments["db_path"],
                arguments["action"],
                arguments["query_name"],
                sql=arguments.get("sql"),
                new_name=arguments.get("new_name"),
                confirm=bool(arguments.get("confirm", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Indexes ──────────────────────────────────────────────────────
        elif name == "access_list_indexes":
            result = ac_list_indexes(arguments["db_path"], arguments["table_name"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_manage_index":
            result = ac_manage_index(
                arguments["db_path"],
                arguments["table_name"],
                arguments["action"],
                arguments["index_name"],
                fields=arguments.get("fields"),
                primary=bool(arguments.get("primary", False)),
                unique=bool(arguments.get("unique", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Compile VBA ──────────────────────────────────────────────────
        elif name == "access_compile_vba":
            result = ac_compile_vba(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Run macro ────────────────────────────────────────────────────
        elif name == "access_run_macro":
            result = ac_run_macro(arguments["db_path"], arguments["macro_name"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Output report ────────────────────────────────────────────────
        elif name == "access_output_report":
            result = ac_output_report(
                arguments["db_path"],
                arguments["report_name"],
                output_path=arguments.get("output_path"),
                fmt=arguments.get("format", "pdf"),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Transfer data ────────────────────────────────────────────────
        elif name == "access_transfer_data":
            result = ac_transfer_data(
                arguments["db_path"],
                arguments["action"],
                arguments["file_path"],
                arguments["table_name"],
                has_headers=bool(arguments.get("has_headers", True)),
                file_type=arguments.get("file_type", "xlsx"),
                range_=arguments.get("range"),
                spec_name=arguments.get("spec_name"),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Field properties ─────────────────────────────────────────────
        elif name == "access_get_field_properties":
            result = ac_get_field_properties(
                arguments["db_path"],
                arguments["table_name"],
                arguments["field_name"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_set_field_property":
            result = ac_set_field_property(
                arguments["db_path"],
                arguments["table_name"],
                arguments["field_name"],
                arguments["property_name"],
                arguments["value"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Startup options ──────────────────────────────────────────────
        elif name == "access_list_startup_options":
            result = ac_list_startup_options(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Create database ─────────────────────────────────────────────
        elif name == "access_create_database":
            result = ac_create_database(arguments["db_path"])
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Create table via DAO ─────────────────────────────────────────
        elif name == "access_create_table":
            result = ac_create_table(
                arguments["db_path"],
                arguments["table_name"],
                arguments["fields"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Alter table via DAO ──────────────────────────────────────────
        elif name == "access_alter_table":
            result = ac_alter_table(
                arguments["db_path"],
                arguments["table_name"],
                arguments["action"],
                arguments["field_name"],
                new_name=arguments.get("new_name"),
                field_type=arguments.get("field_type", "text"),
                size=int(arguments.get("size", 0)),
                required=bool(arguments.get("required", False)),
                default=arguments.get("default"),
                description=arguments.get("description"),
                confirm=bool(arguments.get("confirm", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Delete object ────────────────────────────────────────────────
        elif name == "access_delete_object":
            result = ac_delete_object(
                arguments["db_path"],
                arguments["object_type"],
                arguments["object_name"],
                confirm=bool(arguments.get("confirm", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Run VBA ──────────────────────────────────────────────────────
        elif name == "access_run_vba":
            result = ac_run_vba(
                arguments["db_path"],
                arguments["procedure"],
                args=arguments.get("args"),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Delete relationship ──────────────────────────────────────────
        elif name == "access_delete_relationship":
            result = ac_delete_relationship(
                arguments["db_path"],
                arguments["name"],
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        # ── Find usages ─────────────────────────────────────────────────
        elif name == "access_find_usages":
            result = ac_find_usages(
                arguments["db_path"],
                arguments["search_text"],
                bool(arguments.get("match_case", False)),
                int(arguments.get("max_results", 200)),
                bool(arguments.get("use_regex", False)),
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        else:
            text = f"ERROR: herramienta desconocida '{name}'"

    except Exception as exc:
        log.error("Error en %s: %s", name, exc, exc_info=True)

        # Build detailed error message for the LLM
        tb_lines = traceback.format_exc().splitlines()

        # Create safe representation of arguments (hide full code)
        safe_args_display = {}
        for k, v in arguments.items():
            if k == "code":
                safe_args_display[k] = f"<VBA code provided: length {len(v)} chars>"
            else:
                safe_args_display[k] = v

        error_msg = (
            f"ERROR in tool '{name}'\n"
            f"Type: {type(exc).__name__}\n"
            f"Message: {exc}\n\n"
            f"Arguments received:\n{json.dumps(safe_args_display, indent=2, ensure_ascii=False)}\n\n"
            f"Stack trace (last 5 lines):\n" + "\n".join(tb_lines[-5:])
        )
        text = error_msg

    log.info("<<< %s  OK", name)
    return [types.TextContent(type="text", text=text)]


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
async def _main() -> None:
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


if __name__ == "__main__":
    asyncio.run(_main())
