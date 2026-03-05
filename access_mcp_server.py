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
    "RecSrcDt", "GUID",
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
    search_text: str, match_case: bool = False
) -> dict:
    """
    Busca texto en un módulo y devuelve todas las líneas que coinciden.
    Usa _vbe_code_cache para evitar releer el módulo si ya fue leído.
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    cache_key = f"{object_type}:{object_name}"
    all_code = _cm_all_code(cm, cache_key)
    if not all_code:
        return {"found": False, "match_count": 0, "matches": []}
    needle = search_text if match_case else search_text.lower()
    matches: list[dict] = []
    for i, raw_line in enumerate(all_code.splitlines(), start=1):
        haystack = raw_line if match_case else raw_line.lower()
        if needle in haystack:
            matches.append({"line": i, "content": raw_line.rstrip("\r")})
    return {"found": bool(matches), "match_count": len(matches), "matches": matches}


def ac_vbe_search_all(
    db_path: str, search_text: str, match_case: bool = False
) -> dict:
    """
    Busca texto en TODOS los módulos VBA (modules, forms, reports) de la BD.
    Devuelve {total_matches, results: [{object_type, object_name, matches: [{line, content}]}]}.
    """
    app = _Session.connect(db_path)
    objects = ac_list_objects(db_path, "all")
    needle = search_text if match_case else search_text.lower()
    results: list[dict] = []
    total = 0

    for obj_type in ("module", "form", "report"):
        for obj_name in objects.get(obj_type, []):
            try:
                cm = _get_code_module(app, obj_type, obj_name)
                cache_key = f"{obj_type}:{obj_name}"
                all_code = _cm_all_code(cm, cache_key)
                if not all_code:
                    continue
                obj_matches: list[dict] = []
                for i, raw_line in enumerate(all_code.splitlines(), start=1):
                    haystack = raw_line if match_case else raw_line.lower()
                    if needle in haystack:
                        obj_matches.append({"line": i, "content": raw_line.rstrip("\r")})
                if obj_matches:
                    results.append({
                        "object_type": obj_type,
                        "object_name": obj_name,
                        "matches": obj_matches,
                    })
                    total += len(obj_matches)
            except Exception:
                continue  # skip objects without accessible CodeModule

    return {"total_matches": total, "results": results}


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
    depth = 0
    for i in range(form_begin, len(lines)):
        s = lines[i].rstrip("\r\n").lstrip()
        if re.match(r"^Begin\b", s):
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
    section     = int(_coerce_prop(props.pop("section",     0)))
    parent      = str(props.pop("parent",      ""))
    column_name = str(props.pop("column_name", ""))
    left        = int(_coerce_prop(props.pop("left",   -1)))
    top         = int(_coerce_prop(props.pop("top",    -1)))
    width       = int(_coerce_prop(props.pop("width",  -1)))
    height      = int(_coerce_prop(props.pop("height", -1)))

    _open_in_design(app, object_type, object_name)
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


# ---------------------------------------------------------------------------
# Logica de negocio
# ---------------------------------------------------------------------------
def ac_list_objects(db_path: str, object_type: str = "all") -> dict:
    """Devuelve un dict {tipo: [nombres...]} con los objetos de la BD."""
    app = _Session.connect(db_path)

    # CurrentData  → objetos de datos (tablas, queries)
    # CurrentProject → objetos de codigo (forms, reports, modulos, macros)
    containers = {
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
        result[k] = [col.Item(i).Name for i in range(col.Count)]

    return result


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


def ac_set_code(db_path: str, object_type: str, name: str, code: str) -> str:
    """
    Importa texto como definicion de un objeto Access (crea o sobreescribe).
    Para formularios e informes re-inyecta automáticamente las secciones binarias
    (PrtMip, PrtDevMode…) desde el export actual, de modo que el caller no necesita
    incluirlas en el código que envía.
    """
    if object_type not in AC_TYPE:
        raise ValueError(
            f"object_type '{object_type}' invalido. Validos: {list(AC_TYPE)}"
        )
    app = _Session.connect(db_path)

    # Si el código no contiene secciones binarias (fue devuelto por ac_get_code
    # con el filtrado activo), las restauramos desde el form/report actual.
    if object_type in ("form", "report") and not any(
        s in code for s in _BINARY_SECTIONS
    ):
        log.info("ac_set_code: restaurando secciones binarias para '%s'", name)
        code = _restore_binary_sections(app, object_type, name, code)

    fd, tmp = tempfile.mkstemp(suffix=".txt", prefix="access_mcp_")
    os.close(fd)
    try:
        # Módulos VBA (.bas) esperan ANSI/cp1252; forms/reports/queries/macros esperan UTF-16LE con BOM
        enc = "cp1252" if object_type == "module" else "utf-16"
        _write_tmp(tmp, code, encoding=enc)
        app.LoadFromText(AC_TYPE[object_type], name, tmp)
        # Invalidar caches para este objeto (el código y los controles cambiaron)
        cache_key = f"{object_type}:{name}"
        _vbe_code_cache.pop(cache_key, None)
        _Session._cm_cache.pop(cache_key, None)
        _parsed_controls_cache.pop(cache_key, None)
        return f"OK: '{name}' ({object_type}) importado correctamente en {db_path}"
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


def ac_execute_sql(db_path: str, sql: str) -> dict:
    """
    Ejecuta SQL en la BD via DAO.
    SELECT  → devuelve {rows: [...], count: N}
    Otros   → devuelve {affected_rows: N}
    """
    app = _Session.connect(db_path)
    db = app.CurrentDb()

    if sql.strip().upper().startswith("SELECT"):
        rs = db.OpenRecordset(sql)
        fields = [rs.Fields(i).Name for i in range(rs.Fields.Count)]
        rows: list[dict] = []
        if not rs.EOF:
            rs.MoveFirst()
            while not rs.EOF:
                rows.append(
                    {fields[i]: rs.Fields(i).Value for i in range(len(fields))}
                )
                rs.MoveNext()
        rs.Close()
        return {"rows": rows, "count": len(rows)}
    else:
        db.Execute(sql)
        return {"affected_rows": db.RecordsAffected}


# Mapa DAO Type → nombre legible
_DAO_FIELD_TYPE: dict[int, str] = {
    1: "Boolean", 2: "Byte", 3: "Integer", 4: "Long", 5: "Currency",
    6: "Single", 7: "Double", 8: "Date/Time", 10: "Text",
    11: "OLE Object", 12: "Memo", 15: "GUID", 16: "BigInt",
    20: "Decimal",
}


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

    return (
        f"OK: estructura exportada a `{output_path}` — "
        f"{len(modules)} módulos, {len(forms)} formularios, "
        f"{len(reports)} informes, {len(queries)} queries."
    )


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
        description=(
            "USAR SIEMPRE COMO PRIMER PASO al trabajar con una base de datos Access (.accdb/.mdb). "
            "Lista todos los objetos de la BD agrupados por tipo: modulos VBA, formularios, "
            "informes, queries y macros. "
            "Sin llamar a esta tool el agente no sabe que formularios, modulos ni objetos "
            "existen en la base de datos — no intentes adivinarlos. "
            "object_type puede ser 'module', 'form', 'report', 'query', 'macro' o 'all'."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {
                    "type": "string",
                    "description": "Ruta completa al fichero .accdb o .mdb",
                },
                "object_type": {
                    "type": "string",
                    "enum": ["all", "module", "form", "report", "query", "macro"],
                    "default": "all",
                    "description": "Tipo de objeto a listar. Valores: 'module' (VBA módulo), 'form' (formulario), 'report' (informe), 'query' (consulta), 'macro' (macro), 'all' (todos). Por defecto: 'all'",
                },
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_get_code",
        description=(
            "Lee y devuelve el codigo VBA o la definicion completa de cualquier objeto "
            "de una base de datos Access (.accdb/.mdb). "
            "DEBES llamar a esta tool antes de modificar cualquier objeto — nunca escribas "
            "codigo VBA de Access sin haber leido primero el original con esta tool. "
            "Para modulos VBA devuelve codigo .bas limpio. "
            "Para formularios e informes devuelve el formato interno de Access "
            "(propiedades + seccion Class Module con el VBA). "
            "Es la unica forma de leer el codigo fuente de un objeto Access desde Claude Code."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {
                    "type": "string",
                    "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)",
                },
                "object_type": {
                    "type": "string",
                    "enum": ["module", "form", "report", "query", "macro"],
                    "description": "Tipo del objeto: 'module' (VBA módulo), 'form' (formulario), 'report' (informe), 'query' (consulta), 'macro' (macro)",
                },
                "object_name": {
                    "type": "string",
                    "description": "Nombre exacto del objeto (case-sensitive). Para obtener los nombres válidos, usa primero access_list_objects",
                },
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_set_code",
        description=(
            "Escribe (importa) el codigo modificado de un objeto en la base de datos Access. "
            "Si el objeto ya existe lo SOBREESCRIBE; si no existe lo CREA. "
            "Es la unica forma de guardar cambios en VBA o en la definicion de un objeto Access. "
            "OBLIGATORIO: llamar siempre a access_get_code primero para obtener "
            "el texto original y modificar solo lo necesario — especialmente en formularios "
            "e informes donde el formato incluye propiedades de controles que no debes alterar."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {
                    "type": "string",
                    "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)",
                },
                "object_type": {
                    "type": "string",
                    "enum": ["module", "form", "report", "query", "macro"],
                    "description": "Tipo del objeto: 'module' (VBA módulo), 'form' (formulario), 'report' (informe), 'query' (consulta), 'macro' (macro)",
                },
                "object_name": {
                    "type": "string",
                    "description": "Nombre exacto del objeto (case-sensitive). Para obtener los nombres válidos, usa primero access_list_objects",
                },
                "code": {
                    "type": "string",
                    "description": "Contenido completo del objeto. IMPORTANTE: para módulos VBA usar texto limpio; para formularios/informes usar el formato completo devuelto por access_get_code. Nunca modifiques el texto sin haber antes obtenido el original con access_get_code",
                },
            },
            "required": ["db_path", "object_type", "object_name", "code"],
        },
    ),
    types.Tool(
        name="access_execute_sql",
        description=(
            "Ejecuta SQL directamente en una base de datos Access (.accdb/.mdb) via DAO. "
            "Es la unica forma de consultar o modificar datos de tablas Access desde Claude Code. "
            "SELECT devuelve las filas como JSON. "
            "INSERT/UPDATE/DELETE devuelven el numero de filas afectadas. "
            "Funciona con tablas locales y tablas linkadas a SQL Server."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {
                    "type": "string",
                    "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)",
                },
                "sql": {
                    "type": "string",
                    "description": "Sentencia SQL a ejecutar (SELECT, INSERT, UPDATE, DELETE, etc.)",
                },
            },
            "required": ["db_path", "sql"],
        },
    ),
    types.Tool(
        name="access_table_info",
        description=(
            "Devuelve la estructura de una tabla local o linkada de Access: "
            "campos con nombre, tipo DAO, tamaño, required; record_count; is_linked. "
            "Usa DAO TableDef.Fields — mas preciso que SELECT TOP 1 para ver tipos de datos."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {
                    "type": "string",
                    "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)",
                },
                "table_name": {
                    "type": "string",
                    "description": "Nombre exacto de la tabla (case-sensitive)",
                },
            },
            "required": ["db_path", "table_name"],
        },
    ),
    types.Tool(
        name="access_export_structure",
        description=(
            "Genera un fichero Markdown con la estructura completa de una base de datos Access: "
            "modulos VBA con sus funciones/subs, formularios, informes, queries y macros. "
            "USAR AL INICIO de cualquier proyecto Access para obtener un indice completo "
            "de la BD antes de empezar a trabajar. Tambien util para actualizar el indice "
            "despues de añadir o eliminar objetos."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {
                    "type": "string",
                    "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)",
                },
                "output_path": {
                    "type": "string",
                    "description": (
                        "Ruta donde guardar el fichero .md de estructura. "
                        "Si no se especifica, se crea 'db_structure.md' en el mismo directorio que la base de datos"
                    ),
                },
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_close",
        description=(
            "Cierra la sesion COM con Access y libera el fichero .accdb/.mdb. "
            "Llamar al terminar una sesion de edicion para que otros procesos puedan abrir la BD."
        ),
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    # ── VBE line-level tools ─────────────────────────────────────────────────
    types.Tool(
        name="access_vbe_get_lines",
        description=(
            "Lee un rango de lineas de un modulo VBA de Access directamente via VBE COM, "
            "sin exportar el fichero entero. Mas eficiente que access_get_code "
            "cuando solo necesitas inspeccionar una zona concreta antes de editarla. "
            "object_type: 'module', 'form' o 'report'."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"], "description": "Tipo: 'module' (VBA módulo), 'form' (formulario), 'report' (informe)"},
                "object_name": {"type": "string", "description": "Nombre exacto (case-sensitive). Usa access_list_objects para obtener los nombres válidos"},
                "start_line":  {"type": "integer", "description": "Primera linea a leer (1-based)"},
                "count":       {"type": "integer", "description": "Numero de lineas a leer"},
            },
            "required": ["db_path", "object_type", "object_name", "start_line", "count"],
        },
    ),
    types.Tool(
        name="access_vbe_get_proc",
        description=(
            "Devuelve el codigo completo de un procedimiento VBA (Sub/Function/Property) "
            "de un objeto Access buscandolo por nombre via VBE. "
            "Mucho mas eficiente que access_get_code cuando solo interesa una funcion concreta. "
            "Devuelve: start_line, body_line (donde empieza el cuerpo), count, code."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"], "description": "Tipo: 'module' (VBA módulo), 'form' (formulario), 'report' (informe)"},
                "object_name": {"type": "string", "description": "Nombre exacto (case-sensitive). Usa access_list_objects para obtener los nombres válidos"},
                "proc_name":   {"type": "string", "description": "Nombre exacto del Sub/Function/Property"},
            },
            "required": ["db_path", "object_type", "object_name", "proc_name"],
        },
    ),
    types.Tool(
        name="access_vbe_module_info",
        description=(
            "Devuelve el numero total de lineas y el indice completo de procedimientos "
            "(con start_line, body_line y count de cada uno) de un modulo VBA de Access. "
            "USAR ANTES de access_vbe_get_proc o access_vbe_replace_lines para saber "
            "exactamente donde esta cada funcion sin descargar el codigo completo."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"], "description": "Tipo: 'module' (VBA módulo), 'form' (formulario), 'report' (informe)"},
                "object_name": {"type": "string", "description": "Nombre exacto (case-sensitive). Usa access_list_objects para obtener los nombres válidos"},
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_vbe_replace_lines",
        description=(
            "Reemplaza lineas en un modulo VBA de Access directamente via VBE COM, "
            "sin exportar ni reimportar el modulo entero. "
            "Usar para ediciones quirurgicas: modificar una funcion, añadir declaraciones, etc. "
            "Borra 'count' lineas desde 'start_line' e inserta 'new_code' en su lugar. "
            "count=0 → insercion pura. new_code='' → borrado puro. "
            "new_code puede ser multilinea (separado por \\n). "
            "Valida limites automaticamente: si count excede el final del modulo, se ajusta. "
            "Devuelve el nuevo total de lineas del modulo. "
            "Para reemplazar funciones enteras, preferir access_vbe_replace_proc (mas seguro)."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"], "description": "Tipo: 'module' (VBA módulo), 'form' (formulario), 'report' (informe)"},
                "object_name": {"type": "string", "description": "Nombre exacto (case-sensitive). Usa access_list_objects para obtener los nombres válidos"},
                "start_line":  {"type": "integer", "description": "Primera linea afectada (1-based)"},
                "count":       {"type": "integer", "description": "Lineas a eliminar (0 = solo insertar)"},
                "new_code":    {"type": "string",  "description": "Codigo nuevo ('' = solo borrar)"},
            },
            "required": ["db_path", "object_type", "object_name", "start_line", "count", "new_code"],
        },
    ),
    types.Tool(
        name="access_vbe_find",
        description=(
            "Busca texto en un modulo VBA de Access y devuelve todas las lineas que coinciden "
            "con su numero de linea. Usar para localizar variables, llamadas a funciones, "
            "o cualquier patron de texto antes de editar. "
            "Devuelve: found, match_count, matches [{line, content}]."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"], "description": "Tipo: 'module' (VBA módulo), 'form' (formulario), 'report' (informe)"},
                "object_name": {"type": "string", "description": "Nombre exacto (case-sensitive). Usa access_list_objects para obtener los nombres válidos"},
                "search_text": {"type": "string", "description": "Texto a buscar"},
                "match_case":  {"type": "boolean", "description": "Distinguir mayusculas (default: false)", "default": False},
            },
            "required": ["db_path", "object_type", "object_name", "search_text"],
        },
    ),
    types.Tool(
        name="access_vbe_search_all",
        description=(
            "Busca texto en TODOS los modulos VBA de la base de datos Access "
            "(modules, forms, reports) de una sola vez. "
            "Ideal para refactoring o localizar donde se usa una funcion/variable "
            "sin tener que llamar access_vbe_find una vez por cada objeto. "
            "Devuelve: total_matches, results [{object_type, object_name, matches: [{line, content}]}]."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "search_text": {"type": "string", "description": "Texto a buscar en todos los modulos"},
                "match_case":  {"type": "boolean", "description": "Distinguir mayusculas (default: false)", "default": False},
            },
            "required": ["db_path", "search_text"],
        },
    ),
    types.Tool(
        name="access_vbe_replace_proc",
        description=(
            "Reemplaza un procedimiento VBA completo (Sub/Function/Property) por nombre. "
            "Calcula los limites automaticamente via COM — NO necesitas saber los numeros de linea. "
            "Si new_code esta vacio, elimina el procedimiento. "
            "Mas seguro que access_vbe_replace_lines para reemplazar funciones enteras."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"], "description": "Tipo: 'module' (VBA módulo), 'form' (formulario), 'report' (informe)"},
                "object_name": {"type": "string", "description": "Nombre exacto (case-sensitive). Usa access_list_objects para obtener los nombres válidos"},
                "proc_name":   {"type": "string", "description": "Nombre exacto del Sub/Function/Property a reemplazar o eliminar"},
                "new_code":    {"type": "string", "description": "Codigo nuevo completo del procedimiento. Vacio ('') para eliminar el proc"},
            },
            "required": ["db_path", "object_type", "object_name", "proc_name", "new_code"],
        },
    ),
    types.Tool(
        name="access_vbe_append",
        description=(
            "Añade codigo al final de un modulo VBA de Access. "
            "Mas seguro que replace_lines para insertar nuevas funciones o declaraciones "
            "sin necesidad de calcular numeros de linea."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"], "description": "Tipo: 'module' (VBA módulo), 'form' (formulario), 'report' (informe)"},
                "object_name": {"type": "string", "description": "Nombre exacto (case-sensitive). Usa access_list_objects para obtener los nombres válidos"},
                "new_code":    {"type": "string", "description": "Codigo a añadir al final del modulo. Puede ser multilinea (separado por \\n)"},
            },
            "required": ["db_path", "object_type", "object_name", "new_code"],
        },
    ),
    # ── Control-level tools ──────────────────────────────────────────────────
    types.Tool(
        name="access_list_controls",
        description=(
            "Lista todos los controles de un formulario o informe Access con sus "
            "propiedades clave: nombre, tipo, caption, control_source, posicion y linea. "
            "USAR como paso previo a access_get_control o access_set_control_props "
            "para conocer los nombres exactos de los controles antes de modificarlos. "
            "No incluye controles dentro de TabControl (estan a mayor profundidad)."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del formulario o informe"},
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_get_control",
        description=(
            "Devuelve la definicion completa (bloque Begin...End) de un control concreto "
            "de un formulario o informe Access, buscado por nombre. "
            "Para modificar propiedades del control usa access_set_control_props."
            "No lo uses para nombres de controles sin nombre."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del formulario o informe"},
                "control_name": {"type": "string", "description": "Nombre exacto del control (case-sensitive). Usa access_list_controls para obtenerlos"},
            },
            "required": ["db_path", "object_type", "object_name", "control_name"],
        },
    ),
    types.Tool(
        name="access_create_control",
        description=(
            "Crea un control nuevo en un formulario o informe Access via COM (CreateControl). "
            "Abre el objeto en vista diseno, crea el control, asigna propiedades y guarda. "
            "control_type: nombre ('CommandButton', 'TextBox', 'Label'...) o numero (104, 109, 100...). "
            "props claves especiales: section (0=Detail,1=Header,2=Footer), parent, column_name, "
            "left, top, width, height. Resto de claves: Name, Caption, ControlSource, OnClick, etc."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del formulario o informe"},
                "control_type": {"type": "string", "description": "Tipo del control: nombre ('CommandButton', 'TextBox', 'Label') o valor numerico (104, 109, 100...)"},
                "props": {
                    "type": "object",
                    "description": "Propiedades del control. Claves especiales: section (0=Detail,1=Header,2=Footer), parent (nombre del contenedor), column_name, left, top, width, height. Otras propiedades: Name, Caption, ControlSource, Visible, OnClick, etc.",
                    "additionalProperties": True,
                },
            },
            "required": ["db_path", "object_type", "object_name", "control_type", "props"],
        },
    ),
    types.Tool(
        name="access_delete_control",
        description=(
            "Elimina un control de un formulario o informe Access via COM (DeleteControl). "
            "Abre el objeto en vista diseno, elimina el control y guarda."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del formulario o informe"},
                "control_name": {"type": "string", "description": "Nombre exacto del control a eliminar (case-sensitive). Usa access_list_controls para obtenerlo"},
            },
            "required": ["db_path", "object_type", "object_name", "control_name"],
        },
    ),
    types.Tool(
        name="access_set_control_props",
        description=(
            "Modifica propiedades de un control existente en un formulario o informe Access "
            "via COM en vista diseno. Asigna las propiedades indicadas y guarda. "
            "Los valores numericos y booleanos se convierten automaticamente. "
            "Devuelve {applied: [...], errors: {...}} para confirmar que se aplico."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Ruta completa al fichero .accdb o .mdb (ej: C:\\ruta\\database.accdb)"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del formulario o informe"},
                "control_name": {"type": "string", "description": "Nombre exacto del control (case-sensitive). Usa access_list_controls para obtenerlo"},
                "props": {
                    "type": "object",
                    "description": "Diccionario de propiedades a modificar. Ejemplos: {Caption: 'Nuevo Texto', Left: 1000, Visible: true, ControlSource: 'Campo'}. Valores numéricos y booleanos se convierten automáticamente",
                    "additionalProperties": True,
                },
            },
            "required": ["db_path", "object_type", "object_name", "control_name", "props"],
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
            result = ac_execute_sql(arguments["db_path"], arguments["sql"])
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
            )
            text = json.dumps(result, ensure_ascii=False, indent=2)

        elif name == "access_vbe_search_all":
            result = ac_vbe_search_all(
                arguments["db_path"],
                arguments["search_text"],
                bool(arguments.get("match_case", False)),
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
