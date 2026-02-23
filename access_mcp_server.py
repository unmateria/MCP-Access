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
# Sesion COM — singleton, mantiene Access vivo entre llamadas
# ---------------------------------------------------------------------------
class _Session:
    """
    Mantiene una instancia de Access.Application entre tool calls.
    Si se pide una BD distinta a la abierta, cierra la actual y abre la nueva.
    """
    _app: Optional[Any] = None
    _db_open: Optional[str] = None

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
    Requiere 'Confiar en el acceso al modelo de objetos de proyectos VBA'
    habilitado en las opciones de confianza de Access.
    """
    if object_type not in _VBE_PREFIX:
        raise ValueError(
            f"object_type '{object_type}' no soporta VBE. Usa 'module', 'form' o 'report'."
        )
    component_name = _VBE_PREFIX[object_type] + object_name
    try:
        project = app.VBE.VBProjects(1)
        component = project.VBComponents(component_name)
        return component.CodeModule
    except Exception as exc:
        raise RuntimeError(
            f"No se pudo acceder al CodeModule '{component_name}'. "
            f"¿Está habilitado 'Confiar en el acceso al modelo de objetos de proyectos VBA' "
            f"en las opciones de confianza de Access?\nError: {exc}"
        )


def ac_vbe_get_lines(
    db_path: str, object_type: str, object_name: str,
    start_line: int, count: int
) -> str:
    """Lee un rango de líneas sin exportar el módulo entero."""
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    total = cm.CountOfLines
    if start_line < 1 or start_line > total:
        raise ValueError(f"start_line {start_line} fuera de rango (1-{total})")
    actual = min(count, total - start_line + 1)
    return cm.Lines(start_line, actual)


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
        start = cm.ProcStartLine(proc_name, 0)   # 0 = vbext_pk_Proc
        body  = cm.ProcBodyLine(proc_name, 0)
        count = cm.ProcCountLines(proc_name, 0)
    except Exception as exc:
        raise RuntimeError(
            f"Procedimiento '{proc_name}' no encontrado en '{object_name}': {exc}"
        )
    code = cm.Lines(start, count)
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
    total = cm.CountOfLines
    procs: list[dict] = []
    if total > 0:
        all_code = cm.Lines(1, total)
        seen: set[str] = set()
        for i, raw_line in enumerate(all_code.splitlines(), start=1):
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
                    body   = cm.ProcBodyLine(pname, 0)
                    pcount = cm.ProcCountLines(pname, 0)
                    procs.append({"name": pname, "start_line": i,
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
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    if count > 0:
        cm.DeleteLines(start_line, count)
    inserted = 0
    if new_code:
        # Access VBA espera \r\n como separador de líneas
        normalized = new_code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
        cm.InsertLines(start_line, normalized)
        inserted = len(new_code.splitlines())
    end = start_line + count - 1 if count > 0 else start_line
    return (
        f"OK: líneas {start_line}–{end} reemplazadas "
        f"({count} eliminadas, {inserted} insertadas)"
    )


def ac_vbe_find(
    db_path: str, object_type: str, object_name: str,
    search_text: str, match_case: bool = False
) -> dict:
    """
    Busca texto en un módulo y devuelve todas las líneas que coinciden.
    Una sola llamada COM para leer el módulo; la búsqueda se hace en Python.
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    total = cm.CountOfLines
    if total == 0:
        return {"found": False, "match_count": 0, "matches": []}
    all_code = cm.Lines(1, total)
    needle = search_text if match_case else search_text.lower()
    matches: list[dict] = []
    for i, raw_line in enumerate(all_code.splitlines(), start=1):
        haystack = raw_line if match_case else raw_line.lower()
        if needle in haystack:
            matches.append({"line": i, "content": raw_line.rstrip("\r")})
    return {"found": bool(matches), "match_count": len(matches), "matches": matches}


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
      ctrl_indent    — indentación de los bloques Begin de controles directos
      form_begin_idx — índice 0-based de la línea "Begin Form/Report"
      form_end_idx   — índice 0-based del "End" que cierra el bloque Form/Report
    Nota: controles dentro de TabControl aparecen a mayor profundidad y no se listan.
    """
    lines = form_text.splitlines(keepends=True)
    result = {
        "controls": [],
        "form_indent": "",
        "ctrl_indent": "",
        "form_begin_idx": -1,
        "form_end_idx": -1,
    }

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

    form_indent = result["form_indent"]
    form_begin = result["form_begin_idx"]

    # 2. Encontrar el "End" que cierra el bloque Form (rastreo de profundidad)
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

    # 3. Detectar ctrl_indent: primer "Begin" sin calificador después de Begin Form
    for i in range(form_begin + 1, result["form_end_idx"]):
        raw = lines[i].rstrip("\r\n")
        s = raw.lstrip()
        if s == "Begin":
            result["ctrl_indent"] = raw[: len(raw) - len(s)]
            break

    ctrl_indent = result["ctrl_indent"]
    if not ctrl_indent:
        return result  # form sin controles

    # 4. Extraer bloques de controles al nivel ctrl_indent
    i = form_begin + 1
    while i < result["form_end_idx"]:
        raw = lines[i].rstrip("\r\n")
        s = raw.lstrip()
        indent = raw[: len(raw) - len(s)]

        # Saltar ClassModule
        if re.match(r"^Begin\s+ClassModule\s*$", s, re.IGNORECASE):
            break

        if s == "Begin" and indent == ctrl_indent:
            ctrl_start = i
            block: list[str] = [lines[i]]
            props: dict[str, str] = {}
            depth = 1
            ctrl_end = i
            j = i + 1
            while j < len(lines):
                bl = lines[j]
                bl_r = bl.rstrip("\r\n")
                bl_s = bl_r.lstrip()
                block.append(bl)
                if depth == 1:
                    m = re.match(r"^(\w+)\s*=(.*)", bl_s)
                    if m:
                        props[m.group(1)] = m.group(2).strip().strip('"')
                if bl_s == "Begin":
                    depth += 1
                elif bl_s == "End":
                    depth -= 1
                    if depth == 0:
                        ctrl_end = j
                        break
                j += 1

            name = props.get("Name", props.get("ControlName", ""))
            try:
                ctype = int(props.get("ControlType", -1))
            except (ValueError, TypeError):
                ctype = -1

            result["controls"].append({
                "name":           name,
                "control_type":   ctype,
                "type_name":      _CTRL_TYPE.get(ctype, f"Type{ctype}"),
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


def ac_list_controls(db_path: str, object_type: str, object_name: str) -> dict:
    """
    Lista todos los controles directos de un formulario o informe con sus
    propiedades clave (sin raw_block para no saturar el resultado).
    Nota: controles dentro de TabControl no aparecen (están a mayor profundidad).
    """
    if object_type not in ("form", "report"):
        raise ValueError("ac_list_controls solo admite object_type 'form' o 'report'")
    text = ac_get_code(db_path, object_type, object_name)
    parsed = _parse_controls(text)
    return {
        "count": len(parsed["controls"]),
        "controls": [
            {k: v for k, v in c.items() if k != "raw_block"}
            for c in parsed["controls"]
        ],
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
    text = ac_get_code(db_path, object_type, object_name)
    parsed = _parse_controls(text)
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
        _write_tmp(tmp, code, encoding="utf-16")
        app.LoadFromText(AC_TYPE[object_type], name, tmp)
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
    lines.append(f"## Módulos VBA ({len(modules)})\n")
    for mod_name in modules:
        lines.append(f"### `{mod_name}`")
        try:
            code = ac_get_code(db_path, "module", mod_name)
            sigs = []
            for line in code.splitlines():
                stripped = line.strip()
                if re.match(
                    r"^(Public\s+|Private\s+|Friend\s+)?(Function|Sub)\s+\w+",
                    stripped,
                    re.IGNORECASE,
                ):
                    # Limpiar: quitar Attribute lines y quedarse solo con la firma
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
TOOLS = [
    types.Tool(
        name="access_list_objects",
        description=(
            "Lista los objetos de una base de datos Access (.accdb/.mdb). "
            "Devuelve un JSON con los nombres agrupados por tipo. "
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
                    "description": "Tipo de objeto a listar (por defecto: all)",
                },
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_get_code",
        description=(
            "Exporta el codigo o definicion completa de un objeto Access y lo devuelve como texto. "
            "Para modulos VBA devuelve codigo .bas limpio. "
            "Para formularios e informes devuelve el formato interno de Access "
            "(propiedades + seccion Class Module con el VBA). "
            "Usa este tool ANTES de access_set_code para obtener el texto original."
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
                    "enum": ["module", "form", "report", "query", "macro"],
                    "description": "Tipo del objeto",
                },
                "object_name": {
                    "type": "string",
                    "description": "Nombre exacto del objeto (sensible a mayusculas)",
                },
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_set_code",
        description=(
            "Importa texto como definicion de un objeto en Access. "
            "Si el objeto ya existe lo SOBREESCRIBE; si no existe lo CREA. "
            "IMPORTANTE: llama siempre a access_get_code primero para obtener "
            "el texto original y modificar solo lo necesario, especialmente "
            "en formularios e informes donde el formato incluye propiedades de controles."
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
                    "enum": ["module", "form", "report", "query", "macro"],
                    "description": "Tipo del objeto",
                },
                "object_name": {
                    "type": "string",
                    "description": "Nombre exacto del objeto",
                },
                "code": {
                    "type": "string",
                    "description": "Contenido completo del objeto en formato texto de Access",
                },
            },
            "required": ["db_path", "object_type", "object_name", "code"],
        },
    ),
    types.Tool(
        name="access_execute_sql",
        description=(
            "Ejecuta una sentencia SQL en la base de datos via DAO. "
            "SELECT devuelve las filas como JSON. "
            "INSERT/UPDATE/DELETE devuelven el numero de filas afectadas. "
            "Util para tablas locales (las tablas linkadas a SQL Server tambien funcionan)."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {
                    "type": "string",
                    "description": "Ruta completa al fichero .accdb o .mdb",
                },
                "sql": {
                    "type": "string",
                    "description": "Sentencia SQL a ejecutar",
                },
            },
            "required": ["db_path", "sql"],
        },
    ),
    types.Tool(
        name="access_export_structure",
        description=(
            "Genera un fichero Markdown (db_structure.md) con la estructura completa "
            "de la base de datos: todos los modulos VBA con sus funciones/subs, "
            "formularios, informes, queries y macros. "
            "Usalo al inicio del proyecto para crear el indice, y cada vez que "
            "añadas o elimines objetos en la BD para mantenerlo actualizado."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {
                    "type": "string",
                    "description": "Ruta completa al fichero .accdb o .mdb",
                },
                "output_path": {
                    "type": "string",
                    "description": (
                        "Ruta donde guardar el fichero .md. "
                        "Por defecto: db_structure.md en el mismo directorio que el .accdb"
                    ),
                },
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_close",
        description=(
            "Cierra la sesion COM con Access y libera el fichero .accdb. "
            "Recomendado al terminar una sesion de edicion para que otros "
            "procesos puedan abrir la BD."
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
            "Lee un rango de líneas de un módulo VBA directamente via VBE COM, "
            "sin exportar el fichero entero. Ideal para inspeccionar una zona concreta "
            "antes de editarla. object_type: 'module', 'form' o 'report'."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path":     {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del módulo/formulario/informe"},
                "start_line":  {"type": "integer", "description": "Primera línea a leer (1-based)"},
                "count":       {"type": "integer", "description": "Número de líneas a leer"},
            },
            "required": ["db_path", "object_type", "object_name", "start_line", "count"],
        },
    ),
    types.Tool(
        name="access_vbe_get_proc",
        description=(
            "Devuelve el código completo de un procedimiento (Sub/Function/Property) "
            "buscándolo por nombre via VBE. Mucho más eficiente que access_get_code "
            "cuando solo interesa una función. "
            "Devuelve: start_line, body_line (donde empieza el cuerpo), count, code."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path":     {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del módulo/formulario/informe"},
                "proc_name":   {"type": "string", "description": "Nombre exacto del Sub/Function/Property"},
            },
            "required": ["db_path", "object_type", "object_name", "proc_name"],
        },
    ),
    types.Tool(
        name="access_vbe_module_info",
        description=(
            "Devuelve el número total de líneas y la lista de procedimientos "
            "(con start_line, body_line y count de cada uno) de un módulo VBA. "
            "Úsalo como índice rápido para saber qué hay y dónde antes de editar, "
            "sin necesidad de descargar el código completo."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path":     {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del módulo/formulario/informe"},
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_vbe_replace_lines",
        description=(
            "Reemplaza líneas en un módulo VBA directamente via VBE COM, "
            "sin exportar ni reimportar el módulo entero. "
            "Borra 'count' líneas desde 'start_line' e inserta 'new_code' en su lugar. "
            "count=0 → inserción pura. new_code='' → borrado puro. "
            "new_code puede ser multilínea (separado por \\n)."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path":     {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del módulo/formulario/informe"},
                "start_line":  {"type": "integer", "description": "Primera línea afectada (1-based)"},
                "count":       {"type": "integer", "description": "Líneas a eliminar (0 = solo insertar)"},
                "new_code":    {"type": "string",  "description": "Código nuevo ('' = solo borrar)"},
            },
            "required": ["db_path", "object_type", "object_name", "start_line", "count", "new_code"],
        },
    ),
    types.Tool(
        name="access_vbe_find",
        description=(
            "Busca texto en un módulo VBA y devuelve todas las líneas que coinciden "
            "con su número de línea. Una sola llamada COM; la búsqueda se hace en Python. "
            "Devuelve: found, match_count, matches [{line, content}]."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path":     {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del módulo/formulario/informe"},
                "search_text": {"type": "string", "description": "Texto a buscar"},
                "match_case":  {"type": "boolean", "description": "Distinguir mayúsculas (default: false)", "default": False},
            },
            "required": ["db_path", "object_type", "object_name", "search_text"],
        },
    ),
    # ── Control-level tools ──────────────────────────────────────────────────
    types.Tool(
        name="access_list_controls",
        description=(
            "Lista todos los controles directos de un formulario o informe con sus "
            "propiedades clave: nombre, tipo, caption, control_source, posición y línea. "
            "No incluye controles dentro de TabControl (están a mayor profundidad). "
            "Usa este tool como paso previo a access_get_control / access_set_control_props."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path":     {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Nombre del formulario o informe"},
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_get_control",
        description=(
            "Devuelve la definición completa (bloque Begin...End) de un control "
            "concreto de un formulario o informe, buscado por nombre. "
            "El raw_block es solo lectura; para modificar propiedades usa access_set_control_props."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path":      {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type":  {"type": "string", "enum": ["form", "report"]},
                "object_name":  {"type": "string", "description": "Nombre del formulario o informe"},
                "control_name": {"type": "string", "description": "Nombre exacto del control"},
            },
            "required": ["db_path", "object_type", "object_name", "control_name"],
        },
    ),
    # ── Control COM tools (CreateControl / DeleteControl / design mode) ──────
    types.Tool(
        name="access_create_control",
        description=(
            "Crea un control nuevo en un formulario o informe via COM (CreateControl). "
            "Abre el objeto en vista diseño, crea el control, asigna propiedades y guarda. "
            "control_type: nombre ('CommandButton', 'TextBox', 'Label'...) o número (104, 109, 100...). "
            "props: dict con propiedades del control. Claves especiales (pasadas a CreateControl): "
            "section (0=Detail, 1=Header, 2=Footer), parent, column_name, left, top, width, height. "
            "Resto de claves se asignan como propiedades COM (Name, Caption, ControlSource, OnClick...)."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path":      {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type":  {"type": "string", "enum": ["form", "report"]},
                "object_name":  {"type": "string", "description": "Nombre del formulario o informe"},
                "control_type": {"type": "string", "description": "Tipo: nombre ('CommandButton') o número (104)"},
                "props": {
                    "type": "object",
                    "description": (
                        "Propiedades del control. Especiales: section, parent, column_name, "
                        "left, top, width, height. Resto: Name, Caption, ControlSource, OnClick, etc."
                    ),
                    "additionalProperties": True,
                },
            },
            "required": ["db_path", "object_type", "object_name", "control_type", "props"],
        },
    ),
    types.Tool(
        name="access_delete_control",
        description=(
            "Elimina un control de un formulario o informe via COM (DeleteControl). "
            "Abre el objeto en vista diseño, elimina el control y guarda."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path":      {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type":  {"type": "string", "enum": ["form", "report"]},
                "object_name":  {"type": "string", "description": "Nombre del formulario o informe"},
                "control_name": {"type": "string", "description": "Nombre exacto del control a eliminar"},
            },
            "required": ["db_path", "object_type", "object_name", "control_name"],
        },
    ),
    types.Tool(
        name="access_set_control_props",
        description=(
            "Modifica propiedades de un control existente via COM en vista diseño. "
            "Abre el formulario/informe en diseño, asigna las propiedades y guarda. "
            "Los valores numéricos y booleanos se convierten automáticamente. "
            "Devuelve {applied: [...], errors: {...}} para saber qué se aplicó."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path":      {"type": "string", "description": "Ruta al .accdb/.mdb"},
                "object_type":  {"type": "string", "enum": ["form", "report"]},
                "object_name":  {"type": "string", "description": "Nombre del formulario o informe"},
                "control_name": {"type": "string", "description": "Nombre exacto del control"},
                "props": {
                    "type": "object",
                    "description": "Propiedades a modificar: {Caption: 'OK', Left: 1000, Visible: true, ...}",
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


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[types.TextContent]:
    # Log sin el campo 'code' para no saturar el log
    log_args = {k: v for k, v in arguments.items() if k != "code"}
    log.info(">>> %s  %s", name, log_args)

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
        text = f"ERROR: {exc}"

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
