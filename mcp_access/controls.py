"""
Controls: parsing, listing, design-view manipulation, export/import text.
"""

import os
import re
import tempfile
from typing import Any, Optional

from .core import (
    AC_TYPE, _Session, _vbe_code_cache, _parsed_controls_cache, log,
    invalidate_object_caches, invalidate_all_caches,
)
from .constants import (
    CTRL_TYPE, CONTAINER_TYPES, AC_CONTROL_TYPE_NAMES,
    AC_DESIGN, AC_FORM, AC_REPORT, AC_SAVE_YES, AC_SAVE_NO,
    CTRL_TYPE_BY_NAME, SECTION_MAP,
)
from .helpers import coerce_prop, serialize_value, read_tmp, write_tmp


# ---------------------------------------------------------------------------
# _parse_controls — extract control blocks from SaveAsText export
# ---------------------------------------------------------------------------

def _parse_controls(form_text: str) -> dict:
    """
    Parses the exported text of a form/report and extracts the control blocks.
    Returns a dict with:
      controls       — list of controls with their properties and position in the text
      form_indent    — indentation of the "Begin Form/Report" line
      ctrl_indent    — (legacy, kept for compatibility) indent of the first control found
      form_begin_idx — 0-based index of the "Begin Form/Report" line
      form_end_idx   — 0-based index of the "End" that closes the Form/Report block

    Access export structure:
      Begin Form              <- form level
          Begin               <- defaults block (contains Begin Label, Begin CommandButton with default props)
          End
          Begin Section       <- section (Detail, FormHeader, FormFooter)
              ...
              Begin           <- control container within the section
                  Begin Label <- REAL CONTROL (has Name =, ControlType =, etc.)
                  End
                  Begin CommandButton
                  End
              End
          End
          Begin ClassModule   <- VBA code of the form
          End
      End Form

    The parser looks for controls WITHIN sections, identifying them by having
    a known type (Begin <TypeName>) where TypeName is a value from CTRL_TYPE.
    """
    lines = form_text.splitlines(keepends=True)
    result: dict = {
        "controls": [],
        "form_indent": "",
        "ctrl_indent": "",
        "form_begin_idx": -1,
        "form_end_idx": -1,
    }

    # Set of type names for fast detection
    ctrl_type_names = {v for v in CTRL_TYPE.values()}

    # 1. Locate "Begin Form" or "Begin Report"
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

    # 2. Find the "End" that closes the Form/Report block (depth tracking)
    #    Important: detect both "Begin ..." and "Property = Begin" (e.g.: NameMap = Begin)
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

    # 3. Scan ALL "Begin <TypeName>" blocks inside the form/report
    #    where TypeName matches a known control type.
    #    Controls can be at any depth within sections.
    i = form_begin + 1
    container_stack: list[tuple[str, int]] = []  # [(name, ctrl_end_idx), ...]
    while i < result["form_end_idx"]:
        # Clean up containers we've already passed
        while container_stack and i > container_stack[-1][1]:
            container_stack.pop()
        raw = lines[i].rstrip("\r\n")
        s = raw.lstrip()
        indent = raw[: len(raw) - len(s)]

        # Skip ClassModule — contains no controls, only VBA
        if re.match(r"^Begin\s+ClassModule\s*$", s, re.IGNORECASE):
            break

        # Detect "Begin <TypeName>" where TypeName is a known control type
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
                # Only parse properties at the top level of the control (depth == 1)
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

            # Count ConditionalFormat entries in the raw block
            raw_text = "".join(block)
            fmt_count = sum(
                1 for bl in block
                if re.match(r"^\s+ConditionalFormat\d*\s*=\s*Begin\s*$",
                            bl.rstrip("\r\n"))
            )

            # Save ctrl_indent of the first control found (legacy compat)
            if not result["ctrl_indent"] and name:
                result["ctrl_indent"] = indent

            ctrl_entry = {
                "name":           name,
                "control_type":   ctype,
                "type_name":      CTRL_TYPE.get(ctype, m_ctrl.group(1)),
                "caption":        props.get("Caption", ""),
                "control_source": props.get("ControlSource", ""),
                "left":           props.get("Left", ""),
                "top":            props.get("Top", ""),
                "width":          props.get("Width", ""),
                "height":         props.get("Height", ""),
                "visible":        props.get("Visible", ""),
                "start_line":     ctrl_start + 1,  # 1-based
                "end_line":       ctrl_end + 1,     # 1-based inclusive
                "raw_block":      raw_text,
            }
            if fmt_count > 0:
                ctrl_entry["format_conditions"] = fmt_count
            # Annotate parent if inside a container
            if container_stack:
                ctrl_entry["parent"] = container_stack[-1][0]
            result["controls"].append(ctrl_entry)
            if m_ctrl.group(1) in CONTAINER_TYPES:
                container_stack.append((name, ctrl_end))
                i = ctrl_start + 1  # re-scan inside the container
            else:
                i = ctrl_end + 1
            continue

        i += 1

    return result


# ---------------------------------------------------------------------------
# _get_parsed_controls — cache wrapper
# ---------------------------------------------------------------------------

def _get_parsed_controls(db_path: str, object_type: str, object_name: str) -> dict:
    """
    Returns the result of _parse_controls using _parsed_controls_cache.
    If not in cache, exports and parses (and saves in cache for future calls).
    """
    cache_key = f"{object_type}:{object_name}"
    if cache_key not in _parsed_controls_cache:
        # Lazy import to avoid circular dependency with .code module
        from .code import ac_get_code
        text = ac_get_code(db_path, object_type, object_name)
        _parsed_controls_cache[cache_key] = _parse_controls(text)
    return _parsed_controls_cache[cache_key]


# ---------------------------------------------------------------------------
# ac_list_controls / ac_get_control
# ---------------------------------------------------------------------------

def ac_list_controls(db_path: str, object_type: str, object_name: str) -> dict:
    if object_type not in ("form", "report"):
        raise ValueError("ac_list_controls only accepts object_type 'form' or 'report'")
    parsed = _get_parsed_controls(db_path, object_type, object_name)
    controls = [
        {k: v for k, v in c.items() if k != "raw_block"}
        for c in parsed["controls"]
        if c.get("name", "").strip()  # exclude controls without a name
    ]
    return {
        "count": len(controls),
        "controls": controls,
    }


def ac_get_control(
    db_path: str, object_type: str, object_name: str, control_name: str
) -> dict:
    """
    Returns the full definition (raw_block) of a specific control by name.
    The raw_block can be passed modified to ac_set_control to update the control.
    """
    if object_type not in ("form", "report"):
        raise ValueError("ac_get_control only accepts object_type 'form' or 'report'")
    parsed = _get_parsed_controls(db_path, object_type, object_name)
    for c in parsed["controls"]:
        if c["name"].lower() == control_name.lower():
            return c
    names = [c["name"] for c in parsed["controls"]]
    raise ValueError(
        f"Control '{control_name}' not found in '{object_name}'. "
        f"Available controls: {names}"
    )


# ---------------------------------------------------------------------------
# Section / control type resolution
# ---------------------------------------------------------------------------

def _resolve_section(section_val) -> int:
    """Accepts number (0) or name ('detail', 'header', 'reportheader', etc.)."""
    if isinstance(section_val, str):
        key = section_val.lower().replace(" ", "").replace("_", "")
        if key in SECTION_MAP:
            return SECTION_MAP[key]
        try:
            return int(key)
        except ValueError:
            valid = sorted(set(SECTION_MAP.keys()))
            raise ValueError(
                f"Section '{section_val}' not recognized. "
                f"Valid: {valid} or number (0-8)"
            )
    return int(coerce_prop(section_val))


def _resolve_ctrl_type(ctrl_type) -> int:
    """Accepts name ('CommandButton') or number (104)."""
    if isinstance(ctrl_type, int):
        return ctrl_type
    try:
        return int(ctrl_type)
    except (ValueError, TypeError):
        key = str(ctrl_type).lower()
        if key in CTRL_TYPE_BY_NAME:
            return CTRL_TYPE_BY_NAME[key]
        raise ValueError(
            f"Unknown control type: '{ctrl_type}'. "
            f"Use a number or one of: {list(CTRL_TYPE.values())}"
        )


# ---------------------------------------------------------------------------
# Design-view helpers
# ---------------------------------------------------------------------------

def _open_in_design(app: Any, object_type: str, object_name: str) -> None:
    """Opens a form/report in Design view."""
    try:
        if object_type == "form":
            app.DoCmd.OpenForm(object_name, AC_DESIGN)
        else:
            app.DoCmd.OpenReport(object_name, AC_DESIGN)
    except Exception as exc:
        raise RuntimeError(
            f"Could not open '{object_name}' in Design view. "
            f"If it is open in Normal view, close it first.\nError: {exc}"
        )


def _save_and_close(app: Any, object_type: str, object_name: str) -> None:
    """Saves and closes a form/report open in Design view."""
    ac_type = AC_FORM if object_type == "form" else AC_REPORT
    try:
        app.DoCmd.Close(ac_type, object_name, AC_SAVE_YES)
    except Exception as exc:
        log.warning("Error closing '%s': %s", object_name, exc)


def _get_design_obj(app: Any, object_type: str, object_name: str) -> Any:
    """Returns the Form or Report object open in Design view."""
    return app.Forms(object_name) if object_type == "form" else app.Reports(object_name)


# ---------------------------------------------------------------------------
# ac_create_control
# ---------------------------------------------------------------------------

def ac_create_control(
    db_path: str, object_type: str, object_name: str,
    control_type: Any, props: dict, class_name: Optional[str] = None,
) -> dict:
    """
    Creates a new control in a form/report by opening it in Design view.
    control_type: name ('CommandButton') or number (104).
    props: dict of properties. Special keys passed to CreateControl:
      section (default 0=Detail), parent (''), column_name (''),
      left, top, width, height (twips; -1 = automatic).
    The rest are assigned as COM properties on the created control.

    For ActiveX controls (type 119 = acCustomControl), pass class_name with the ProgID
    of the control (e.g.: 'Shell.Explorer.2', 'MSCAL.Calendar.7') to initialize the OLE.

    For WebBrowser, use type 128 (acWebBrowser) instead of 119 — creates a native
    WebBrowser control without needing ActiveX.
    """
    if object_type not in ("form", "report"):
        raise ValueError("Only 'form' or 'report'")

    app = _Session.connect(db_path)
    ctype = _resolve_ctrl_type(control_type)

    # Extract positional/structural params from props (not assigned as prop)
    props = dict(props)  # copia para no mutar el original
    section     = _resolve_section(props.pop("section", 0))
    parent      = str(props.pop("parent",      ""))
    column_name = str(props.pop("column_name", ""))
    left        = int(coerce_prop(props.pop("left",   -1)))
    top         = int(coerce_prop(props.pop("top",    -1)))
    width       = int(coerce_prop(props.pop("width",  -1)))
    height      = int(coerce_prop(props.pop("height", -1)))

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
            section_names = [k for k, v in SECTION_MAP.items() if v == section]
            raise RuntimeError(
                f"Error creating control in section={section} "
                f"({', '.join(section_names) or 'unknown'}): {exc}. "
                f"Verify that the section exists in the {object_type}. "
                f"Valid sections: 0=Detail, 1=Header, 2=Footer, "
                f"3=PageHeader, 4=PageFooter"
            )

        # ActiveX: set ProgID via Class property to initialize OLE
        if class_name and ctype == 119:  # acCustomControl
            try:
                ctrl.Class = class_name
            except Exception as exc:
                log.warning("Could not set Class='%s': %s", class_name, exc)

        errors: dict[str, str] = {}
        for key, val in props.items():
            try:
                setattr(ctrl, key, coerce_prop(val))
            except Exception as exc:
                errors[key] = str(exc)

        result: dict = {
            "name":         ctrl.Name,
            "control_type": ctype,
            "type_name":    CTRL_TYPE.get(ctype, f"Type{ctype}"),
        }
        if errors:
            result["property_errors"] = errors
    finally:
        _save_and_close(app, object_type, object_name)
        # Invalidate caches — form changed in Design view
        invalidate_object_caches(object_type, object_name)

    return result


# ---------------------------------------------------------------------------
# ac_delete_control
# ---------------------------------------------------------------------------

def ac_delete_control(
    db_path: str, object_type: str, object_name: str, control_name: str
) -> str:
    """Deletes a control from a form/report by opening it in Design view."""
    if object_type not in ("form", "report"):
        raise ValueError("Only 'form' or 'report'")

    app = _Session.connect(db_path)
    _open_in_design(app, object_type, object_name)
    try:
        if object_type == "form":
            app.DeleteControl(object_name, control_name)
        else:
            app.DeleteReportControl(object_name, control_name)
    finally:
        _save_and_close(app, object_type, object_name)
        # Invalidate caches — form changed in Design view
        invalidate_object_caches(object_type, object_name)

    return f"OK: control '{control_name}' deleted from '{object_name}'"


# ---------------------------------------------------------------------------
# ac_export_text / ac_import_text
# ---------------------------------------------------------------------------

def ac_export_text(db_path: str, object_type: str, object_name: str,
                   output_path: str) -> dict:
    """Exports a form/report/module/query/macro as text via SaveAsText.

    Does NOT open the object in Design view — does not recalculate positions.
    The resulting file is UTF-16 LE with BOM.
    """
    if object_type not in ("form", "report", "module", "query", "macro"):
        raise ValueError("object_type must be form, report, module, query or macro")
    app = _Session.connect(db_path)
    app.SaveAsText(AC_TYPE[object_type], object_name, output_path)
    file_size = os.path.getsize(output_path)
    return {"path": output_path, "file_size": file_size,
            "object": object_name, "type": object_type}


def _split_code_behind(code: str) -> tuple[str, str]:
    """
    Splits a form/report text into (form_text, vba_code).
    If the code contains 'CodeBehindForm' or 'CodeBehindReport', it splits it.
    Returns (form_text_without_vba, vba_code) where vba_code may be empty.
    The form_text is cleaned of HasModule if there is VBA (it will be injected later).
    """
    # Find the line that marks the start of VBA code
    for marker in ("CodeBehindForm", "CodeBehindReport"):
        idx = code.find(marker)
        if idx != -1:
            form_part = code[:idx].rstrip() + "\n"
            vba_part = code[idx:].split("\n", 1)
            vba_code = vba_part[1] if len(vba_part) > 1 else ""
            # Remove Attribute VB_ lines from VBA (auto-generated)
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
    Injects VBA code into a form/report after importing it.
    Activates HasModule by opening in Design view, then uses VBE to insert the code.
    """
    if not vba_code.strip():
        return

    # 1. Open in Design view and activate HasModule
    _open_in_design(app, object_type, name)
    try:
        obj = _get_design_obj(app, object_type, name)
        obj.HasModule = True
    finally:
        _save_and_close(app, object_type, name)

    # 2. Clear VBE cache (module was just created)
    cache_key = f"{object_type}:{name}"
    _Session._cm_cache.pop(cache_key, None)
    _vbe_code_cache.pop(cache_key, None)

    # 3. Inject code via VBE (lazy import to avoid circular)
    from .vbe import _get_code_module
    cm = _get_code_module(app, object_type, name)
    total = cm.CountOfLines

    # Delete auto-generated content by Access (Option Compare Database, etc.)
    # to avoid duplicates with the VBA we are about to inject
    if total > 0:
        cm.DeleteLines(1, total)

    # Normalize line endings to \r\n (VBE requires it)
    vba_code = vba_code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    if not vba_code.endswith("\r\n"):
        vba_code += "\r\n"

    cm.InsertLines(1, vba_code)

    # Invalidate caches
    _vbe_code_cache.pop(cache_key, None)
    _Session._cm_cache.pop(cache_key, None)


def ac_import_text(db_path: str, object_type: str, object_name: str,
                   input_path: str) -> dict:
    """Imports a form/report/module/query/macro from text via LoadFromText.

    REPLACES the object if it already exists (deletes it first).
    Does NOT open in Design view — does not recalculate control positions.
    The file must be UTF-16 LE with BOM (the SaveAsText format).

    For forms/reports with CodeBehindForm section: automatically splits the VBA
    and injects via VBE (same as access_set_code) to avoid encoding errors.
    """
    if object_type not in ("form", "report", "module", "query", "macro"):
        raise ValueError("object_type must be form, report, module, query or macro")
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"File not found: {input_path}")

    # For forms/reports: read the content and detect CodeBehindForm
    if object_type in ("form", "report"):
        content, _enc = read_tmp(input_path)
        if "CodeBehindForm" in content or "CodeBehindReport" in content:
            log.info("ac_import_text: detectado CodeBehindForm en '%s', usando VBA split", object_name)
            # Split VBA from form text (avoids encoding errors in LoadFromText)
            form_code, vba_code = _split_code_behind(content)
            if vba_code:
                form_code = re.sub(r"^\s*HasModule\s*=.*$", "", form_code, flags=re.MULTILINE)

            app = _Session.connect(db_path)
            try:
                app.DoCmd.Close(AC_TYPE[object_type], object_name, AC_SAVE_NO)
            except Exception:
                pass
            try:
                app.DoCmd.DeleteObject(AC_TYPE[object_type], object_name)
            except Exception:
                pass

            fd, tmp = tempfile.mkstemp(suffix=".txt", prefix="access_mcp_imp_")
            os.close(fd)
            try:
                write_tmp(tmp, form_code, encoding="utf-16")
                try:
                    app.LoadFromText(AC_TYPE[object_type], object_name, tmp)
                except Exception as e:
                    raise RuntimeError(
                        f"LoadFromText failed for '{object_name}': {e}\n"
                        f"Check the form text syntax (without VBA). "
                        f"Use access_set_code for more error details."
                    ) from e
            finally:
                try:
                    os.unlink(tmp)
                except OSError:
                    pass

            if vba_code:
                _inject_vba_after_import(app, object_type, object_name, vba_code)

            invalidate_object_caches(object_type, object_name)
            return {"status": "imported", "object": object_name, "type": object_type,
                    "source": input_path, "vba_injected": bool(vba_code)}

    app = _Session.connect(db_path)
    # Close if open and delete existing
    try:
        app.DoCmd.Close(AC_TYPE[object_type], object_name, AC_SAVE_NO)
    except Exception:
        pass
    try:
        app.DoCmd.DeleteObject(AC_TYPE[object_type], object_name)
    except Exception:
        pass
    # Import from text
    try:
        app.LoadFromText(AC_TYPE[object_type], object_name, input_path)
    except Exception as e:
        raise RuntimeError(
            f"LoadFromText failed for '{object_name}': {e}\n"
            f"For forms/reports with VBA: use access_set_code instead of access_import_text."
        ) from e
    # Invalidate caches
    invalidate_object_caches(object_type, object_name)
    return {"status": "imported", "object": object_name, "type": object_type,
            "source": input_path}


# ---------------------------------------------------------------------------
# ac_set_control_props
# ---------------------------------------------------------------------------

def ac_set_control_props(
    db_path: str, object_type: str, object_name: str,
    control_name: str, props: dict
) -> dict:
    """
    Modifies properties of an existing control by opening the form/report in Design view.
    props: dict {property: value}. Values are automatically converted
    to int/bool when appropriate.
    Returns {"applied": [...], "errors": {...}}.
    """
    if object_type not in ("form", "report"):
        raise ValueError("Only 'form' or 'report'")

    app = _Session.connect(db_path)
    _open_in_design(app, object_type, object_name)
    applied: list[str] = []
    errors: dict[str, str] = {}
    try:
        obj  = _get_design_obj(app, object_type, object_name)
        ctrl = obj.Controls(control_name)
        for key, val in props.items():
            try:
                setattr(ctrl, key, coerce_prop(val))
                applied.append(key)
            except Exception as exc:
                errors[key] = str(exc)
    finally:
        _save_and_close(app, object_type, object_name)
        # Invalidate caches — form changed in Design view
        invalidate_object_caches(object_type, object_name)

    return {"applied": applied, "errors": errors}


# ---------------------------------------------------------------------------
# ac_set_form_property / ac_get_form_property
# ---------------------------------------------------------------------------

def ac_set_form_property(
    db_path: str, object_type: str, object_name: str, props: dict
) -> dict:
    """
    Sets properties at the form/report level by opening in Design view.
    Useful for changing RecordSource, Caption, DefaultView, HasModule, etc.
    props: dict {property: value}. Values are automatically converted to int/bool.
    Returns {"applied": [...], "errors": {...}}.
    """
    if object_type not in ("form", "report"):
        raise ValueError("Only 'form' or 'report'")

    app = _Session.connect(db_path)
    _open_in_design(app, object_type, object_name)
    applied: list[str] = []
    errors: dict[str, str] = {}
    try:
        obj = _get_design_obj(app, object_type, object_name)
        for key, val in props.items():
            try:
                setattr(obj, key, coerce_prop(val))
                applied.append(key)
            except Exception as exc:
                errors[key] = str(exc)
    finally:
        _save_and_close(app, object_type, object_name)
        # Invalidate caches — form properties changed
        invalidate_object_caches(object_type, object_name)

    return {"applied": applied, "errors": errors}


def ac_get_form_property(
    db_path: str, object_type: str, object_name: str,
    property_names: list[str] | None = None,
) -> dict:
    """
    Reads properties of a form/report by opening it in Design view.
    If property_names is None, reads all readable properties.
    Returns {"object": str, "type": str, "properties": {...}}.
    """
    if object_type not in ("form", "report"):
        raise ValueError("Only 'form' or 'report'")

    app = _Session.connect(db_path)
    _open_in_design(app, object_type, object_name)
    properties: dict = {}
    errors: dict[str, str] = {}
    try:
        obj = _get_design_obj(app, object_type, object_name)
        if property_names:
            for pname in property_names:
                try:
                    properties[pname] = serialize_value(obj.Properties(pname).Value)
                except Exception as exc:
                    errors[pname] = str(exc)
        else:
            # Read all readable properties
            for i in range(obj.Properties.Count):
                try:
                    p = obj.Properties(i)
                    properties[p.Name] = serialize_value(p.Value)
                except Exception:
                    pass  # Skip unreadable properties
    finally:
        _save_and_close(app, object_type, object_name)

    result: dict = {
        "object": object_name,
        "type": object_type,
        "properties": properties,
    }
    if errors:
        result["errors"] = errors
    return result


# ---------------------------------------------------------------------------
# ac_set_multiple_controls
# ---------------------------------------------------------------------------

def ac_set_multiple_controls(
    db_path: str, object_type: str, object_name: str,
    controls: list[dict],
) -> dict:
    """
    Modifies properties of multiple controls in a single operation.
    Opens the form/report in Design view only once.
    controls: [{"name": str, "props": {prop: val, ...}}, ...]
    Returns {"results": [{"name": str, "applied": [...], "errors": {...}}, ...]}.
    """
    if object_type not in ("form", "report"):
        raise ValueError("Only 'form' or 'report'")
    if not controls:
        return {"error": "No controls provided."}

    app = _Session.connect(db_path)
    _open_in_design(app, object_type, object_name)
    results: list[dict] = []
    try:
        obj = _get_design_obj(app, object_type, object_name)
        for ctrl_spec in controls:
            ctrl_name = ctrl_spec["name"]
            ctrl_props = ctrl_spec.get("props", {})
            applied: list[str] = []
            errors: dict[str, str] = {}
            try:
                ctrl = obj.Controls(ctrl_name)
                for key, val in ctrl_props.items():
                    try:
                        setattr(ctrl, key, coerce_prop(val))
                        applied.append(key)
                    except Exception as exc:
                        errors[key] = str(exc)
            except Exception as exc:
                errors["_control"] = f"Control '{ctrl_name}' not found: {exc}"
            entry: dict = {"name": ctrl_name, "applied": applied}
            if errors:
                entry["errors"] = errors
            results.append(entry)
    finally:
        _save_and_close(app, object_type, object_name)
        invalidate_object_caches(object_type, object_name)

    return {"results": results}
