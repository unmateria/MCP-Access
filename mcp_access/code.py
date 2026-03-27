"""
Object management: list, get, set, delete objects, create form, export structure.
"""

import os
import re
import tempfile
from pathlib import Path
from typing import Any

from .core import (
    AC_TYPE, _Session, _vbe_code_cache, _parsed_controls_cache, log,
    invalidate_all_caches, invalidate_object_caches,
)
from .constants import BINARY_SECTIONS, AC_FORM, AC_SAVE_NO
from .helpers import read_tmp, write_tmp, strip_binary_sections, restore_binary_sections


# ---------------------------------------------------------------------------
# Design-view helpers (used by _inject_vba_after_import)
# ---------------------------------------------------------------------------
# These are small private helpers also used by controls module.  Duplicated
# here to keep the module self-contained; the canonical copy lives in helpers
# once that module is extended.

_AC_DESIGN   = 1   # acDesign / acViewDesign
_AC_SAVE_YES = 1   # acSaveYes
_AC_REPORT   = 3   # acReport


def _open_in_design(app: Any, object_type: str, object_name: str) -> None:
    """Opens a form/report in Design view."""
    try:
        if object_type == "form":
            app.DoCmd.OpenForm(object_name, _AC_DESIGN)
        else:
            app.DoCmd.OpenReport(object_name, _AC_DESIGN)
    except Exception as exc:
        raise RuntimeError(
            f"Could not open '{object_name}' in Design view. "
            f"If it is open in Normal view, close it first.\nError: {exc}"
        )


def _save_and_close(app: Any, object_type: str, object_name: str) -> None:
    """Saves and closes a form/report open in Design view."""
    ac_type = AC_FORM if object_type == "form" else _AC_REPORT
    try:
        app.DoCmd.Close(ac_type, object_name, _AC_SAVE_YES)
    except Exception as exc:
        log.warning("Error closing '%s': %s", object_name, exc)


def _get_design_obj(app: Any, object_type: str, object_name: str) -> Any:
    """Returns the Form or Report object open in Design view."""
    return app.Forms(object_name) if object_type == "form" else app.Reports(object_name)


# ---------------------------------------------------------------------------
# List objects
# ---------------------------------------------------------------------------

def ac_list_objects(db_path: str, object_type: str = "all") -> dict:
    """Returns a dict {type: [names...]} with the database objects."""
    app = _Session.connect(db_path)

    # CurrentData  -> data objects (tables, queries)
    # CurrentProject -> code objects (forms, reports, modules, macros)
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
    """Deletes an Access object (module, form, report, query, macro) via DoCmd.DeleteObject."""
    if object_type not in AC_TYPE:
        raise ValueError(
            f"Invalid object_type '{object_type}'. Valid: {list(AC_TYPE)}"
        )
    if not confirm:
        raise ValueError(
            "Destructive operation: confirm=true is required to delete an object."
        )
    app = _Session.connect(db_path)
    try:
        app.DoCmd.DeleteObject(AC_TYPE[object_type], object_name)
    except Exception as exc:
        raise RuntimeError(
            f"Error deleting {object_type} '{object_name}': {exc}"
        )
    finally:
        invalidate_all_caches()
    return {
        "action": "deleted",
        "object_type": object_type,
        "object_name": object_name,
    }


# ---------------------------------------------------------------------------
# Get code (export)
# ---------------------------------------------------------------------------

def ac_get_code(db_path: str, object_type: str, name: str) -> str:
    """
    Exports an Access object to text and returns the content.
    For forms and reports, strips binary sections (PrtMip, PrtDevMode...)
    that are irrelevant for editing VBA/controls and represent 95% of the size.
    ac_set_code restores them automatically before importing.
    """
    if object_type not in AC_TYPE:
        raise ValueError(
            f"Invalid object_type '{object_type}'. Valid: {list(AC_TYPE)}"
        )
    app = _Session.connect(db_path)

    fd, tmp = tempfile.mkstemp(suffix=".txt", prefix="access_mcp_")
    os.close(fd)
    try:
        app.SaveAsText(AC_TYPE[object_type], name, tmp)
        text, _enc = read_tmp(tmp)
        if object_type in ("form", "report"):
            text = strip_binary_sections(text)
        return text
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Split CodeBehind from form/report text
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Inject VBA after import
# ---------------------------------------------------------------------------

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

    # 3. Inject code via VBE (lazy import from .vbe)
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


# ---------------------------------------------------------------------------
# Set code (import)
# ---------------------------------------------------------------------------

def ac_set_code(db_path: str, object_type: str, name: str, code: str) -> str:
    """
    Imports text as an Access object definition (creates or overwrites).
    For forms and reports, automatically re-injects binary sections
    (PrtMip, PrtDevMode...) from the current export, so the caller doesn't need
    to include them in the code they send.

    If the code contains a CodeBehindForm/CodeBehindReport section, it is automatically
    split: first the form/report is imported without VBA, then the VBA code is injected
    via VBE (avoiding encoding issues with LoadFromText).
    """
    if object_type not in AC_TYPE:
        raise ValueError(
            f"Invalid object_type '{object_type}'. Valid: {list(AC_TYPE)}"
        )
    app = _Session.connect(db_path)

    # Split CodeBehindForm/CodeBehindReport if present
    vba_code = ""
    if object_type in ("form", "report"):
        code, vba_code = _split_code_behind(code)
        # Remove HasModule from form text — it will be activated when injecting VBA
        if vba_code:
            code = re.sub(r"^\s*HasModule\s*=.*$", "", code, flags=re.MULTILINE)

    # If the code doesn't contain binary sections (returned by ac_get_code
    # with filtering active), restore them from the current form/report.
    if object_type in ("form", "report") and not any(
        s in code for s in BINARY_SECTIONS
    ):
        log.info("ac_set_code: restoring binary sections for '%s'", name)
        code = restore_binary_sections(app, object_type, name, code)

    # Backup existing object in case import fails
    backup_tmp = None
    if object_type in ("form", "report", "module"):
        try:
            fd_bk, backup_tmp = tempfile.mkstemp(suffix=".txt", prefix="access_mcp_bk_")
            os.close(fd_bk)
            app.SaveAsText(AC_TYPE[object_type], name, backup_tmp)
        except Exception:
            # Doesn't exist yet — no backup needed
            if backup_tmp:
                try:
                    os.unlink(backup_tmp)
                except OSError:
                    pass
            backup_tmp = None

    fd, tmp = tempfile.mkstemp(suffix=".txt", prefix="access_mcp_")
    os.close(fd)
    try:
        # VBA modules (.bas) expect ANSI/cp1252; forms/reports/queries/macros expect UTF-16LE with BOM
        enc = "cp1252" if object_type == "module" else "utf-16"
        write_tmp(tmp, code, encoding=enc)
        try:
            app.LoadFromText(AC_TYPE[object_type], name, tmp)
        except Exception as import_exc:
            # Restaurar backup si existe
            if backup_tmp and os.path.exists(backup_tmp):
                log.warning("ac_set_code: import failed, restoring backup for '%s'", name)
                try:
                    app.LoadFromText(AC_TYPE[object_type], name, backup_tmp)
                except Exception:
                    log.error("ac_set_code: could not restore backup for '%s'", name)
            raise import_exc

        # Invalidate caches for this object (code and controls changed)
        invalidate_object_caches(object_type, name)

        # Inject VBA if there was CodeBehindForm
        vba_msg = ""
        if vba_code:
            _inject_vba_after_import(app, object_type, name, vba_code)
            vba_msg = " (with VBA injected via VBE)"

        return f"OK: '{name}' ({object_type}) imported successfully into {db_path}{vba_msg}"
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


# ---------------------------------------------------------------------------
# Create form
# ---------------------------------------------------------------------------

def ac_create_form(db_path: str, form_name: str, has_header: bool = False) -> dict:
    """Creates a new form avoiding the 'Save As' MsgBox that blocks COM.

    CreateForm() generates a form with an auto name (Form1, Form2...).
    DoCmd.Save saves with that name (no dialog).
    DoCmd.Close with acSaveNo closes (already saved, no dialog).
    DoCmd.Rename renames to the desired name.
    """
    app = _Session.connect(db_path)
    auto_name = None
    try:
        form = app.CreateForm()
        auto_name = form.Name  # e.g. "Form1"

        if has_header:
            app.RunCommand(36)  # acCmdFormHdrFtr — toggle header/footer

        # Save with auto-name — no dialog (DoCmd.Save uses current name)
        app.DoCmd.Save(AC_FORM, auto_name)
        # Close without prompt (already saved)
        app.DoCmd.Close(AC_FORM, auto_name, AC_SAVE_NO)

        # Rename to desired name
        if auto_name != form_name:
            app.DoCmd.Rename(form_name, AC_FORM, auto_name)

        return {"name": form_name, "created_from": auto_name, "has_header": has_header}
    except Exception as exc:
        if auto_name:
            try:
                app.DoCmd.Close(AC_FORM, auto_name, AC_SAVE_NO)
            except Exception:
                pass
            try:
                app.DoCmd.DeleteObject(AC_FORM, auto_name)
            except Exception:
                pass
        raise RuntimeError(f"Error creating form '{form_name}': {exc}")
    finally:
        invalidate_all_caches()


# ---------------------------------------------------------------------------
# Export structure
# ---------------------------------------------------------------------------

def ac_export_structure(db_path: str, output_path: str | None = None) -> str:
    """
    Generates a Markdown file with the complete database structure:
    VBA modules with their function signatures, forms, reports and queries.
    """
    from datetime import datetime

    if output_path is None:
        output_path = str(Path(db_path).parent / "db_structure.md")

    objects = ac_list_objects(db_path, "all")
    modules  = objects.get("module",  [])
    forms    = objects.get("form",    [])
    reports  = objects.get("report",  [])
    queries  = objects.get("query",   [])
    macros   = objects.get("macro",   [])

    lines: list[str] = []
    lines.append(f"# Structure of `{Path(db_path).name}`")
    lines.append(f"\n**Path**: `{db_path}`  ")
    lines.append(f"**Generated**: {datetime.now().strftime('%Y-%m-%d %H:%M')}  ")
    lines.append(
        f"**Summary**: {len(modules)} modules · {len(forms)} forms · "
        f"{len(reports)} reports · {len(queries)} queries · {len(macros)} macros\n"
    )

    # -- VBA Modules with signatures --
    # Read modules via VBE (no SaveAsText/disk) and warming up the code cache
    # Lazy imports from .vbe
    from .vbe import _get_code_module, _cm_all_code

    app = _Session.connect(db_path)
    lines.append(f"## VBA Modules ({len(modules)})\n")
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
                lines.append("  *(no public functions/subs)*")
        except Exception as exc:
            lines.append(f"  *(error reading: {exc})*")
        lines.append("")

    # -- Forms --
    lines.append(f"## Forms ({len(forms)})\n")
    if forms:
        for name in forms:
            lines.append(f"- `{name}`")
    else:
        lines.append("*(none)*")
    lines.append("")

    # -- Reports --
    lines.append(f"## Reports ({len(reports)})\n")
    if reports:
        for name in reports:
            lines.append(f"- `{name}`")
    else:
        lines.append("*(none)*")
    lines.append("")

    # -- Queries --
    lines.append(f"## Queries ({len(queries)})\n")
    if queries:
        for name in queries:
            lines.append(f"- `{name}`")
    else:
        lines.append("*(none)*")
    lines.append("")

    # -- Macros --
    if macros:
        lines.append(f"## Macros ({len(macros)})\n")
        for name in macros:
            lines.append(f"- `{name}`")
        lines.append("")

    content = "\n".join(lines)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)

    return f"[Saved to `{output_path}`]\n\n{content}"
