"""
VBA compilation and linting.
"""

import re
import threading
from typing import Optional

from .core import _Session, _vbe_code_cache, log
from .constants import AC_CMD_COMPILE


# ---------------------------------------------------------------------------
# VBE error location after compile error
# ---------------------------------------------------------------------------

def _get_vbe_error_location(app) -> Optional[dict]:
    """After a compile error, VBE positions the cursor on the offending line.
    Try to read ActiveCodePane to extract module name, line number, and code.
    Returns dict with error location or None if unavailable.
    """
    try:
        pane = app.VBE.ActiveCodePane
        if pane is None:
            return None
        cm = pane.CodeModule
        module_name = cm.Parent.Name
        # GetSelection returns (StartLine, StartCol, EndLine, EndCol)
        start_line, start_col, end_line, end_col = pane.GetSelection()
        # Read a few lines around the error
        first = max(1, start_line - 2)
        last = min(cm.CountOfLines, start_line + 2)
        lines = []
        for i in range(first, last + 1):
            prefix = ">>> " if i == start_line else "    "
            lines.append(f"{prefix}{i}: {cm.Lines(i, 1)}")
        return {
            "module": module_name,
            "line": start_line,
            "code_context": "\n".join(lines),
        }
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Lint form modules
# ---------------------------------------------------------------------------

def _lint_form_modules(app) -> list:
    """Lint form modules: detect orphan event handlers and Me.X refs to missing controls.

    Returns list of warning strings. Empty if no issues found.
    Iterates all VBComponents of type 100 (Access form/report modules), opens each
    form in Design view to collect control names, then scans VBA code for:
      - Event handler subs whose ctrl prefix doesn't match any control
      - Me.X references to names that aren't controls or known Form properties
    """
    _FORM_PROPS = {
        "recordsource", "filter", "caption", "visible", "enabled", "dirty",
        "newrecord", "allowedits", "allowadditions", "allowdeletions", "requery",
        "refresh", "undo", "setfocus", "repaint", "recalc", "controls", "name",
        "tag", "filterstring", "orderbyon", "orderby", "dataentry", "cycle",
        "filteron", "openargs", "recordset", "bookmark", "currentrecord",
        "module", "hasmodule", "width", "painting", "popup", "modal",
        "borderstyle", "defaultview", "autocenter", "autoresize",
        "minmaxbuttons", "controlbox", "scrollbars", "navigbuttons",
        "gridx", "gridy", "picture", "picturetype", "layoutforprint",
        "fastlaserprinting", "allowlayoutview", "allowformview", "allowdataview",
        "splitformorientation", "whenclosed", "whenloaded", "whennothinghaschanged",
        "insidewidth", "insideheight", "currentview", "painted",
    }
    _event_re = re.compile(
        r"^\s*(?:Private\s+|Public\s+)?Sub\s+(\w+)_"
        r"(Click|BeforeUpdate|AfterUpdate|LostFocus|Change|GotFocus|KeyDown|"
        r"Enter|Exit|DblClick|MouseDown|MouseMove|KeyUp|KeyPress)\s*\(",
        re.IGNORECASE | re.MULTILINE,
    )
    _me_re = re.compile(r"\bMe\.(\w+)\b", re.IGNORECASE)

    warnings = []
    try:
        vbe = app.VBE
        proj = vbe.ActiveVBProject
        for comp in proj.VBComponents:
            if comp.Type != 100:  # vbext_ct_Document -- Access form/report modules
                continue
            form_name = comp.Name
            # Try to open as form in Design view to get control names
            ctrl_names = set()
            already_open = False
            try:
                try:
                    _ = app.Forms(form_name)
                    already_open = True
                except Exception:
                    pass
                if not already_open:
                    app.DoCmd.OpenForm(form_name, 1)  # acDesign=1
                form_obj = app.Forms(form_name)
                for ctrl in form_obj.Controls:
                    try:
                        ctrl_names.add(ctrl.Name.lower())
                    except Exception:
                        pass
                if not already_open:
                    app.DoCmd.Close(2, form_name, 2)  # acForm=2, acSaveNo=2
            except Exception:
                continue  # Not a form (maybe a report), can't open -- skip
            if not ctrl_names:
                continue
            # Get VBA code for this form module
            try:
                cm = comp.CodeModule
                if cm.CountOfLines == 0:
                    continue
                code = cm.Lines(1, cm.CountOfLines)
            except Exception:
                continue
            # Check orphan event handlers
            for m in _event_re.finditer(code):
                ctrl_part = m.group(1)
                if ctrl_part.lower().startswith("form"):
                    continue  # Form_Load, Form_Open, etc. -- valid
                if ctrl_part.lower() not in ctrl_names:
                    warnings.append(
                        f"{form_name}: event handler '{ctrl_part}_{m.group(2)}'"
                        f" -- control '{ctrl_part}' not found"
                    )
            # Check Me.X references (deduplicated per prop within this form)
            seen_me: set = set()
            for m in _me_re.finditer(code):
                prop = m.group(1)
                key = prop.lower()
                if key in seen_me:
                    continue
                seen_me.add(key)
                if key in _FORM_PROPS:
                    continue  # known Form property -- not a control
                if key not in ctrl_names:
                    warnings.append(
                        f"{form_name}: 'Me.{prop}' -- control '{prop}' not found"
                    )
    except Exception:
        pass  # VBE not accessible -- skip lint
    return warnings


# ---------------------------------------------------------------------------
# Compile VBA
# ---------------------------------------------------------------------------

def ac_compile_vba(db_path: str, timeout: Optional[int] = None) -> dict:
    """Attempts to compile VBA via RunCommand(126) + per-module verification.

    RunCommand(126) via COM has limitations (doesn't open the VBE, doesn't
    actually compile in many cases). As additional verification, iterates all
    standard modules and calls Application.Run on a function in each to force
    on-demand compilation. Reports any compilation failures.

    With timeout, a watchdog auto-dismisses error MsgBox.
    Returns dict with status + optional error_detail, error_location, dialog_screenshot.
    """
    # Lazy import to avoid circular dependency
    from .vba_exec import _dialog_watchdog

    app = _Session.connect(db_path)

    # 1. Intentar RunCommand(126) -- puede no hacer nada via COM, pero intentamos
    stop_event = None
    dialog_screenshots: list = []
    dismissed: list = []
    if timeout:
        _h = app.hWndAccessApp
        hwnd = int(_h() if callable(_h) else _h)
        stop_event = threading.Event()
        watchdog = threading.Thread(
            target=_dialog_watchdog,
            args=[hwnd, stop_event, dismissed, dialog_screenshots, 2.0],
            daemon=True,
        )
        watchdog.start()
    try:
        app.RunCommand(AC_CMD_COMPILE)
    except Exception as exc:
        err_loc = _get_vbe_error_location(app)
        result = {
            "status": "error",
            "error_detail": f"VBA compilation error: {exc}",
        }
        if err_loc:
            result["error_location"] = err_loc
        if dialog_screenshots:
            result["dialog_screenshot"] = dialog_screenshots[0]
        return result
    finally:
        if stop_event:
            stop_event.set()

    _vbe_code_cache.clear()
    _Session._cm_cache.clear()

    if dismissed:
        result = {
            "status": "error",
            "error_detail": "VBA compilation error -- error MsgBox auto-dismissed.",
        }
        err_loc = _get_vbe_error_location(app)
        if err_loc:
            result["error_location"] = err_loc
        if dialog_screenshots:
            result["dialog_screenshot"] = dialog_screenshots[0]
        return result

    # 2. Verificar IsCompiled -- si True, pasar directamente al lint
    try:
        if bool(app.IsCompiled):
            warnings = _lint_form_modules(app)
            result = {"status": "compiled"}
            if warnings:
                result["warnings"] = warnings
            return result
    except Exception:
        pass

    # 3. IsCompiled=False: RunCommand didn't compile. Verify per-module using
    #    Application.Run to force on-demand compilation of each standard module.
    errors = []
    try:
        vbe = app.VBE
        proj = vbe.ActiveVBProject
        for comp in proj.VBComponents:
            if comp.Type != 1:  # standard modules only (vbext_ct_StdModule)
                continue
            cm = comp.CodeModule
            # Find the first Public Function to call
            func_name = None
            for line_num in range(cm.CountOfDeclarationLines + 1, cm.CountOfLines + 1):
                try:
                    pname = cm.ProcOfLine(line_num, 0)  # vbext_pk_Proc=0
                    if pname:
                        # Verify it's a Function (not Sub) to use Run
                        proc_line = cm.ProcBodyLine(pname, 0)
                        proc_text = cm.Lines(proc_line, 1).strip().lower()
                        if proc_text.startswith("public function"):
                            func_name = pname
                            break
                except Exception:
                    continue
            if not func_name:
                continue
            # Try Application.Run -- forces compilation of the entire module
            try:
                app.Run(f"{comp.Name}.{func_name}")
            except Exception as exc:
                err_str = str(exc).lower()
                if "compile" in err_str or "expected" in err_str or "type mismatch" in err_str or "byref" in err_str:
                    errors.append(f"{comp.Name}.{func_name}: {exc}")
                # Other runtime errors (division by zero, etc.) are OK -- the module compiled
    except Exception as exc:
        # If we can't access VBE, report warning
        return {
            "status": "compiled",
            "note": f"IsCompiled=False, per-module verification failed: {exc}"
        }

    if errors:
        return {
            "status": "error",
            "error_detail": "Compilation errors detected via Application.Run:\n" + "\n".join(errors),
        }

    # 4. Lint form/report modules: orphan event handlers + Me.X refs to missing controls
    warnings = _lint_form_modules(app)
    result = {"status": "compiled"}
    if warnings:
        result["warnings"] = warnings
    return result
