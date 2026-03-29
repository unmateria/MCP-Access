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

def _verify_module_structure(app) -> list:
    """Verify structural integrity of ALL VBA modules (standard + form/report).

    RunCommand(acCmdCompileAndSaveAllModules) via COM may not detect errors in
    form/report modules even with VBE open.  This function checks that no
    executable code exists outside Sub/Function/Property/Type/Enum blocks.

    Catches the specific bug pattern: Sub/Function header accidentally deleted,
    leaving orphan code after End Sub that VBA silently absorbs into the
    previous procedure.

    Returns list of error strings. Empty if all OK.
    """
    # Regex for valid module-level statements (outside any proc)
    _MODULE_LEVEL = re.compile(
        r"(?:Option\s|Dim\s|Private\s|Public\s|Global\s|Const\s|Declare\s|"
        r"#If|#ElseIf|#Else|#End\s|#Const\s|Attribute\s|Implements\s|Event\s|"
        r"Friend\s|Static\s|Sub\s|Function\s|Property\s|Type\s|Enum\s|DefInt\s|"
        r"DefLng\s|DefSng\s|DefDbl\s|DefCur\s|DefStr\s|DefBool\s|DefDate\s|"
        r"DefVar\s|DefObj\s|DefByte\s)",
        re.IGNORECASE,
    )
    _PROC_START = re.compile(
        r"(?:Private\s+|Public\s+|Friend\s+)?(?:Static\s+)?"
        r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s",
        re.IGNORECASE,
    )
    _BLOCK_START = re.compile(
        r"(?:Private\s+|Public\s+)?(?:Type|Enum)\s", re.IGNORECASE
    )
    _BLOCK_END = re.compile(r"End\s+(?:Type|Enum)", re.IGNORECASE)
    _PROC_END = re.compile(
        r"End\s+(?:Sub|Function|Property)", re.IGNORECASE
    )

    errors = []
    try:
        vbe = app.VBE
        proj = vbe.ActiveVBProject
        for comp in proj.VBComponents:
            if comp.Type not in (1, 100):  # standard modules + form/report
                continue
            cm = comp.CodeModule
            total = cm.CountOfLines
            if total == 0:
                continue
            code = cm.Lines(1, total)

            in_proc = False
            in_block = False  # Type / Enum
            continuation = False

            for i, line in enumerate(code.split("\n"), 1):
                stripped = line.strip()

                # Line continuation from previous line
                if continuation:
                    continuation = stripped.endswith(" _")
                    continue
                if stripped.endswith(" _"):
                    continuation = True
                    # Still process the first line of the continuation

                # Skip blank / comment
                if not stripped or stripped.startswith("'"):
                    continue

                # Type/Enum blocks
                if not in_proc and _BLOCK_START.match(stripped):
                    in_block = True
                    continue
                if in_block:
                    if _BLOCK_END.match(stripped):
                        in_block = False
                    continue

                # Proc start/end
                if _PROC_START.match(stripped):
                    in_proc = True
                    continue
                if _PROC_END.match(stripped):
                    in_proc = False
                    continue

                # Inside a proc: anything goes
                if in_proc:
                    continue

                # Module level: only declarations/directives are valid
                if not _MODULE_LEVEL.match(stripped):
                    errors.append(
                        f"{comp.Name} line {i}: code outside Sub/Function: "
                        f"{stripped[:80]}"
                    )
                    break  # one error per module is enough

    except Exception:
        pass  # VBE not accessible -- skip
    return errors


def ac_compile_vba(db_path: str, timeout: Optional[int] = None) -> dict:
    """Compile VBA with VBE open + structural verification.

    Opens the VBE MainWindow so RunCommand(126) behaves like clicking
    Debug > Compile.  Additionally runs _verify_module_structure() to
    catch orphan code outside Sub/Function blocks (a pattern that
    RunCommand sometimes misses for form/report modules).

    With timeout, a watchdog auto-dismisses error MsgBox.
    Returns dict with status + optional error_detail, error_location, dialog_screenshot.
    """
    # Lazy import to avoid circular dependency
    from .vba_exec import _dialog_watchdog

    app = _Session.connect(db_path)

    # 0. Force project to "not compiled" state.
    #    VBE edits via COM don't always invalidate IsCompiled, so RunCommand on
    #    an already-compiled project can be a no-op.
    vbe_was_visible = False
    try:
        vbe_was_visible = bool(app.VBE.MainWindow.Visible)
    except Exception:
        pass
    try:
        _proj = app.VBE.ActiveVBProject
        for _comp in _proj.VBComponents:
            if _comp.Type == 1 and _comp.CodeModule.CountOfLines > 0:
                _cm = _comp.CodeModule
                _cm.InsertLines(_cm.CountOfLines + 1, "' _compile_dirty_check")
                _cm.DeleteLines(_cm.CountOfLines, 1)
                break
    except Exception:
        pass

    # 1. Open VBE so RunCommand(126) compiles ALL modules (including form/report).
    #    Without VBE visible, RunCommand often silently skips form/report modules.
    try:
        app.VBE.MainWindow.Visible = True
    except Exception:
        pass

    # 2. RunCommand(126) = acCmdCompileAndSaveAllModules
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
        # Restore VBE visibility
        if not vbe_was_visible:
            try:
                app.VBE.MainWindow.Visible = False
            except Exception:
                pass

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

    # 3. Structural verification: catch orphan code outside Sub/Function
    #    (RunCommand can still miss this for form/report modules)
    struct_errors = _verify_module_structure(app)
    if struct_errors:
        return {
            "status": "error",
            "error_detail": "Structural errors in VBA modules:\n" + "\n".join(struct_errors),
        }

    # 4. Lint form/report modules: orphan event handlers + Me.X refs
    warnings = _lint_form_modules(app)
    result = {"status": "compiled"}
    if warnings:
        result["warnings"] = warnings
    return result
