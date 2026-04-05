"""
VBA compilation and linting.
"""

import re
import threading
import time
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


def _find_block_mismatches(app) -> list:
    """Parse ALL VBA modules for mismatched block structures.

    Catches: If/End If, For/Next, Do/Loop, While/Wend,
    Select Case/End Select, With/End With.

    Returns list of error dicts: {module, line, error}.
    """
    # Patterns for block openers (multiline only — single-line If is excluded)
    _LINE_CONT = re.compile(r"\s+_$")

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
            lines = code.split("\n")
            _check_blocks_in_module(comp.Name, lines, errors)
    except Exception as exc:
        log.warning("_find_block_mismatches failed: %s", exc)
    return errors


def _check_blocks_in_module(module_name: str, lines: list, errors: list):
    """Check block structure in a single module's lines."""
    # Stack of (block_type, line_number)
    stack: list = []
    in_proc = False
    i = 0
    while i < len(lines):
        raw = lines[i]
        # Join continuation lines
        full_line = raw
        while full_line.rstrip().endswith(" _") and i + 1 < len(lines):
            i += 1
            full_line = full_line.rstrip()[:-1] + " " + lines[i].strip()
        stripped = full_line.strip()
        line_num = i + 1  # 1-based

        # Skip blank / comment
        if not stripped or stripped.startswith("'"):
            i += 1
            continue

        upper = stripped.upper()

        # Conditional compilation directives: #If / #ElseIf / #Else / #End If
        # These are separate from runtime If/End If — handle first.
        if upper.startswith("#"):
            if re.match(r"#IF\s+", upper):
                stack.append(("#If", line_num))
            elif re.match(r"#END\s+IF", upper):
                if stack and stack[-1][0] == "#If":
                    stack.pop()
            # #ElseIf, #Else — no stack change
            i += 1
            continue

        # Track proc boundaries to reset stack per-proc
        if re.match(r"(?:PRIVATE\s+|PUBLIC\s+|FRIEND\s+)?(?:STATIC\s+)?(?:SUB|FUNCTION|PROPERTY\s+(?:GET|LET|SET))\s",
                     upper):
            in_proc = True
            stack = []
            i += 1
            continue

        if re.match(r"END\s+(?:SUB|FUNCTION|PROPERTY)", upper):
            # Check for unclosed blocks at end of proc
            if stack:
                blk_type, blk_line = stack[0]  # report the first unclosed
                errors.append({
                    "module": module_name,
                    "line": blk_line,
                    "error": f"Block {blk_type} without End {blk_type} (unclosed at End Sub/Function, line {line_num})",
                })
            in_proc = False
            stack = []
            i += 1
            continue

        if not in_proc:
            i += 1
            continue

        # --- Block openers ---

        # Multiline If: "If ... Then" where Then is at end of line (not single-line If)
        # Single-line If has executable code after Then on the same line.
        # Accept optional trailing comment: If x Then  ' comment
        m_if = re.match(r"(?:#?)IF\s+.+\sTHEN\s*(?:'.*)?$", upper)
        if m_if:
            stack.append(("If", line_num))
            i += 1
            continue

        # ElseIf — doesn't change stack depth, just validate there's an If open
        if re.match(r"ELSEIF\s+", upper):
            i += 1
            continue

        # Else — same
        if upper == "ELSE" or upper.startswith("ELSE ") or upper == "ELSE:":
            i += 1
            continue

        # For Each ... / For ...
        # Single-line: "For Each x In y: doSomething: Next" — has Next on same line
        if re.match(r"FOR\s+(?:EACH\s+)?\w+", upper):
            if not re.search(r":\s*NEXT\b", upper):
                stack.append(("For", line_num))
            i += 1
            continue

        # Do ...
        # Single-line: "Do While x: something: Loop" — has Loop on same line
        if upper == "DO" or re.match(r"DO\s+(?:WHILE|UNTIL)\s", upper):
            if not re.search(r":\s*LOOP\b", upper):
                stack.append(("Do", line_num))
            i += 1
            continue

        # While ...
        if re.match(r"WHILE\s+", upper) and not re.match(r"WHILE\s+WEND", upper):
            stack.append(("While", line_num))
            i += 1
            continue

        # Select Case ...
        if re.match(r"SELECT\s+CASE\s", upper):
            stack.append(("Select", line_num))
            i += 1
            continue

        # With ...
        if re.match(r"WITH\s+", upper):
            stack.append(("With", line_num))
            i += 1
            continue

        # --- Block closers ---

        if re.match(r"END\s+IF", upper):
            if stack and stack[-1][0] == "If":
                stack.pop()
            elif not stack:
                errors.append({
                    "module": module_name,
                    "line": line_num,
                    "error": "End If without matching If",
                })
            else:
                # Mismatched — the top of stack is not If
                blk_type, blk_line = stack[-1]
                errors.append({
                    "module": module_name,
                    "line": line_num,
                    "error": f"End If but expected End {blk_type} (opened at line {blk_line})",
                })
            i += 1
            continue

        if upper.startswith("NEXT"):
            if stack and stack[-1][0] == "For":
                stack.pop()
            i += 1
            continue

        if upper == "LOOP" or re.match(r"LOOP\s+(?:WHILE|UNTIL)\s", upper):
            if stack and stack[-1][0] == "Do":
                stack.pop()
            i += 1
            continue

        if upper == "WEND":
            if stack and stack[-1][0] == "While":
                stack.pop()
            i += 1
            continue

        if re.match(r"END\s+SELECT", upper):
            if stack and stack[-1][0] == "Select":
                stack.pop()
            i += 1
            continue

        if re.match(r"END\s+WITH", upper):
            if stack and stack[-1][0] == "With":
                stack.pop()
            i += 1
            continue

        i += 1


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

    # 2b. Verify IsCompiled — RunCommand(126) can silently fail for form/report
    #     modules without raising an exception or showing a MsgBox.
    #     If not compiled, analyze VBA code for block mismatches.
    try:
        if not app.IsCompiled:
            log.warning("IsCompiled=False after RunCommand — analyzing VBA blocks...")
            block_errors = _find_block_mismatches(app)
            if block_errors:
                detail_lines = []
                for e in block_errors[:10]:  # limit to 10
                    detail_lines.append(f"  {e['module']} line {e['line']}: {e['error']}")
                return {
                    "status": "error",
                    "error_detail": "VBA compile errors detected (IsCompiled=False):\n"
                                   + "\n".join(detail_lines),
                    "errors": block_errors[:10],
                }
            # No block mismatches found but still not compiled — generic error
            return {
                "status": "error",
                "error_detail": "VBA project is NOT compiled after RunCommand. "
                                "No block mismatches found — the error may be a "
                                "missing reference, undeclared variable, or type mismatch. "
                                "Open VBE and do Debug > Compile manually to see the exact error.",
            }
    except Exception:
        pass  # IsCompiled not available — continue with structural checks

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
