"""
VBA execution: run_vba, run_macro, eval_vba, and dialog helpers.
"""

import ctypes
import json
import os
import tempfile
import threading
import time
from typing import Any, Optional

from .core import _Session, log


# ---------------------------------------------------------------------------
# Application.Run via InvokeTypes
# ---------------------------------------------------------------------------

def _invoke_app_run(app, procedure: str, call_args: list):
    """Call Application.Run via InvokeTypes -- proper COM optional-param protocol.

    pywin32's late-bound Dispatch uses Invoke() which passes only the provided
    args.  Access.Application.Run has 31 params (1 required + 30 optional) and
    its COM server rejects calls missing VT_ERROR markers for the optional params
    with DISP_E_BADPARAMCOUNT (-2147352562).

    InvokeTypes converts pythoncom.Missing -> VT_ERROR/DISP_E_PARAMNOTFOUND,
    matching what early-bound wrappers (EnsureDispatch / MakeDispatchFuncMethod)
    generate.
    """
    import pythoncom

    dispid = app._oleobj_.GetIDsOfNames(0, "Run")

    # Application.Run signature:
    #   Function Run(Procedure As String, [Arg1], ..., [Arg30]) As Variant
    # Arg types: (VT, PARAMFLAGS)
    #   Procedure:   VT_BSTR(8),    PARAMFLAG_FIN(1)
    #   Arg1..Arg30: VT_VARIANT(12), PARAMFLAG_FIN|PARAMFLAG_FOPT(17)
    arg_types = tuple([(8, 1)] + [(12, 17)] * 30)

    # Fill: procedure + user args + padding with Missing
    n = len(call_args)
    all_args = [procedure] + list(call_args) + [pythoncom.Missing] * (30 - n)

    return app._oleobj_.InvokeTypes(
        dispid,
        0,                           # LCID
        pythoncom.DISPATCH_METHOD,   # wFlags
        (12, 0),                     # return type: VT_VARIANT
        arg_types,
        *all_args,
    )


# ---------------------------------------------------------------------------
# Dialog dismissal helpers
# ---------------------------------------------------------------------------

def _dismiss_access_dialogs(hwnd_access: int, screenshot_holder: Optional[list] = None) -> bool:
    """Dismiss modal dialogs owned by the Access process.

    Uses process-ID matching (not just owner hwnd) to also catch VBE-owned
    dialogs.  Tries clicking End/OK/Finalizar buttons first (more reliable
    for VBA runtime-error dialogs), then falls back to WM_CLOSE.
    Returns True if any dialog was found and dismissed.
    """
    import win32gui
    import win32process

    try:
        _, access_pid = win32process.GetWindowThreadProcessId(hwnd_access)
    except Exception:
        return False

    found = []
    def _cb(hwnd, found):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid != access_pid:
                return True
            cls = win32gui.GetClassName(hwnd)
            if cls == '#32770':
                found.append(hwnd)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(_cb, found)
    except Exception:
        return False

    if not found:
        return False

    # Capture screenshot of first dialog before dismissing
    if screenshot_holder is not None:
        try:
            # Lazy import to avoid circular dependency
            from .ui import _capture_window
            img, _, _ = _capture_window(found[0], max_width=800)
            path = os.path.join(tempfile.gettempdir(),
                                f"access_dialog_{int(time.time())}.png")
            img.save(path)
            screenshot_holder.append(path)
        except Exception:
            pass  # screenshot is best-effort

    for dlg in found:
        _try_click_button(dlg)
        # Fallback: WM_CLOSE
        try:
            if win32gui.IsWindow(dlg):
                win32gui.PostMessage(dlg, 0x0010, 0, 0)  # WM_CLOSE
        except Exception:
            pass

    return True


def _try_click_button(dialog_hwnd: int):
    """Find and click End/OK/Finalizar/Aceptar button in a dialog."""
    import win32gui

    target_texts = {"end", "finalizar", "ok", "aceptar",
                    "&end", "&finalizar", "&ok", "&aceptar"}
    found_btn = [None]

    def _cb(hwnd, _):
        try:
            if win32gui.GetClassName(hwnd) == 'Button':
                text = win32gui.GetWindowText(hwnd).lower().strip()
                if text in target_texts or text.lstrip('&') in target_texts:
                    found_btn[0] = hwnd
                    return False
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(dialog_hwnd, _cb, None)
    except Exception:
        pass

    if found_btn[0]:
        try:
            win32gui.PostMessage(found_btn[0], 0x00F5, 0, 0)  # BM_CLICK
        except Exception:
            pass


def _dialog_watchdog(hwnd_access: int, stop_event: threading.Event,
                     dismissed: list, screenshot_holder: list,
                     interval: float = 2.0):
    """Poll for Access dialogs every *interval* seconds and dismiss them."""
    while not stop_event.is_set():
        if _dismiss_access_dialogs(hwnd_access,
                                   screenshot_holder if not dismissed else None):
            dismissed.append(True)
        stop_event.wait(interval)


# ---------------------------------------------------------------------------
# Run macro
# ---------------------------------------------------------------------------

def ac_run_macro(db_path: str, macro_name: str) -> dict:
    """Runs an Access macro."""
    app = _Session.connect(db_path)
    try:
        app.DoCmd.RunMacro(macro_name)
    except Exception as exc:
        raise RuntimeError(f"Error running macro '{macro_name}': {exc}")
    return {"macro_name": macro_name, "status": "executed"}


# ---------------------------------------------------------------------------
# Run VBA procedure
# ---------------------------------------------------------------------------

def ac_run_vba(
    db_path: str, procedure: str, args: Optional[list] = None,
    timeout: Optional[int] = None,
) -> dict:
    """Runs a VBA Sub/Function via Application.Run (or COM Forms() for form modules).

    Supports 3 syntaxes:
    - 'MyModule.MySub' or 'MySub' -> Application.Run (standard modules)
    - 'Forms.FormName.Method' -> COM Forms() access (form modules, form must be open)

    If the procedure shows MsgBox/InputBox and timeout is passed, the dialog is
    auto-dismissed after timeout seconds and a timeout error is returned.
    """
    app = _Session.connect(db_path)
    call_args = args or []
    if len(call_args) > 30:
        raise ValueError("Application.Run supports max 30 arguments.")

    # Forms.FormName.Method -> direct COM access (form modules)
    if "." in procedure:
        parts = procedure.split(".", 2)
        if parts[0] == "Forms" and len(parts) == 3:
            form_name, method_name = parts[1], parts[2]
            try:
                form = app.Forms(form_name)
                if call_args:
                    result = getattr(form, method_name)(*call_args)
                else:
                    # Try method call first, fall back to property read
                    attr = getattr(form, method_name)
                    try:
                        result = attr() if callable(attr) else attr
                    except (TypeError, AttributeError):
                        result = attr
            except Exception as exc:
                raise RuntimeError(
                    f"Error calling Forms('{form_name}').{method_name}: {exc}. "
                    f"Make sure the form is open."
                )
            if result is not None:
                try:
                    json.dumps(result)
                except (TypeError, ValueError):
                    result = str(result)
            return {"procedure": procedure, "result": result, "status": "executed"}

    # Standard Application.Run with optional watchdog timeout (polling every 2s)
    stop_event = None
    dismissed: list = []
    dialog_screenshots: list = []
    if timeout:
        # Capture hwnd on main thread (COM is apartment-threaded)
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
        result = _invoke_app_run(app, procedure, call_args)
    except Exception as exc:
        if dismissed:
            detail = f"'{procedure}' -- VBA runtime error (dialog auto-dismissed)."
            if dialog_screenshots:
                detail += f" Screenshot: {dialog_screenshots[0]}"
            raise RuntimeError(detail)
        raise RuntimeError(f"Error running '{procedure}': {exc}")
    finally:
        if stop_event:
            stop_event.set()
    # COM may return non-serializable types; convert to str if needed
    if result is not None:
        try:
            json.dumps(result)
        except (TypeError, ValueError):
            result = str(result)
    return {"procedure": procedure, "result": result, "status": "executed"}


# ---------------------------------------------------------------------------
# Eval VBA expression
# ---------------------------------------------------------------------------

def _invoke_app_eval(app, expression: str):
    """Call Application.Eval via InvokeTypes -- same pattern as _invoke_app_run."""
    import pythoncom
    dispid = app._oleobj_.GetIDsOfNames(0, "Eval")
    # Eval(StringExpr As String) As Variant -- 1 required param
    return app._oleobj_.InvokeTypes(
        dispid, 0, pythoncom.DISPATCH_METHOD,
        (12, 0),       # return: VT_VARIANT
        ((8, 1),),     # 1 param: VT_BSTR, PARAMFLAG_FIN
        expression,
    )


def ac_eval_vba(db_path: str, expression: str) -> dict:
    """Evaluates a VBA/Access expression via Application.Eval.

    Allows:
    - Calling form module functions: Eval("Forms!frmX.MyFunction()")
    - Reading properties of open forms: Eval("Forms!frmX.MARGEN_SEG")
    - Domain functions: Eval("DLookup(""Company"",""Sales"",""numc=1"")")
    - VBA built-in functions: Eval("Date()")
    - Functions only (not Subs). Form must be open.
    """
    app = _Session.connect(db_path)
    try:
        result = _invoke_app_eval(app, expression)
    except Exception as exc:
        raise RuntimeError(f"Error evaluating '{expression}': {exc}")
    if result is not None:
        try:
            json.dumps(result)
        except (TypeError, ValueError):
            result = str(result)
    return {"expression": expression, "result": result, "status": "evaluated"}
