"""
Compact and repair operations.
"""

import os
import threading
import time
from pathlib import Path

from .core import _Session, _vbe_code_cache, _parsed_controls_cache, log


# ---------------------------------------------------------------------------
# CompactRepair watchdog -- polls for blocking dialogs (wizards, recovery
# prompts, ODBC credential prompts) while CompactRepair runs, and dismisses
# them via proper button click (Cancel first — never VK_RETURN).
# ---------------------------------------------------------------------------

def _call_with_dialog_watchdog(app, label: str, callable_fn) -> None:
    """Run a blocking COM call with a polling dialog-dismiss watchdog.

    Captures the Access hwnd on the caller (COM STA) thread, spawns a daemon
    thread that polls every 0.5s after a 1s grace period, and dismisses any
    Access-owned dialog via `_dismiss_access_dialogs` (which now uses the
    Cancel-first button priority and wizard-title detection).

    Used around any COM call that could block on an unexpected dialog:
    `CompactRepair`, `RunCommand`, etc.  `label` is used in log messages.
    """
    # Lazy import to avoid circular dependency with vba_exec
    from .vba_exec import _dismiss_access_dialogs

    # Capture hwnd on the COM-worker thread (same apartment as `app`)
    try:
        _h = app.hWndAccessApp
        hwnd = int(_h() if callable(_h) else _h)
    except Exception as e_hwnd:
        log.warning("Could not capture Access hwnd for %s watchdog: %s", label, e_hwnd)
        hwnd = 0

    stop_event = threading.Event()
    dismissed: list = []

    def _watchdog():
        if stop_event.wait(1.0):  # 1s grace period
            return
        while not stop_event.is_set():
            if hwnd:
                try:
                    if _dismiss_access_dialogs(hwnd):
                        if not dismissed:
                            log.warning("Dialog dismissed during %s", label)
                        dismissed.append(True)
                except Exception as e_wd:
                    log.warning("%s watchdog error: %s", label, e_wd)
            stop_event.wait(0.5)

    watchdog_thread = threading.Thread(target=_watchdog, daemon=True)
    watchdog_thread.start()
    try:
        callable_fn()
    finally:
        stop_event.set()


def _compact_with_watchdog(app, src: str, dst: str) -> None:
    """Call app.CompactRepair(src, dst) with a polling dialog watchdog.
    Thin wrapper around `_call_with_dialog_watchdog`.
    """
    _call_with_dialog_watchdog(app, "CompactRepair", lambda: app.CompactRepair(src, dst))


def ac_compact_repair(db_path: str) -> dict:
    """Compacts and repairs the database. Closes, compacts to temp, replaces and reopens."""
    resolved = str(Path(db_path).resolve())
    app = _Session.connect(resolved)
    original_size = os.path.getsize(resolved)

    # Close current database (keep Access alive)
    try:
        app.CloseCurrentDatabase()
    except Exception as exc:
        raise RuntimeError(f"Could not close the database for compacting: {exc}")
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
            _compact_with_watchdog(app, resolved, tmp_path)
        except Exception as exc:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)
            raise RuntimeError(f"Error en CompactRepair: {exc}")

        if not os.path.exists(tmp_path):
            raise RuntimeError("CompactRepair did not generate the output file")
        compacted_size = os.path.getsize(tmp_path)

        # Atomic swap: original -> .bak, tmp -> original
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
                _Session.reopen(resolved)
        except Exception:
            pass
        raise

    # Reopen compacted database (with SHIFT to bypass AutoExec/startup)
    try:
        _Session.reopen(resolved)
    except Exception as exc:
        raise RuntimeError(f"Database compacted OK but error reopening: {exc}")

    saved = original_size - compacted_size
    return {
        "original_size": original_size,
        "compacted_size": compacted_size,
        "saved_bytes": saved,
        "saved_pct": round(saved / original_size * 100, 1) if original_size > 0 else 0,
        "status": "compacted",
    }


def ac_decompile_compact(db_path: str) -> dict:
    """Removes orphaned VBA p-code (/decompile), recompiles and compacts. Typical reduction 60-70%."""
    import subprocess, time
    resolved = str(Path(db_path).resolve())
    if not os.path.exists(resolved):
        raise FileNotFoundError(f"Database not found: {resolved}")

    original_size = os.path.getsize(resolved)

    # 1. Close COM session and release the file completely
    try:
        app = _Session.connect(resolved)
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        _Session._db_open = None
        _Session._cm_cache.clear()
        _vbe_code_cache.clear()
        _parsed_controls_cache.clear()
        try:
            app.Quit(1)  # acQuitSaveNone=1
        except Exception:
            pass
        _Session._app = None
    except Exception:
        pass  # si no habia sesion abierta, continuar igualmente

    # 2. Lanzar MSACCESS /decompile
    msaccess_candidates = [
        r"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE",
    ]
    msaccess = next((p for p in msaccess_candidates if os.path.exists(p)), None)
    if not msaccess:
        raise RuntimeError("MSACCESS.EXE not found in known Office 16 paths")

    # Hold SHIFT during /decompile to bypass AutoExec/startup forms
    import ctypes
    VK_SHIFT = 0x10
    KEYEVENTF_KEYUP = 0x0002
    _kbd = ctypes.windll.user32.keybd_event
    shift_held = False
    try:
        _kbd(VK_SHIFT, 0, 0, 0)       # Press SHIFT
        time.sleep(0.3)                # Let key state register
        shift_held = True
    except Exception:
        pass  # SHIFT simulation failed — AutoExec may run

    proc = subprocess.Popen(
        [msaccess, resolved, "/decompile"],
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
    )

    # Polling loop: 16 × 0.5s = 8s total.  Release SHIFT at the ~3s mark,
    # and poll for any blocking dialogs (wizards, recovery prompts) via
    # _dismiss_dialogs_by_pid on the subprocess PID.
    from .vba_exec import _dismiss_dialogs_by_pid
    for i in range(16):
        if i == 6 and shift_held:  # ~3s mark
            try:
                _kbd(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)  # Release SHIFT
            except Exception:
                pass
            shift_held = False
        if proc.poll() is not None:
            break
        try:
            _dismiss_dialogs_by_pid(proc.pid)
        except Exception:
            pass
        time.sleep(0.5)

    # Ensure SHIFT released even on early exit
    if shift_held:
        try:
            _kbd(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
        except Exception:
            pass
    try:
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
            capture_output=True, timeout=10,
        )
    except Exception:
        pass
    time.sleep(1)  # let Windows evict the dead process's ROT entry

    decompile_size = os.path.getsize(resolved)

    # 3. Reabrir via COM y recompilar VBA
    app2 = _Session.connect(resolved)
    # v0.7.22 BUG FIX: 137 = acCmdNewObjectReport (opens Report Wizard!),
    # NOT acCmdCompileAllModules.  Correct values per Microsoft docs:
    #   acCmdCompileAllModules        = 125
    #   acCmdCompileAndSaveAllModules = 126
    # Every call to ac_decompile_compact was silently launching the Report
    # Wizard and blocking the COM thread indefinitely.  This is the root
    # cause of the wizard hang reported by @CaptainStormfield + @unmateria.
    # Defence-in-depth: also wrap in the dialog watchdog so any future
    # unexpected dialog during compile is dismissed automatically.
    try:
        _call_with_dialog_watchdog(
            app2, "RunCommand(compile)", lambda: app2.RunCommand(126)
        )  # acCmdCompileAndSaveAllModules
    except Exception:
        pass  # compiling is not critical for the compact
    try:
        app2.CloseCurrentDatabase()
    except Exception:
        pass
    _Session._db_open = None
    _Session._cm_cache.clear()
    _vbe_code_cache.clear()
    _parsed_controls_cache.clear()

    # 4. Compact & Repair
    db_dir = os.path.dirname(resolved)
    db_name, db_ext = os.path.splitext(os.path.basename(resolved))
    tmp_path = os.path.join(db_dir, f"{db_name}_compact_tmp{db_ext}")
    bak_path = os.path.join(db_dir, f"{db_name}_compact_bak{db_ext}")
    for p in (tmp_path, bak_path):
        if os.path.exists(p):
            os.unlink(p)

    try:
        _compact_with_watchdog(app2, resolved, tmp_path)
    except Exception as exc:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise RuntimeError(f"Error en CompactRepair: {exc}")

    if not os.path.exists(tmp_path):
        raise RuntimeError("CompactRepair did not generate the output file")

    compacted_size = os.path.getsize(tmp_path)
    os.rename(resolved, bak_path)
    try:
        os.rename(tmp_path, resolved)
    except Exception:
        os.rename(bak_path, resolved)
        raise
    try:
        os.unlink(bak_path)
    except OSError:
        pass

    # Reopen (with SHIFT to bypass AutoExec/startup)
    _Session.reopen(resolved)

    saved = original_size - compacted_size
    return {
        "original_size": original_size,
        "decompile_size": decompile_size,
        "compacted_size": compacted_size,
        "saved_bytes": saved,
        "saved_pct": round(saved / original_size * 100, 1) if original_size > 0 else 0,
        "status": "decompiled_and_compacted",
    }
