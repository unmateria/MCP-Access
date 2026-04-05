"""
Core: COM session singleton, caches, logging, COM thread pool.
All other modules import shared state from here.
"""

import asyncio
import atexit
import concurrent.futures
import ctypes
import logging
import os
import subprocess
import time
import sys
import threading
from pathlib import Path
from typing import Any, Optional

# DPI awareness -- must be set before any window operations
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    pass

# ---------------------------------------------------------------------------
# Logging -- goes to stderr to avoid polluting the JSON-RPC stdout channel
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    stream=sys.stderr,
)
log = logging.getLogger("access-mcp")

# ---------------------------------------------------------------------------
# COM thread pool -- single thread so all COM calls stay in the same STA.
# ---------------------------------------------------------------------------

def _com_thread_init():
    """Initializer for the COM worker thread -- calls CoInitialize once."""
    import pythoncom
    pythoncom.CoInitialize()
    log.info("COM thread initialized (thread=%s)", threading.current_thread().name)

_com_executor = concurrent.futures.ThreadPoolExecutor(
    max_workers=1,
    thread_name_prefix="com-worker",
    initializer=_com_thread_init,
)

# ---------------------------------------------------------------------------
# Access COM constants
# ---------------------------------------------------------------------------
AC_TYPE: dict[str, int] = {
    "query":  1,   # acQuery
    "form":   2,   # acForm
    "report": 3,   # acReport
    "macro":  4,   # acMacro
    "module": 5,   # acModule
}

# ---------------------------------------------------------------------------
# Caches to reduce COM calls in long sessions
# ---------------------------------------------------------------------------
_vbe_code_cache: dict = {}        # "type:name" -> full text of VBE module
_parsed_controls_cache: dict = {} # "form:name" / "report:name" -> _parse_controls() result

# ---------------------------------------------------------------------------
# COM Session -- singleton, keeps Access alive between calls
# ---------------------------------------------------------------------------
class _Session:
    """
    Maintains a single Access.Application instance across tool calls.
    If a different DB is requested, closes the current one and opens the new one.
    """
    _app: Optional[Any] = None
    _db_open: Optional[str] = None
    _cm_cache: dict = {}   # "type:name" -> CodeModule COM object
    _decompiled_dbs: set = set()  # DBs already decompiled in this session

    @classmethod
    def connect(cls, db_path: str) -> Any:
        resolved = str(Path(db_path).resolve())
        if cls._app is not None:
            # Health check: verify COM session is still alive
            try:
                _ = cls._app.Visible  # cheap COM property access
            except Exception:
                log.warning("COM session stale — auto-reconnecting...")
                cls._force_cleanup()
        if cls._app is None:
            cls._launch()
        if cls._db_open != resolved:
            cls._switch(resolved)
        return cls._app

    @classmethod
    def _force_cleanup(cls):
        """Reset state without calling methods on a dead COM object."""
        cls._app = None
        cls._db_open = None
        cls._cm_cache.clear()
        cls._decompiled_dbs.clear()
        _vbe_code_cache.clear()
        _parsed_controls_cache.clear()

    @classmethod
    def _launch(cls) -> None:
        try:
            import win32com.client
        except ImportError:
            raise RuntimeError(
                "pywin32 not installed. Run: pip install pywin32"
            )
        log.info("Launching Access.Application...")
        cls._app = win32com.client.Dispatch("Access.Application")
        try:
            cls._app.Visible = True   # required for VBE to be accessible via COM
        except Exception as e:
            log.warning("Could not set Visible=True: %s (continuing anyway)", e)
        log.info("Access launched OK")

    @classmethod
    def reopen(cls, path: str) -> None:
        """Forces reopen with SHIFT (bypass AutoExec) via _switch().
        Use after CloseCurrentDatabase+CompactRepair in maintenance."""
        cls._db_open = None
        cls._switch(path)

    @classmethod
    def _decompile(cls, path: str) -> None:
        """Run MSACCESS /decompile + SHIFT on the DB before opening via COM.
        Strips orphaned p-code so compile errors are real, not phantom."""
        msaccess_candidates = [
            r"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE",
        ]
        msaccess = next((p for p in msaccess_candidates if os.path.exists(p)), None)
        if not msaccess:
            log.warning("MSACCESS.EXE not found — skipping /decompile")
            cls._decompiled_dbs.add(path)  # don't retry
            return

        # Close COM session completely so the file is unlocked
        if cls._app is not None:
            log.info("Closing COM session for /decompile...")
            try:
                if cls._db_open:
                    cls._app.CloseCurrentDatabase()
            except Exception:
                pass
            try:
                cls._app.Quit(1)  # acQuitSaveNone
            except Exception:
                pass
            cls._app = None
            cls._db_open = None
            cls._cm_cache.clear()
            _vbe_code_cache.clear()
            _parsed_controls_cache.clear()

        log.info("Decompiling %s ...", path)

        # Hold SHIFT while launching /decompile
        VK_SHIFT = 0x10
        KEYEVENTF_KEYUP = 0x0002
        _kbd = ctypes.windll.user32.keybd_event
        shift_held = False
        try:
            _kbd(VK_SHIFT, 0, 0, 0)
            time.sleep(0.3)
            shift_held = True
        except Exception:
            pass

        proc = subprocess.Popen(
            [msaccess, path, "/decompile"],
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
        )
        time.sleep(3)
        if shift_held:
            try:
                _kbd(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
            except Exception:
                pass
        time.sleep(5)  # total ~8s for /decompile to finish
        try:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                capture_output=True, timeout=10,
            )
        except Exception:
            pass

        cls._decompiled_dbs.add(path)
        log.info("Decompile done for %s", path)

        # Re-launch COM (was killed above)
        cls._launch()

    @classmethod
    def _switch(cls, path: str) -> None:
        if not os.path.isfile(path):
            raise FileNotFoundError(f"File not found: {path}")

        # Auto-decompile on first open of each DB in this session
        if path not in cls._decompiled_dbs:
            cls._decompile(path)

        if cls._db_open is not None:
            log.info("Closing previous DB: %s", cls._db_open)
            try:
                cls._app.CloseCurrentDatabase()
            except Exception as e:
                log.warning("Error closing previous DB: %s", e)
        log.info("Opening DB: %s", path)

        # Hold Shift during OpenCurrentDatabase to bypass AutoExec/startup forms
        VK_SHIFT = 0x10
        KEYEVENTF_KEYUP = 0x0002
        _kbd = ctypes.windll.user32.keybd_event
        shift_held = False
        try:
            _kbd(VK_SHIFT, 0, 0, 0)  # Press SHIFT
            time.sleep(0.3)  # Let the key state register before COM call
            shift_held = True
            log.info("SHIFT held for bypass")
        except Exception:
            log.warning("Could not simulate Shift — AutoExec may run")

        try:
            cls._app.OpenCurrentDatabase(path)
        except Exception as e:
            if "already have the database open" in str(e).lower():
                log.info("DB was already open — syncing state")
            else:
                raise
        finally:
            if shift_held:
                try:
                    _kbd(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)  # Release SHIFT
                    log.info("SHIFT released")
                except Exception:
                    pass

        cls._db_open = path

        # Close any auto-opened forms (safety net)
        try:
            for i in range(cls._app.Forms.Count - 1, -1, -1):
                try:
                    name = cls._app.Forms(i).Name
                    cls._app.DoCmd.Close(2, name)  # acForm
                    log.info("Closed auto-opened form: %s", name)
                except Exception:
                    pass
        except Exception:
            pass

        # Clear caches on DB switch
        cls._cm_cache.clear()
        _vbe_code_cache.clear()
        _parsed_controls_cache.clear()
        log.info("DB opened OK")

    @classmethod
    def quit(cls) -> None:
        if cls._app is not None:
            log.info("Closing Access...")
            try:
                if cls._db_open:
                    cls._app.CloseCurrentDatabase()
                cls._app.Quit()
                log.info("Access closed OK")
            except Exception as e:
                log.warning("Error closing Access: %s", e)
            finally:
                cls._app = None
                cls._db_open = None
                cls._cm_cache.clear()
                cls._decompiled_dbs.clear()
                _vbe_code_cache.clear()
                _parsed_controls_cache.clear()


atexit.register(_Session.quit)


def invalidate_all_caches():
    """Convenience: clear all 3 caches at once."""
    _vbe_code_cache.clear()
    _parsed_controls_cache.clear()
    _Session._cm_cache.clear()


def invalidate_object_caches(object_type: str, object_name: str):
    """Clear caches for a specific object."""
    cache_key = f"{object_type}:{object_name}"
    _vbe_code_cache.pop(cache_key, None)
    _parsed_controls_cache.pop(cache_key, None)
    _Session._cm_cache.pop(cache_key, None)
