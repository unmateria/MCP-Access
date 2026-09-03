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
import winreg
from pathlib import Path
from typing import Any, Optional

from .security import exclusive_open_enabled, shift_bypass_enabled

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
_parsed_controls_cache: dict = {} # "form:name" / "report:name" -> _parse_controls() result


# ---------------------------------------------------------------------------
# SHIFT key safety net — always release simulated SHIFT on exit/cleanup
# ---------------------------------------------------------------------------
_VK_SHIFT = 0x10
_KEYEVENTF_KEYUP = 0x0002

def _release_shift():
    """Release SHIFT key at OS level.  Idempotent — safe to call even if not pressed."""
    try:
        ctypes.windll.user32.keybd_event(_VK_SHIFT, 0, _KEYEVENTF_KEYUP, 0)
    except Exception:
        pass

atexit.register(_release_shift)


def _press_shift_bypass(context: str) -> bool:
    """Press SHIFT for the AutoExec/startup bypass.  Returns True if it is held.

    The ONLY place in this package that synthesises a SHIFT key-down.  Every
    caller (``_switch``, ``_Session._decompile``, ``ac_decompile_compact``) went
    through its own copy of press/sleep/release before v0.7.53; funnelling them
    here is what makes the ``MCP_ACCESS_SHIFT_BYPASS`` opt-out impossible to
    half-apply.  Release with ``_release_shift()`` (idempotent).

    ``keybd_event`` is a GLOBAL key-down — it is not scoped to Access, so it
    shifts whatever the human is typing anywhere on the machine for as long as
    it is held.  That is the whole reason the switch exists; see
    ``security.shift_bypass_enabled``.

    A False return means "not held" for both reasons — opted out, or the
    injection failed — which is exactly what every caller needs to decide
    whether to release later.
    """
    if not shift_bypass_enabled():
        log.info(
            "SHIFT bypass disabled via MCP_ACCESS_SHIFT_BYPASS (%s) — an "
            "unguarded AutoExec in the target database will run; "
            "AutomationSecurity and the dialog watchdog still apply", context,
        )
        return False
    try:
        ctypes.windll.user32.keybd_event(_VK_SHIFT, 0, 0, 0)
        time.sleep(0.3)  # let the key state register before the blocking call
        log.info("SHIFT held for bypass (%s)", context)
        return True
    except Exception as exc:
        log.warning("Could not simulate SHIFT (%s) — AutoExec may run: %s",
                    context, exc)
        return False


# ---------------------------------------------------------------------------
# Lock file reader — "who has this database open" for exclusive-open failures
# ---------------------------------------------------------------------------
_LOCK_ENTRY = 64   # bytes per entry in an Access lock file
_LOCK_FIELD = 32   # first 32 = computer name, second 32 = security name
_LOCK_MAX = 255 * _LOCK_ENTRY   # Access caps concurrent users at 255 (16 KB)

_GENERIC_READ = 0x80000000
_OPEN_EXISTING = 3
_INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_kernel32.CreateFileW.restype = ctypes.c_void_p
_kernel32.CreateFileW.argtypes = (
    ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p,
    ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p,
)
_kernel32.CloseHandle.restype = ctypes.c_int
_kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)


def _lock_file_path(db_path: str) -> str:
    """The lock file Access keeps beside a database: .laccdb, or .ldb for .mdb."""
    base, ext = os.path.splitext(db_path)
    return base + (".ldb" if ext.lower() == ".mdb" else ".laccdb")


def _lock_file_in_use(db_path: str) -> bool:
    """True when the database's lock file exists AND a process still holds it.

    Existence alone does not mean the database is in use: an Access that died
    without closing leaves an orphan behind, and Access opens exclusively right
    over one, leaving the stale file untouched (tested). Opening it with
    ``dwShareMode = 0`` separates the two — a live shared session (ours
    included) makes Windows answer ERROR_SHARING_VIOLATION, an orphan opens.

    False on any error we can't interpret: this decides whether to refuse an
    open, and refusing on a permission quirk would be worse than the shared
    session it is meant to catch — the post-open `CurrentDb()` check still
    stands behind it.
    """
    path = _lock_file_path(db_path)
    if not os.path.exists(path):
        return False
    try:
        handle = _kernel32.CreateFileW(
            path, _GENERIC_READ, 0, None, _OPEN_EXISTING, 0, None,
        )
    except Exception:
        return False
    if handle == _INVALID_HANDLE_VALUE:
        return ctypes.get_last_error() == 32   # ERROR_SHARING_VIOLATION
    _kernel32.CloseHandle(handle)
    return False


def _db_file_in_use(db_path: str) -> bool:
    """True when another process holds the database FILE itself.

    The lock file only answers for SHARED sessions: a database somebody opened
    exclusively has no ``.laccdb`` at all (tested), so `_lock_file_in_use` says
    False for the very case that locks everyone else out. Probing the .accdb
    with ``dwShareMode = 0`` catches both — shared occupants and an exclusive
    one — because either keeps a handle open on the file.

    Only meaningful once OUR session has let go of it (i.e. after a failed
    open): while we have the database open we are an occupant ourselves.
    False on anything unexpected — this only picks which diagnosis a failure
    message carries.
    """
    if not os.path.exists(db_path):
        return False
    try:
        handle = _kernel32.CreateFileW(
            db_path, _GENERIC_READ, 0, None, _OPEN_EXISTING, 0, None,
        )
    except Exception:
        return False
    if handle == _INVALID_HANDLE_VALUE:
        return ctypes.get_last_error() == 32   # ERROR_SHARING_VIOLATION
    _kernel32.CloseHandle(handle)
    return False


def _lock_file_holders(db_path: str) -> list[str]:
    """Best-effort "who has this database open", read from its lock file.

    Access keeps a sibling ``.laccdb`` (``.ldb`` for ``.mdb``) while a database
    is open for shared use, with one 64-byte entry per user: 32 bytes of
    computer name then 32 bytes of security name, both NUL-padded. Format and
    the 255-user cap are documented by Microsoft:
    https://learn.microsoft.com/troubleshoot/microsoft-365-apps/access/lock-files-introduction

    Returns ``["COMPUTER (user)", ...]``, de-duplicated in file order — one
    session can own several entries. Returns an empty list if the file is
    absent, unreadable or malformed: this only ever decorates an error message,
    so it must never raise one of its own.
    """
    try:
        with open(_lock_file_path(db_path), "rb") as fh:
            raw = fh.read(_LOCK_MAX)
    except Exception:
        return []

    def _field(chunk: bytes) -> str:
        return chunk.split(b"\x00")[0].decode("latin-1", "replace").strip()

    holders: list[str] = []
    for off in range(0, len(raw) - _LOCK_ENTRY + 1, _LOCK_ENTRY):
        entry = raw[off:off + _LOCK_ENTRY]
        machine = _field(entry[:_LOCK_FIELD])
        user = _field(entry[_LOCK_FIELD:])
        if not machine and not user:
            continue
        who = f"{machine} ({user})" if machine and user else (machine or user)
        if who not in holders:
            holders.append(who)
    return holders


def _exclusive_open_failure(path: str, reason: str) -> str:
    """Message for an exclusive open that did not end up exclusive.

    Access reports none of these three outcomes as an error, so the reason has
    to be spelled out or the user is left with a bare COM failure at best and
    silence at worst. Name the occupants when the lock file can tell us.
    """
    holders = _lock_file_holders(path)
    who = ("; ".join(holders) if holders
           else "could not tell (the lock file is unreadable or already gone)")
    return (
        f"Could not open '{path}' exclusively: {reason}\n"
        f"Sessions holding it: {who}\n"
        "MCP_ACCESS_EXCLUSIVE is on, so this is all-or-nothing: the database is "
        "NOT open and nothing ran. Close the other sessions and retry, or remove "
        "MCP_ACCESS_EXCLUSIVE from this server's env and restart to work shared "
        "(design-lock work may then fail silently, per table)."
    )


def _shared_open_advisory(path: str, holders: list) -> str:
    """Warning for a SHARED open onto a database somebody else already has open.

    Not a failure: shared is the default mode and plenty of read-only work is
    fine that way. But it is the exact situation where design work silently
    half-lands — Access drops a design lock it cannot take and reports success
    per table — and the person most likely to be pointed at a live production
    front-end is the one least likely to know that (issue #36, @GPGeorge).

    So it is surfaced on the tool result rather than only in the log, and it
    names the occupants. It never blocks: turning this into a refusal is what
    MCP_ACCESS_EXCLUSIVE is for, and that has to stay an explicit choice.
    """
    who = "; ".join(holders) if holders else "could not read the lock file"
    return (
        f"'{path}' is already open in another Access session, and this server "
        f"opened it SHARED. Sessions holding it: {who}. "
        "Design changes (table/form/report design, Data Macros, anything "
        "routing through Design view) can be dropped WITHOUT an error while "
        "another session holds the object — a run can report success and have "
        "changed nothing. Read-only work is unaffected. Work on a private copy "
        "on a build machine, or set MCP_ACCESS_EXCLUSIVE=1 in this server's env "
        "and restart to make this a hard failure at open time instead."
    )


def _get_com_pid(app) -> Optional[int]:
    """Get the PID of the Access process behind a COM object via its window handle."""
    try:
        hwnd = app.hWndAccessApp()
        if not hwnd:
            return None
        pid = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return pid.value if pid.value else None
    except Exception:
        return None


def _wait_pid_gone(pid: int, timeout: float = 10.0) -> bool:
    """Poll until a PID no longer appears in the process list.  Returns True if gone."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if pid not in _list_msaccess_pids():
            return True
        time.sleep(0.3)
    return pid not in _list_msaccess_pids()


def _list_msaccess_pids() -> set[int]:
    """Return PIDs of all running msaccess.exe processes.
    Defensive helper used around the /decompile subprocess to catch forked
    children that escape `taskkill /T /F`.  Returns an empty set on any error.
    """
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq msaccess.exe", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=5,
        )
        pids: set[int] = set()
        for line in (out.stdout or "").splitlines():
            # CSV: "msaccess.exe","1234","Console","1","45,678 K"
            parts = [p.strip().strip('"') for p in line.split(",")]
            if len(parts) >= 2 and parts[0].lower() == "msaccess.exe":
                try:
                    pids.add(int(parts[1]))
                except ValueError:
                    pass
        return pids
    except Exception:
        return set()


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
    _pid: Optional[int] = None  # captured in _launch() on the COM worker thread (atexit can't query COM)
    _wd_thread: Optional[Any] = None  # global dialog watchdog thread
    _wd_stop: Optional[Any] = None    # threading.Event to stop the watchdog
    _cm_cache: dict = {}   # "type:name" -> CodeModule COM object
    _decompiled_dbs: set = set()  # DBs already decompiled in this session
    _attached: bool = False  # True if we attached via GetActiveObject; False if we spawned via DispatchEx
    # time.monotonic() when a tool call entered the COM executor; None when
    # idle.  Set/cleared by server.call_tool around run_in_executor.  Lets the
    # global dialog watchdog distinguish a modal raised by OUR blocked COM
    # call (dismiss it) from one the interactive user is looking at (leave it)
    # on attached instances.  Plain attribute read/write — atomic under GIL.
    _tool_started: Optional[float] = None
    # (time.monotonic(), dialog_title) of the last dialog auto-dismissed by
    # any watchdog (all funnel through _dismiss_dialogs_by_pid).  server.call_tool
    # compares the timestamp against the call's start to append a diagnostic
    # note ("a modal dialog was auto-dismissed during this call") to the result.
    _last_dismissed: Optional[tuple] = None
    # (time.monotonic(), message) of the last SHARED open onto a database that
    # another Access session already had open.  Set by _switch, read by
    # server.call_tool the same way as _last_dismissed, so the advisory reaches
    # the model on the call that opened the database instead of dying in the
    # log.  Only ever a warning — refusing is what MCP_ACCESS_EXCLUSIVE does.
    _shared_open_warning: Optional[tuple] = None
    # Detected Office install. Defaults are the hardcoded fallback used pre-0.7.36.
    # _detect_office_install() runs once per process and updates these in place.
    _office_version: str = "16.0"
    _office_msaccess: Optional[str] = None
    _office_detected: bool = False

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
            else:
                # The app proxy can be alive while the database has been
                # closed under us (e.g. by its own startup code, or by the
                # user in an attached instance).  A stale _db_open wedges
                # every later CurrentDb/CurrentData call with "object is
                # closed or doesn't exist" — and on attached instances
                # reconnects re-attach to the same broken instance forever.
                # CurrentDb() costs one extra COM round-trip per tool call;
                # acceptable (~ms) for the recovery it buys.
                if cls._db_open is not None:
                    db_alive = False
                    try:
                        db_alive = cls._app.CurrentDb() is not None
                    except Exception:
                        pass
                    if not db_alive:
                        log.warning(
                            "Database closed under the session (%s) — "
                            "auto-reconnecting...", cls._db_open,
                        )
                        cls._force_cleanup()
        if cls._app is None:
            cls._launch(resolved)
        if cls._db_open != resolved:
            cls._switch(resolved)
        return cls._app

    @classmethod
    def _force_cleanup(cls):
        """Reset state without calling methods on a dead COM object."""
        cls._stop_dialog_watchdog()
        _release_shift()
        cls._app = None
        cls._db_open = None
        cls._pid = None
        cls._attached = False
        cls._cm_cache.clear()
        cls._decompiled_dbs.clear()
        _parsed_controls_cache.clear()

    # -----------------------------------------------------------------------
    # Global dialog watchdog
    # -----------------------------------------------------------------------
    @classmethod
    def _start_dialog_watchdog(cls) -> None:
        """Start a background thread that dismisses Access-owned modal dialogs
        which persist past a grace period.

        This backstops operations that have NO watchdog of their own — most
        importantly VBE code-module access (get_proc / find_definition /
        module_info ...).  Without it, a blocking modal raised during the
        startup form of a DB such as
            'Error accessing file. Network connection may have been lost.'
        hangs the COM call indefinitely (observed: a 1-hour hang on a DB
        with a startup form).  Operation-specific watchdogs (open / compile /
        run_vba) react faster and screenshot first; the grace period keeps
        this thread from stealing their dialog.

        On spawned instances every persistent dialog is ours to dismiss.
        On ATTACHED instances (user's own Access) we only dismiss while one
        of OUR tool calls is in flight (_tool_started) — a modal raised then
        was almost certainly provoked by that blocked COM call (e.g. a VBA
        project that fails to load popping 'Error accessing file. Network
        connection may have been lost.'), and without dismissal the tool
        call hangs until a human clicks.  A dialog with no tool call in
        flight belongs to the interactive user and is never touched."""
        pid = cls._pid
        if not pid:
            return
        # Always (re)start fresh so the watchdog tracks the CURRENT pid.
        cls._stop_dialog_watchdog()
        stop = threading.Event()
        cls._wd_stop = stop

        def _loop():
            from .vba_exec import _find_dialog_hwnds_by_pid, _dismiss_dialogs_by_pid
            GRACE = 3.0            # s a dialog must persist before we force-dismiss
            GRACE_ATTACHED = 5.0   # more conservative on the user's own Access
            POLL = 0.5
            seen: dict = {}   # hwnd -> monotonic first-seen
            while not stop.wait(POLL):
                try:
                    hwnds = _find_dialog_hwnds_by_pid(pid)
                except Exception:
                    continue
                now = time.monotonic()
                for h in list(seen):
                    if h not in hwnds:
                        del seen[h]
                for h in hwnds:
                    seen.setdefault(h, now)
                attached = cls._attached
                grace = GRACE_ATTACHED if attached else GRACE
                if attached:
                    # Only act while one of our tool calls has been blocked
                    # at least as long as the grace period.
                    started = cls._tool_started
                    if started is None or now - started < grace:
                        continue
                stale = [h for h in hwnds if now - seen.get(h, now) >= grace]
                if stale:
                    log.warning(
                        "Global dialog watchdog: %d dialog(s) persisted >%.0fs "
                        "(pid=%s, attached=%s) -- dismissing",
                        len(stale), grace, pid, attached,
                    )
                    try:
                        _dismiss_dialogs_by_pid(pid)
                    except Exception as e:
                        log.warning("Global watchdog dismiss error: %s", e)
                    seen.clear()  # re-grace any survivors before re-hitting

        t = threading.Thread(target=_loop, name="mcp-access-global-dialog-wd",
                             daemon=True)
        cls._wd_thread = t
        t.start()
        log.info("Global dialog watchdog started (pid=%s, attached=%s)",
                 pid, cls._attached)

    @classmethod
    def _stop_dialog_watchdog(cls) -> None:
        try:
            if cls._wd_stop is not None:
                cls._wd_stop.set()
        except Exception:
            pass
        cls._wd_thread = None
        cls._wd_stop = None

    @classmethod
    def _launch(cls, target_path: Optional[str] = None) -> None:
        """Bring up the COM session.

        `target_path` is the database we are about to open, when the caller
        knows it.  It decides whether attaching to a running Access instance is
        acceptable — see the attach block below.  None keeps the historical
        attach-anything behaviour for callers with no destination.
        """
        cls._suppress_recovery_dialog()
        try:
            import win32com.client
        except ImportError:
            raise RuntimeError(
                "pywin32 not installed. Run: pip install pywin32"
            )
        # Prefer attaching to an already-running Access instance so we don't
        # spawn a second process when the user already has Access open on the
        # database we are about to work on. Fall back to DispatchEx otherwise —
        # it is also required after /decompile kills to bypass stale ROT
        # entries, but in that path no live instance exists.
        # Attaching is skipped entirely under MCP_ACCESS_EXCLUSIVE: a running
        # instance holds its database SHARED, and connect() does not re-open a
        # path that is already recorded in _db_open — so attaching would leave
        # the session shared while reporting itself exclusive, which is the
        # silent-success failure the switch exists to remove.
        cls._app = None   # the DispatchEx below is keyed on this, not on the caller
        if exclusive_open_enabled():
            log.info(
                "MCP_ACCESS_EXCLUSIVE is on — not attaching to a running "
                "Access instance (it would hold the database shared)"
            )
        else:
            try:
                candidate = win32com.client.GetActiveObject("Access.Application")
                # Sanity check: round-trip a cheap property to catch marshalled
                # zombie references from a dying process. If this raises, we
                # treat it as "no running instance" and fall through.
                _ = candidate.Visible  # noqa: F841
                try:
                    current_db = candidate.CurrentDb()
                    db_name = current_db.Name if current_db is not None else None
                    candidate_db = str(Path(db_name).resolve()) if db_name else None
                except Exception:
                    candidate_db = None
                # Attaching to an instance that holds a DIFFERENT database means
                # we close that database out from under whoever is using it:
                # ac_create_database calls CloseCurrentDatabase on whatever it
                # finds, and _switch reopens the instance elsewhere.  Issue #38:
                # that shut down the database whose own VBA was driving this
                # server.  So attach only to an idle instance, or to one that
                # already holds our target.
                #
                # GetActiveObject returns a single ROT entry, so with several
                # Access windows open we can only inspect one of them and may
                # spawn a process even though another window had the target
                # open.  Worse than ideal, better than hijacking whichever
                # instance answered.
                if (
                    target_path is not None
                    and candidate_db is not None
                    and os.path.normcase(candidate_db)
                    != os.path.normcase(str(Path(target_path).resolve()))
                ):
                    # Nothing to release: dropping the reference is all we can
                    # do (Quit would close the user's Access, which is the very
                    # thing this branch exists to prevent).
                    candidate = None
                    cls._attached = False
                    log.info(
                        "Not attaching to the running Access instance: it has "
                        "%s open and we need %s — launching our own instead",
                        candidate_db, target_path,
                    )
                else:
                    cls._app = candidate
                    cls._attached = True
                    cls._db_open = candidate_db
                    log.info("Attached to existing Access.Application instance")
                    if candidate_db:
                        log.info("Existing Access has DB open: %s", candidate_db)
            except Exception:
                pass
        if cls._app is None:
            log.info("Launching new Access.Application...")
            cls._app = win32com.client.DispatchEx("Access.Application")
            cls._attached = False
            log.info("Access launched OK")
        log.info("Decompile/compile strategy: attached=%s", cls._attached)
        try:
            cls._app.Visible = True   # required for VBE to be accessible via COM
        except Exception as e:
            log.warning("Could not set Visible=True: %s (continuing anyway)", e)
        # Capture PID HERE on the COM worker thread.  quit() runs in the atexit
        # main thread where cross-thread COM calls return RPC_E_WRONG_THREAD,
        # so re-querying the proxy from quit() would silently return None and
        # disable the taskkill fallback below.
        cls._pid = _get_com_pid(cls._app)
        log.info("Access PID captured: %s (attached=%s)", cls._pid, cls._attached)
        # One-shot Office install detection — populates _office_version and
        # _office_msaccess so /decompile and the Resiliency registry path
        # match the real install. Never raises; logs a warning and keeps
        # defaults if nothing matched.
        cls._detect_office_install()
        # Global background dialog watchdog — backstops operations without
        # their own watchdog (VBE access) so a blocking modal like
        # "Error accessing file. Network connection may have been lost."
        # can't hang a COM call forever.
        cls._start_dialog_watchdog()

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
        # Prefer the detected MSACCESS.EXE; fall back to the hardcoded Office16
        # paths if detection didn't run / fired blank. Keeps behavior identical
        # on machines with a normal install.
        cls._detect_office_install()  # idempotent
        msaccess = cls._office_msaccess
        if not msaccess or not os.path.exists(msaccess):
            for p in (
                r"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE",
                r"C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE",
            ):
                if os.path.exists(p):
                    msaccess = p
                    break
        if not msaccess:
            log.warning("MSACCESS.EXE not found — skipping /decompile")
            cls._decompiled_dbs.add(path)  # don't retry
            return

        resolved_target = str(Path(path).resolve())

        # Release the file lock so the /decompile subprocess can open it.
        # When we attached to the user's Access, NEVER call Quit(1) — it would
        # kill the user's session.  Only close the current DB if it matches
        # our target; if a *different* DB is open we refuse, to avoid silently
        # closing the user's unsaved work.
        if cls._app is not None:
            log.info("Preparing COM session for /decompile (attached=%s)...", cls._attached)
            if cls._db_open:
                current_norm = str(Path(cls._db_open).resolve())
                if current_norm != resolved_target:
                    raise RuntimeError(
                        f"Cannot run /decompile on {resolved_target}: a different "
                        f"database is currently open ({current_norm}). Close it first."
                    )
                try:
                    cls._app.CloseCurrentDatabase()
                except Exception:
                    pass
            if cls._attached:
                # Keep the user's Access alive — just drop our references to
                # caches/state that become invalid after the subprocess runs.
                cls._db_open = None
                cls._cm_cache.clear()
                _parsed_controls_cache.clear()
            else:
                try:
                    cls._app.Quit(1)  # acQuitSaveNone
                except Exception:
                    pass
                cls._app = None
                cls._db_open = None
                cls._attached = False
                cls._cm_cache.clear()
                _parsed_controls_cache.clear()

        log.info("Decompiling %s ...", path)

        # Snapshot msaccess.exe PIDs so we can kill any forked children that
        # escape `taskkill /T /F`.  When attached, the user's PID is in this
        # set and MUST be preserved.
        pids_before = _list_msaccess_pids()

        # Hold SHIFT while launching /decompile.  This path holds it ~3 s (see
        # the i == 6 release below) — long enough to mangle real typing, hence
        # the MCP_ACCESS_SHIFT_BYPASS opt-out.
        shift_held = _press_shift_bypass("/decompile")

        proc = subprocess.Popen(
            [msaccess, path, "/decompile"],
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
        )

        # Polling loop: 16 × 0.5s = 8s total.  Release SHIFT at ~3s mark
        # and poll for any blocking dialogs (wizards, recovery prompts)
        # via _dismiss_dialogs_by_pid.  Breaks early if the subprocess
        # exits on its own.
        for i in range(16):
            if i == 6 and shift_held:  # ~3s mark
                _release_shift()
                shift_held = False
            if proc.poll() is not None:
                break  # subprocess already exited
            try:
                from .vba_exec import _dismiss_dialogs_by_pid  # lazy — circular
                _dismiss_dialogs_by_pid(proc.pid)
            except Exception:
                pass
            time.sleep(0.5)

        # Ensure SHIFT released even on early exit
        if shift_held:
            _release_shift()
        try:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                capture_output=True, timeout=10,
            )
        except Exception:
            pass

        # Defence-in-depth: kill any msaccess.exe PIDs that appeared during
        # the subprocess (forked children that escaped taskkill /T).  NEVER
        # touch PIDs that were already running before we started.
        try:
            pids_after = _list_msaccess_pids()
            new_pids = pids_after - pids_before
            for pid in new_pids:
                log.warning("Killing leaked msaccess.exe PID %s from /decompile", pid)
                try:
                    subprocess.run(
                        ["taskkill", "/F", "/T", "/PID", str(pid)],
                        capture_output=True, timeout=5,
                    )
                except Exception:
                    pass
        except Exception:
            pass

        cls._decompiled_dbs.add(path)
        log.info("Decompile done for %s", path)

        # Re-launch COM only when we actually killed it.  If we attached to
        # the user's Access we kept _app alive; verify it's still responsive
        # and only relaunch on failure.
        if cls._app is None:
            time.sleep(1)  # let Windows evict the dead process's ROT entry
            cls._launch(path)
        else:
            try:
                _ = cls._app.Visible
            except Exception as e:
                log.warning("Attached COM object died during /decompile (%s) — relaunching", e)
                cls._app = None
                cls._attached = False
                time.sleep(1)
                cls._launch(path)

    @classmethod
    def _detect_office_install(cls) -> None:
        """Best-effort detection of the installed Office version and MSACCESS.EXE.

        Populates `cls._office_version` (string like "16.0") and `cls._office_msaccess`
        (absolute path to MSACCESS.EXE, or None).  ALWAYS leaves the class with
        usable defaults — never raises. Called once per process at the end of
        `_launch()`; later calls are no-ops thanks to `_office_detected`.

        Detection order:
          1. Enumerate `Software\Microsoft\Office\<ver>\Access\InstallRoot\Path`
             under HKLM (64-bit), HKLM\WOW6432Node (32-bit on 64-bit) and HKCU
             (per-user Click-to-Run).  Pick the highest <ver> with a real file.
          2. App Paths\MSACCESS.EXE\(Default) — works even without per-version key.
          3. Final fallback: keep the hardcoded defaults so existing behaviour
             is preserved on machines with a broken or non-standard registry.
        """
        if cls._office_detected:
            return
        cls._office_detected = True  # set first so partial failures still terminate

        def _try_install_root(root, sub_root: str) -> Optional[tuple[str, str]]:
            """Enumerate version subkeys under *root\\sub_root* and return
            (version, MSACCESS.EXE path) for the highest version with a
            working Access\\InstallRoot\\Path. None if nothing found."""
            try:
                with winreg.OpenKey(root, sub_root) as office_key:
                    versions: list[tuple[float, str]] = []
                    i = 0
                    while True:
                        try:
                            sub = winreg.EnumKey(office_key, i)
                        except OSError:
                            break
                        i += 1
                        # Versions look like "16.0", "15.0", "14.0"; ignore
                        # "ClickToRun", "Common", "Outlook", etc.
                        try:
                            v_num = float(sub)
                        except ValueError:
                            continue
                        versions.append((v_num, sub))
                    versions.sort(reverse=True)  # highest first
                    for v_num, v_str in versions:
                        path_sub = f"{sub_root}\\{v_str}\\Access\\InstallRoot"
                        try:
                            with winreg.OpenKey(root, path_sub) as ir:
                                install_dir, _ = winreg.QueryValueEx(ir, "Path")
                                candidate = os.path.join(install_dir, "MSACCESS.EXE")
                                if os.path.exists(candidate):
                                    return v_str, candidate
                        except FileNotFoundError:
                            continue
                        except OSError:
                            continue
            except FileNotFoundError:
                return None
            except OSError:
                return None
            return None

        # Try every plausible hive/wow path
        registry_candidates = [
            (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Office"),
            (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Microsoft\Office"),
            (winreg.HKEY_CURRENT_USER,  r"Software\Microsoft\Office"),
        ]
        for root, sub in registry_candidates:
            hit = _try_install_root(root, sub)
            if hit:
                ver, path = hit
                cls._office_version = ver
                cls._office_msaccess = path
                log.info("Detected Office %s at %s", ver, path)
                return

        # Fallback 2: App Paths
        try:
            with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"Software\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE",
            ) as ap:
                default_val, _ = winreg.QueryValueEx(ap, None)
                if default_val and os.path.exists(default_val):
                    cls._office_msaccess = default_val
                    # Best-effort: parse Office<NN> from the path to set version
                    low = default_val.lower()
                    for v in ("16.0", "15.0", "14.0"):
                        if f"office{v.split('.')[0]}" in low:
                            cls._office_version = v
                            break
                    log.info(
                        "Detected MSACCESS.EXE via App Paths: %s (version=%s)",
                        default_val, cls._office_version,
                    )
                    return
        except FileNotFoundError:
            pass
        except OSError:
            pass

        log.warning(
            "Could not detect Office install via registry — using hardcoded "
            "defaults (version=%s, MSACCESS=%s)",
            cls._office_version, cls._office_msaccess,
        )

    @classmethod
    def _suppress_recovery_dialog(cls) -> None:
        """Nuke the Resiliency keys that cause the 'serious error' dialog."""
        # Detection is idempotent and uses only winreg / os.path — safe to run
        # even on the very first call before any COM object exists. Without
        # this, the Resiliency path would always target "16.0" no matter what
        # version is installed.
        cls._detect_office_install()
        ver = cls._office_version  # detected (or default "16.0")
        base = rf"Software\Microsoft\Office\{ver}\Access\Resiliency"
        # Delete subkeys that track crashed files (DisabledItems, StartupItems)
        for subkey_name in ("DisabledItems", "StartupItems"):
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, base + "\\" + subkey_name)
                log.info("Deleted Resiliency\\%s", subkey_name)
            except FileNotFoundError:
                pass
            except Exception as e:
                log.warning("Could not delete Resiliency\\%s: %s", subkey_name, e)
        # Also set the suppression flags as belt-and-suspenders
        try:
            key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, base, 0, winreg.KEY_SET_VALUE)
            try:
                winreg.SetValueEx(key, "DisableAllCallersWarning", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(key, "DoNotShowUI", 0, winreg.REG_DWORD, 1)
            finally:
                winreg.CloseKey(key)
            log.info("Recovery dialog suppressed via registry")
        except Exception as e:
            log.warning("Could not set Resiliency flags: %s", e)

    @classmethod
    def _already_open(cls, path: str) -> bool:
        """True when the live Access instance already has THIS file open.

        Its own entry is in the lock file, so without this check attaching to a
        user's Access (or a redundant reopen) would report the user's own
        session back to them as a foreign occupant.  Never raises: a False here
        only costs a warning that is one word wrong about who is holding it.
        """
        try:
            current = cls._app.CurrentProject.FullName
        except Exception:
            return False
        return bool(current) and os.path.normcase(str(current)) == os.path.normcase(path)

    @classmethod
    def _switch(cls, path: str) -> None:
        if not os.path.isfile(path):
            raise FileNotFoundError(f"File not found: {path}")

        exclusive = exclusive_open_enabled()   # read once — used either side of the open

        # Refuse before touching the session, not after: asking for an exclusive
        # open on a database somebody else has open does NOT fail — Access opens
        # it shared and says nothing (tested), which would silently break the
        # promise of the switch AND add us to the lock file. Checked again after
        # the open, which is what actually guarantees it.
        if exclusive and _lock_file_in_use(path):
            raise RuntimeError(_exclusive_open_failure(
                path, "the database is already open in another session"))

        # Shared-mode advisory: the switch above refuses; with it off we still
        # know the answer, and staying silent is how someone ends up doing
        # design work on a live front-end and believing it landed.  Read the
        # lock file BEFORE closing/opening anything, so our own session is not
        # one of the holders — except when this very instance already has the
        # file open (attached to the user's Access), which is not a warning.
        cls._shared_open_warning = None
        if not exclusive and _lock_file_in_use(path) and not cls._already_open(path):
            holders = _lock_file_holders(path)
            msg = _shared_open_advisory(path, holders)
            log.warning("Shared open onto a database in use: %s", holders or "?")
            cls._shared_open_warning = (time.monotonic(), msg)

        if cls._db_open is not None:
            log.info("Closing previous DB: %s", cls._db_open)
            try:
                cls._app.CloseCurrentDatabase()
            except Exception as e:
                log.warning("Error closing previous DB: %s", e)
        log.info("Opening DB%s: %s", " EXCLUSIVELY" if exclusive else "", path)

        cls._suppress_recovery_dialog()

        # Hold Shift during OpenCurrentDatabase to bypass AutoExec/startup forms
        shift_held = _press_shift_bypass("OpenCurrentDatabase")

        # Capture Access hwnd on THIS thread (COM worker — same apartment
        # that created _app).  COM STA proxies cannot be accessed from the
        # watchdog thread, so we must resolve hwnd here before spawning it.
        try:
            _h = cls._app.hWndAccessApp
            access_hwnd = int(_h() if callable(_h) else _h)
        except Exception as e_hwnd:
            log.warning("Could not capture Access hwnd for watchdog: %s", e_hwnd)
            access_hwnd = 0

        # Polling watchdog: after a 2s grace period, poll every 0.5s and
        # dismiss any dialog via proper button click (Cancel first — NEVER
        # VK_RETURN, which would advance wizards and create stray objects).
        _open_done = threading.Event()
        _dialog_screenshots: list = []

        def _watchdog():
            # Lazy import to avoid circular dependency with vba_exec
            from .vba_exec import _dismiss_access_dialogs
            # Grace period — normal opens complete in <2s
            if _open_done.wait(2):
                return
            log.warning("OpenCurrentDatabase blocked >2s — polling for dialogs")
            while not _open_done.is_set():
                if access_hwnd:
                    try:
                        if _dismiss_access_dialogs(
                            access_hwnd,
                            _dialog_screenshots if len(_dialog_screenshots) == 0 else None,
                        ):
                            log.warning("Dialog dismissed during OpenCurrentDatabase")
                    except Exception as e_wd:
                        log.warning("Watchdog dismiss error: %s", e_wd)
                _open_done.wait(0.5)

        watchdog_thread = threading.Thread(target=_watchdog, daemon=True)
        watchdog_thread.start()

        # Call OpenCurrentDatabase in the current thread (COM worker — same
        # apartment that created _app).  The watchdog will dismiss any dialog.
        # AutomationSecurity=3 as defence-in-depth.  Does NOT suppress Access
        # AutoExec macro objects (tested — Access ignores it for those), but
        # may prevent VBA auto-run code in edge cases where the Shift key
        # doesn't register (remote desktop, key event eaten by another app).
        try:
            cls._app.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
        except Exception:
            pass
        try:
            if exclusive:
                # OpenCurrentDatabase(filepath, Exclusive, bstrPassword) — the
                # mode is the second positional argument and defaults to False.
                cls._app.OpenCurrentDatabase(path, True)
            else:
                cls._app.OpenCurrentDatabase(path)
        except Exception as e:
            if exclusive:
                # Nothing is swallowed on this path: the point of the switch is
                # that a refused lock is one visible failure at open time rather
                # than a silent partial one per table later. The previous
                # database was already closed above, so record that nothing is
                # open before unwinding — a stale _db_open would wedge every
                # later CurrentDb call.
                cls._db_open = None
                raise RuntimeError(_exclusive_open_failure(path, str(e))) from e
            if "already have the database open" in str(e).lower():
                log.info("DB was already open — syncing state")
            else:
                raise
        finally:
            _open_done.set()  # signal watchdog to stop
            try:
                cls._app.AutomationSecurity = 1  # msoAutomationSecurityLow — restore
            except Exception:
                pass
            if shift_held:
                _release_shift()
                log.info("SHIFT released")

        if _dialog_screenshots:
            log.warning("A blocking dialog was auto-dismissed. Screenshot: %s",
                        _dialog_screenshots[0])

        # Post-open validation: confirm the database actually stayed open.
        # Startup code can close the database during an automated open — e.g.
        # an AutoExec/startup-form error path firing because backend links are
        # broken, with AllowBypassKey=False defeating the SHIFT bypass (the
        # watchdog's Cancel click then routes through the form's error
        # handler, which closes the db).  Without this check, _db_open would
        # record a database that isn't there, and every later call would die
        # at CurrentDb/CurrentData with "object is closed or doesn't exist",
        # wedging the session permanently.  Also covers the case where the
        # "already have the database open" swallow above masked a dead db.
        db_alive = False
        try:
            db_alive = cls._app.CurrentDb() is not None
        except Exception as e_val:
            log.warning("Post-open validation raised: %s", e_val)
        if not db_alive:
            log.error(
                "Database closed itself during open: %s — resetting session",
                path,
            )
            # quit() does the right thing for both cases: spawned instances
            # are quit (taskkill fallback included); attached instances are
            # released without touching the user's Access.
            try:
                cls.quit()
            except Exception as e_q:
                log.warning("Session teardown after failed open: %s", e_q)
                cls._force_cleanup()
            if exclusive:
                # An exclusive open onto a database somebody else holds
                # exclusively lands here: no exception, no database, no
                # explanation. Blaming AutoExec would send the user hunting
                # through startup code for a lock problem.
                raise RuntimeError(_exclusive_open_failure(
                    path, "Access refused the open and left the session with no "
                          "database (another session most likely holds it "
                          "exclusively)"))
            # Same wrong-diagnosis trap the exclusive path already avoids, one
            # level down: a database another process holds EXCLUSIVELY refuses
            # to open with no exception and no database, and blaming AutoExec
            # sends the user hunting through startup code for a lock problem
            # (observed live: a second Access opening a db a first one created).
            # The lock file cannot answer this — an exclusive holder writes
            # none — so the .accdb itself is probed. Our own session has
            # already been torn down above, so a live handle is somebody else.
            if _db_file_in_use(path):
                holders = _lock_file_holders(path)
                who = ("; ".join(holders) if holders
                       else "no lock file — most likely a single session "
                            "holding it exclusively")
                raise RuntimeError(
                    f"Could not open '{path}': another process has the file "
                    f"open and Access refused the open without an error. "
                    f"Sessions holding it: {who}. Close the other session and "
                    "retry. (Set MCP_ACCESS_EXCLUSIVE=1 in this server's env "
                    "to make lock conflicts a hard failure at open time.)"
                )
            raise RuntimeError(
                f"Database closed itself while opening: {path}. Its startup "
                "code (AutoExec / startup form) most likely failed and closed "
                "the database — common causes: missing backend table links, "
                "or AllowBypassKey=False defeating the SHIFT bypass. The COM "
                "session has been reset; open the file manually in Access "
                "while holding SHIFT to investigate."
            )

        # The open reported success — but "exclusive" is a request, not a
        # guarantee: with the file already in use Access downgrades to shared
        # without a word, and the only visible difference is that it joins the
        # lock file instead of leaving none behind (tested). A live lock file
        # here means we are shared, so close it rather than let design work run
        # against a session that only believes it is exclusive.
        if exclusive and _lock_file_in_use(path):
            log.error("Exclusive open silently downgraded to shared: %s", path)
            try:
                cls._app.CloseCurrentDatabase()
            except Exception as e_c:
                log.warning("Could not close the downgraded session: %s", e_c)
            cls._db_open = None
            raise RuntimeError(_exclusive_open_failure(
                path, "Access downgraded the open to shared mode because the "
                      "database was already in use"))

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
        _parsed_controls_cache.clear()
        log.info("DB opened OK")

    @classmethod
    def quit(cls) -> None:
        cls._stop_dialog_watchdog()
        if cls._app is not None:
            if cls._attached:
                log.info("Releasing attached Access.Application (not quitting user's session)")
                cls._app = None
                cls._db_open = None
                cls._pid = None
                cls._attached = False
                cls._cm_cache.clear()
                cls._decompiled_dbs.clear()
                _parsed_controls_cache.clear()
                return
            _release_shift()
            pid = cls._pid  # captured in _launch() on the COM worker; safe across threads
            log.info("Closing Access (PID %s)...", pid)
            try:
                if cls._db_open:
                    cls._app.CloseCurrentDatabase()
                cls._app.Quit()
            except Exception as e:
                log.warning("Error during Access Quit(): %s", e)
            finally:
                cls._app = None
                cls._db_open = None
                cls._pid = None
                cls._attached = False
                cls._cm_cache.clear()
                cls._decompiled_dbs.clear()
                _parsed_controls_cache.clear()
            if pid:
                if _wait_pid_gone(pid, timeout=10.0):
                    log.info("Access PID %s exited cleanly", pid)
                else:
                    log.warning("Access PID %s still alive after 10s — killing", pid)
                    try:
                        subprocess.run(
                            ["taskkill", "/F", "/T", "/PID", str(pid)],
                            capture_output=True, timeout=5,
                        )
                        _wait_pid_gone(pid, timeout=3.0)
                    except Exception as e:
                        log.warning("taskkill failed for PID %s: %s", pid, e)
                # Let Windows release file locks
                time.sleep(0.5)
            else:
                time.sleep(1.0)
            log.info("Access closed OK")


atexit.register(_Session.quit)


def _get_vb_project(app):
    """Return the VBProject that belongs to the current database.

    ``app.VBE.VBProjects(1)`` may return the wrong project (e.g. the
    ``acwzmain`` wizard library) after a decompile+compact cycle.  This
    helper enumerates all loaded VBProjects and picks the one whose
    ``.FileName`` matches ``_Session._db_open``.  Falls back to index 1
    if no match is found (single-project scenario).
    """
    db_path = _Session._db_open
    try:
        projects = app.VBE.VBProjects
        count = projects.Count
        if db_path:
            db_norm = os.path.normcase(os.path.abspath(db_path))
            for i in range(1, count + 1):
                try:
                    proj = projects(i)
                    fname = getattr(proj, "FileName", "") or ""
                    if fname and os.path.normcase(os.path.abspath(fname)) == db_norm:
                        return proj
                except Exception:
                    continue
        # Fallback: first project
        return projects(1)
    except Exception:
        return app.VBE.VBProjects(1)


def invalidate_all_caches():
    """Convenience: clear all caches at once."""
    _parsed_controls_cache.clear()
    _Session._cm_cache.clear()


def invalidate_object_caches(object_type: str, object_name: str):
    """Clear caches for a specific object."""
    cache_key = f"{object_type}:{object_name}"
    _parsed_controls_cache.pop(cache_key, None)
    _Session._cm_cache.pop(cache_key, None)
