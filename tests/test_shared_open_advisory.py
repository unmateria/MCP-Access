"""
Pure-Python tests for the SHARED-open advisory (issue #36, follow-up).

MCP_ACCESS_EXCLUSIVE turns "somebody else has this database open" into a refusal,
but it is off by default and the person most likely to point the server at a live
production front-end is the one least likely to switch it on. With the switch off
the server now still says so: `_switch` reads the lock file before opening and
parks a warning on `_Session._shared_open_warning`, which `server.call_tool`
appends to the result of the call that opened the database.

It is a warning, never a block — refusing is what the switch is for.

No COM / no Access needed.

Run with:
    python -m pytest tests/test_shared_open_advisory.py
"""

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from mcp_access.core import _Session, _shared_open_advisory  # noqa: E402

PKG = Path(__file__).resolve().parent.parent / "mcp_access"


def test_message_names_the_holders():
    msg = _shared_open_advisory(r"C:\erp\front.accdb", ["PC01 (ana)", "PC02 (luis)"])
    assert "PC01 (ana)" in msg and "PC02 (luis)" in msg
    assert r"C:\erp\front.accdb" in msg


def test_message_survives_an_unreadable_lock_file():
    """The advisory decorates a successful open — it must never raise."""
    msg = _shared_open_advisory(r"C:\erp\front.accdb", [])
    assert "could not read the lock file" in msg


def test_message_says_shared_and_points_at_the_switch():
    """A warning is only useful if it says what to do about it."""
    msg = _shared_open_advisory("db.accdb", ["PC01 (ana)"])
    assert "SHARED" in msg
    assert "MCP_ACCESS_EXCLUSIVE" in msg


def test_message_does_not_claim_the_open_failed():
    """Shared is a legitimate mode: read-only work is fine and nothing was refused."""
    msg = _shared_open_advisory("db.accdb", ["PC01 (ana)"])
    low = msg.lower()
    assert "could not open" not in low
    assert "nothing ran" not in low
    assert "read-only work is unaffected" in low


def test_session_starts_with_no_warning():
    assert _Session._shared_open_warning is None


def test_switch_only_warns_when_not_exclusive():
    """With the switch on the same situation is a refusal, not a warning —
    emitting both would contradict itself."""
    src = (PKG / "core.py").read_text(encoding="utf-8")
    body = src.split("def _switch(", 1)[1]
    assert "if not exclusive and _lock_file_in_use(path)" in body


def test_switch_reads_the_lock_file_before_touching_the_session():
    """Our own entry would be in the lock file once we open: the advisory has to
    be decided before CloseCurrentDatabase/OpenCurrentDatabase run."""
    src = (PKG / "core.py").read_text(encoding="utf-8")
    body = src.split("def _switch(", 1)[1]
    warn = body.index("cls._shared_open_warning = (")
    assert warn < body.index("OpenCurrentDatabase(path")
    assert warn < body.index("CloseCurrentDatabase()")


def test_switch_skips_the_warning_for_our_own_open_database():
    """Attaching to the user's Access must not report the user to themselves."""
    src = (PKG / "core.py").read_text(encoding="utf-8")
    body = src.split("def _switch(", 1)[1]
    assert "not cls._already_open(path)" in body


def test_already_open_never_raises_on_a_dead_com_proxy():
    _Session._app = None
    assert _Session._already_open("db.accdb") is False


def test_already_open_compares_case_insensitively():
    class _Proj:
        FullName = r"C:\ERP\Front.ACCDB"

    class _App:
        CurrentProject = _Proj()

    prev = _Session._app
    try:
        _Session._app = _App()
        assert _Session._already_open(r"c:\erp\front.accdb") is True
        assert _Session._already_open(r"c:\erp\other.accdb") is False
    finally:
        _Session._app = prev


def test_call_tool_surfaces_the_warning_on_the_opening_call():
    """Same mechanism as the auto-dismissed-dialog note: timestamp-gated so a
    later, unrelated tool call doesn't repeat it."""
    src = (PKG / "server.py").read_text(encoding="utf-8")
    assert "_Session._shared_open_warning" in src
    assert re.search(r"w\[0\]\s*>=\s*started", src)
    assert "[mcp-access] Warning: " in src


# --- The "somebody else holds it exclusively" diagnosis (shared mode) --------
# Verified live against two Access 2016 processes: a database another process
# holds EXCLUSIVELY writes NO lock file, and opening it shared leaves our
# session with no database and no exception — the same silent failure the
# exclusive switch handles, one level down. It used to be reported as an
# AutoExec/startup-form problem.

def test_db_file_probe_exists_and_is_used_by_switch():
    src = (PKG / "core.py").read_text(encoding="utf-8")
    assert "def _db_file_in_use(" in src
    body = src.split("def _switch(", 1)[1]
    assert "if _db_file_in_use(path):" in body


def test_lock_diagnosis_precedes_the_autoexec_one():
    """An exclusive holder leaves no lock file, so only the file probe can tell:
    the AutoExec message must be the fallback, not the first answer."""
    src = (PKG / "core.py").read_text(encoding="utf-8")
    body = src.split("def _switch(", 1)[1]
    assert body.index("_db_file_in_use(path)") < body.index("Database closed itself while opening")


def test_missing_file_is_not_in_use(tmp_path):
    from mcp_access.core import _db_file_in_use
    assert _db_file_in_use(str(tmp_path / "nope.accdb")) is False


def test_free_file_is_not_in_use(tmp_path):
    from mcp_access.core import _db_file_in_use
    f = tmp_path / "free.accdb"
    f.write_bytes(b"x")
    assert _db_file_in_use(str(f)) is False


def test_held_file_is_in_use(tmp_path):
    """Any live occupant — shared or exclusive — keeps a handle on the .accdb."""
    from mcp_access.core import _db_file_in_use
    f = tmp_path / "held.accdb"
    f.write_bytes(b"x")
    with open(f, "rb"):
        assert _db_file_in_use(str(f)) is True
    assert _db_file_in_use(str(f)) is False
