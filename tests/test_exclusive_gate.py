"""
Pure-Python tests for the opt-in exclusive-open switch.

Requested by @GPGeorge in issue #36. Shared opens cannot take a design lock, so
attaching Data Macros via SaveAsText/LoadFromText — or anything routing through
DoCmd.OpenTable acViewDesign — is dropped for any table another Access session
has open, and the run can still report success. MCP_ACCESS_EXCLUSIVE turns that
into one visible failure at open time.

It is OFF by default and fails CLOSED (like MCP_ACCESS_ALLOW_CODE_EXEC, unlike
MCP_ACCESS_SHIFT_BYPASS): the COM session holds the database between tool calls,
so an accidental exclusive open locks every other user out of a shared front-end
for as long as the server runs.

No COM / no Access needed.

Run with:
    python -m pytest tests/test_exclusive_gate.py
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from mcp_access.security import exclusive_open_enabled  # noqa: E402

ENV = "MCP_ACCESS_EXCLUSIVE"
PKG = Path(__file__).resolve().parent.parent / "mcp_access"


def test_off_when_unset(monkeypatch):
    """Shared stays the default — this changes nothing for existing users."""
    monkeypatch.delenv(ENV, raising=False)
    assert exclusive_open_enabled() is False


@pytest.mark.parametrize("value", ["1", "true", "TRUE", "yes", "YES", "on", " on "])
def test_truthy_values_enable(monkeypatch, value):
    monkeypatch.setenv(ENV, value)
    assert exclusive_open_enabled() is True


@pytest.mark.parametrize("value", ["", "   ", "0", "false", "no", "off", "maybe", "2"])
def test_everything_else_stays_off(monkeypatch, value):
    """Fails CLOSED: a typo must not lock a workgroup out of its database."""
    monkeypatch.setenv(ENV, value)
    assert exclusive_open_enabled() is False


def test_read_per_call_not_at_import(monkeypatch):
    monkeypatch.delenv(ENV, raising=False)
    assert exclusive_open_enabled() is False
    monkeypatch.setenv(ENV, "1")
    assert exclusive_open_enabled() is True
    monkeypatch.delenv(ENV, raising=False)
    assert exclusive_open_enabled() is False


def test_independent_of_the_other_two_switches(monkeypatch):
    """Three separate concerns: exclusivity must not ride on either of the
    others, and must not turn them on."""
    from mcp_access.security import code_exec_enabled, shift_bypass_enabled

    monkeypatch.delenv("MCP_ACCESS_ALLOW_CODE_EXEC", raising=False)
    monkeypatch.delenv("MCP_ACCESS_SHIFT_BYPASS", raising=False)
    monkeypatch.setenv(ENV, "1")
    assert exclusive_open_enabled() is True
    assert code_exec_enabled() is False
    assert shift_bypass_enabled() is True

    monkeypatch.setenv("MCP_ACCESS_ALLOW_CODE_EXEC", "1")
    monkeypatch.delenv(ENV, raising=False)
    assert exclusive_open_enabled() is False


# ---------------------------------------------------------------------------
# Structural guards — a half-applied gate is worse than no gate
# ---------------------------------------------------------------------------

def test_every_open_site_consults_the_gate():
    """Two places call OpenCurrentDatabase (`core._switch` and
    `database.ac_create_database`). A third one that forgets the switch would
    silently reopen the session shared, so every calling module must import the
    gate."""
    callers = [py for py in PKG.glob("*.py")
               if "OpenCurrentDatabase(" in py.read_text(encoding="utf-8")]
    assert callers, "no OpenCurrentDatabase call sites found — did the API move?"
    for py in callers:
        assert "exclusive_open_enabled" in py.read_text(encoding="utf-8"), (
            f"{py.name} opens a database without consulting "
            f"security.exclusive_open_enabled()"
        )


def test_exclusive_open_is_not_swallowed():
    """`_switch` swallows 'already have the database open' on the shared path.
    Swallowing anything on the exclusive path would restore the silent failure
    the switch removes."""
    src = (PKG / "core.py").read_text(encoding="utf-8")
    body = src.split("def _switch")[1].split("\n    @classmethod")[0]
    assert "OpenCurrentDatabase(path, True)" in body
    assert "_exclusive_open_failure" in body
    swallow = body.index("already have the database open")
    guard = body.index("if exclusive:", body.index("except Exception as e:"))
    assert guard < swallow, (
        "the exclusive branch must be checked before the "
        "'already have the database open' swallow"
    )


def test_switch_verifies_the_mode_after_opening():
    """Asking for exclusive is not getting it: with the file already in use
    Access opens it SHARED and reports nothing (measured against Access 2016 —
    no exception, CurrentDb valid, and our own entry added to the lock file).
    The check before the open is courtesy; the one after it is the guarantee,
    so both must stay."""
    src = (PKG / "core.py").read_text(encoding="utf-8")
    body = src.split("def _switch")[1].split("\n    @classmethod")[0]
    open_call = body.index("OpenCurrentDatabase(path, True)")
    checks = [i for i in range(len(body))
              if body.startswith("_lock_file_in_use(path)", i)]
    assert len(checks) >= 2, "_switch must check the lock file before AND after"
    assert checks[0] < open_call < checks[-1]


def test_launch_does_not_attach_when_exclusive():
    """Attaching to a running Access would hold the file shared while the
    session reported itself exclusive."""
    src = (PKG / "core.py").read_text(encoding="utf-8")
    body = src.split("def _launch")[1].split("\n    @classmethod")[0]
    gate = body.index("exclusive_open_enabled()")
    attach = body.index("GetActiveObject")
    assert gate < attach, "_launch attaches before consulting the exclusive gate"


# ---------------------------------------------------------------------------
# Lock file reader — "who has it open" in the failure message
# ---------------------------------------------------------------------------

def _entry(machine: str, user: str) -> bytes:
    """One 64-byte lock file record: 32 bytes computer, 32 bytes security name."""
    return machine.encode("latin-1").ljust(32, b"\x00") + \
        user.encode("latin-1").ljust(32, b"\x00")


def test_lock_file_holders_reads_entries(tmp_path):
    from mcp_access.core import _lock_file_holders

    db = tmp_path / "sample.accdb"
    db.write_bytes(b"")
    (tmp_path / "sample.laccdb").write_bytes(
        _entry("WORKSTATION1", "Admin") + _entry("WORKSTATION2", "Admin")
    )
    assert _lock_file_holders(str(db)) == [
        "WORKSTATION1 (Admin)", "WORKSTATION2 (Admin)"
    ]


def test_lock_file_holders_dedupes(tmp_path):
    """One session can own several entries; the message should name it once."""
    from mcp_access.core import _lock_file_holders

    db = tmp_path / "sample.accdb"
    db.write_bytes(b"")
    (tmp_path / "sample.laccdb").write_bytes(
        _entry("WORKSTATION1", "Admin") * 3
    )
    assert _lock_file_holders(str(db)) == ["WORKSTATION1 (Admin)"]


def test_lock_file_holders_uses_ldb_for_mdb(tmp_path):
    from mcp_access.core import _lock_file_holders

    db = tmp_path / "legacy.mdb"
    db.write_bytes(b"")
    (tmp_path / "legacy.ldb").write_bytes(_entry("WORKSTATION1", "Admin"))
    assert _lock_file_holders(str(db)) == ["WORKSTATION1 (Admin)"]


@pytest.mark.parametrize("content", [b"", b"\x00" * 64, b"short"])
def test_lock_file_holders_never_raises(tmp_path, content):
    """It only decorates an error message — a missing, empty, truncated or
    all-NUL lock file must degrade to 'unknown', never to a second exception."""
    from mcp_access.core import _lock_file_holders

    db = tmp_path / "sample.accdb"
    db.write_bytes(b"")
    (tmp_path / "sample.laccdb").write_bytes(content)
    assert _lock_file_holders(str(db)) == []
    assert _lock_file_holders(str(tmp_path / "missing.accdb")) == []


def test_failure_message_names_the_holders(tmp_path):
    from mcp_access.core import _exclusive_open_failure

    db = tmp_path / "sample.accdb"
    db.write_bytes(b"")
    (tmp_path / "sample.laccdb").write_bytes(_entry("WORKSTATION1", "Admin"))
    msg = _exclusive_open_failure(str(db), "the database is already open")
    assert "WORKSTATION1 (Admin)" in msg
    assert "the database is already open" in msg
    assert "NOT open" in msg


def test_failure_message_survives_a_missing_lock_file(tmp_path):
    from mcp_access.core import _exclusive_open_failure

    msg = _exclusive_open_failure(str(tmp_path / "gone.accdb"), "boom")
    assert "could not tell" in msg
    assert "boom" in msg


# ---------------------------------------------------------------------------
# In-use detection — an orphan lock file must not read as an occupied database
# ---------------------------------------------------------------------------

def test_lock_file_path_matches_the_database_extension():
    from mcp_access.core import _lock_file_path

    assert _lock_file_path(r"C:\x\sample.accdb").endswith("sample.laccdb")
    assert _lock_file_path(r"C:\x\legacy.mdb").endswith("legacy.ldb")
    assert _lock_file_path(r"C:\x\legacy.MDB").endswith("legacy.ldb")


def test_no_lock_file_is_not_in_use(tmp_path):
    from mcp_access.core import _lock_file_in_use

    assert _lock_file_in_use(str(tmp_path / "sample.accdb")) is False


def test_orphan_lock_file_is_not_in_use(tmp_path):
    """An Access that died without closing leaves a lock file nobody holds.
    Access opens exclusively straight over one (tested), so reading mere
    existence as 'occupied' would refuse a perfectly good open."""
    from mcp_access.core import _lock_file_in_use

    db = tmp_path / "sample.accdb"
    db.write_bytes(b"")
    (tmp_path / "sample.laccdb").write_bytes(_entry("GHOST", "Admin"))
    assert _lock_file_in_use(str(db)) is False


def test_open_lock_file_is_in_use(tmp_path):
    """A live shared session keeps its lock file open; asking for it with
    dwShareMode=0 is what tells that apart from an orphan."""
    from mcp_access.core import _lock_file_in_use

    db = tmp_path / "sample.accdb"
    db.write_bytes(b"")
    lock = tmp_path / "sample.laccdb"
    lock.write_bytes(_entry("WORKSTATION1", "Admin"))
    with open(lock, "rb"):
        assert _lock_file_in_use(str(db)) is True
    assert _lock_file_in_use(str(db)) is False
