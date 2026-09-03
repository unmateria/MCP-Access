"""
Pure-Python tests for the conditional attach policy in `_Session._launch`.

Issue #38 (@Access-Abraxas): the server attached to whatever Access instance
`GetActiveObject` returned, then `ac_create_database` called
`CloseCurrentDatabase()` on it — closing the database whose own VBA was driving
the MCP call.  `_launch` now takes the database we are about to open and only
attaches to an instance that is idle or already holds that file.

No COM / no Access needed: `win32com.client` is imported inside `_launch`, so a
stub in `sys.modules` is enough.

Run with:
    python -m pytest tests/test_attach_policy.py
"""

import os
import sys
import types
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from mcp_access.core import _Session  # noqa: E402

ENV = "MCP_ACCESS_EXCLUSIVE"


class _FakeDb:
    def __init__(self, name):
        self.Name = name


class _FakeApp:
    """Minimal Access.Application stand-in."""

    def __init__(self, tag, db_name=None):
        self.tag = tag
        self.Visible = True
        self._db_name = db_name

    def CurrentDb(self):
        return _FakeDb(self._db_name) if self._db_name else None


class _FakeWin32:
    def __init__(self, running=None):
        self.running = running
        self.dispatched = False

    def GetActiveObject(self, prog_id):
        if self.running is None:
            raise OSError("no running instance")
        return self.running

    def DispatchEx(self, prog_id):
        self.dispatched = True
        return _FakeApp("spawned")


@pytest.fixture
def launch(monkeypatch):
    """Runs `_launch` against a stubbed win32com and neutered side effects."""
    monkeypatch.delenv(ENV, raising=False)

    for name in ("_suppress_recovery_dialog", "_detect_office_install",
                 "_start_dialog_watchdog"):
        monkeypatch.setattr(_Session, name, classmethod(lambda cls: None))
    monkeypatch.setattr("mcp_access.core._get_com_pid", lambda app: 1234)

    def run(running, target_path):
        fake = _FakeWin32(running)
        win32com = types.ModuleType("win32com")
        win32com.client = fake
        monkeypatch.setitem(sys.modules, "win32com", win32com)
        monkeypatch.setitem(sys.modules, "win32com.client", fake)
        monkeypatch.setattr(_Session, "_app", None)
        monkeypatch.setattr(_Session, "_attached", False)
        monkeypatch.setattr(_Session, "_db_open", None)
        _Session._launch(target_path)
        return fake

    return run


def test_instance_with_another_database_is_left_alone(tmp_path, launch):
    """The issue #38 case: we must not touch an Access that is busy with
    somebody else's database."""
    other = str(tmp_path / "user_work.accdb")
    target = str(tmp_path / "brand_new.accdb")

    fake = launch(_FakeApp("user", other), target)

    assert _Session._attached is False
    assert fake.dispatched is True
    assert _Session._app.tag == "spawned"
    assert _Session._db_open is None


def test_instance_with_the_target_open_is_reused(tmp_path, launch):
    """The case attach-first exists for — don't spawn a second Access on the
    database the user already has open."""
    target = str(tmp_path / "shared.accdb")

    fake = launch(_FakeApp("user", target), target)

    assert _Session._attached is True
    assert fake.dispatched is False
    assert _Session._app.tag == "user"
    assert os.path.normcase(_Session._db_open) == os.path.normcase(target)


def test_path_comparison_ignores_case_and_shape(tmp_path, launch):
    """Windows paths are case-insensitive, and callers hand us differently
    spelled but identical paths."""
    target = tmp_path / "Shared.accdb"
    open_as = str(tmp_path / "sub" / ".." / "SHARED.ACCDB")
    (tmp_path / "sub").mkdir()

    fake = launch(_FakeApp("user", open_as), str(target).upper())

    assert _Session._attached is True
    assert fake.dispatched is False


def test_idle_instance_is_reused(tmp_path, launch):
    """No database open means nothing of the user's to close."""
    fake = launch(_FakeApp("user", None), str(tmp_path / "new.accdb"))

    assert _Session._attached is True
    assert fake.dispatched is False
    assert _Session._db_open is None


def test_no_target_keeps_the_old_behaviour(tmp_path, launch):
    """A caller that doesn't know its destination must not change behaviour."""
    other = str(tmp_path / "user_work.accdb")

    fake = launch(_FakeApp("user", other), None)

    assert _Session._attached is True
    assert fake.dispatched is False
    assert os.path.normcase(_Session._db_open) == os.path.normcase(other)


def test_no_running_instance_spawns_one(tmp_path, launch):
    fake = launch(None, str(tmp_path / "new.accdb"))

    assert _Session._attached is False
    assert fake.dispatched is True


def test_exclusive_never_looks_at_the_candidate(tmp_path, monkeypatch, launch):
    """MCP_ACCESS_EXCLUSIVE skips the attach entirely — even for the target
    database, since a running instance holds it SHARED."""
    monkeypatch.setenv(ENV, "1")
    target = str(tmp_path / "shared.accdb")

    class _Exploding(_FakeApp):
        def CurrentDb(self):  # pragma: no cover - must never run
            raise AssertionError("the exclusive path inspected the candidate")

    fake = launch(_Exploding("user", target), target)

    assert _Session._attached is False
    assert fake.dispatched is True
