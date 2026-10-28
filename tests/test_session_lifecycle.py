"""Lifecycle regression tests without opening Microsoft Access."""

import asyncio
import concurrent.futures
import threading
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from mcp_access import core, database, server


@pytest.fixture
def session(monkeypatch):
    for name, value in (
        ("_app", None), ("_db_open", None), ("_pid", None),
        ("_attached", False), ("_cm_cache", {}),
        ("_decompiled_dbs", set()), ("_shared_open_warning", None),
    ):
        monkeypatch.setattr(core._Session, name, value)
    monkeypatch.setattr(core, "_parsed_controls_cache", {})
    monkeypatch.setattr(core, "_release_shift", Mock())
    monkeypatch.setattr(core._Session, "_stop_dialog_watchdog", Mock())
    return core._Session


def test_attached_disconnect_preserves_database_and_clears_proxies(session):
    app = Mock()
    session._app = app
    session._attached = True
    session._db_open = "example.accdb"
    session._pid = 1234
    session._cm_cache["module:modExample"] = object()
    core._parsed_controls_cache["form:frmOrders"] = {}

    result = session.quit()

    app.CloseCurrentDatabase.assert_not_called()
    app.Quit.assert_not_called()
    assert "remains open" in result
    assert session._app is None
    assert session._pid is None
    assert not session._cm_cache
    assert not core._parsed_controls_cache


def test_owned_quit_releases_children_before_quit_and_app_before_wait(session, monkeypatch):
    app = Mock()
    session._app = app
    session._pid = 1234
    session._db_open = "example.accdb"
    session._cm_cache["module:modExample"] = object()

    def quitting():
        assert not session._cm_cache

    def waiting(pid, timeout):
        assert pid == 1234 and timeout == 10.0
        assert session._app is None
        return True

    app.Quit.side_effect = quitting
    monkeypatch.setattr(core, "_wait_pid_gone", waiting)

    assert "exited" in session.quit()
    app.CloseCurrentDatabase.assert_not_called()
    app.Quit.assert_called_once_with()
    assert session._pid is None


def test_owned_timeout_is_reported_without_killing(session, monkeypatch):
    app = Mock()
    session._app = app
    session._pid = 1234
    monkeypatch.setattr(core, "_wait_pid_gone", Mock(return_value=False))
    run = Mock()
    monkeypatch.setattr(core.subprocess, "run", run)

    with pytest.raises(RuntimeError, match="did not exit within 10 seconds"):
        session.quit()

    run.assert_not_called()
    assert session._app is None
    assert session._pid is None


def test_quit_failure_is_not_reported_as_success(session):
    session._app = Mock()
    session._app.Quit.side_effect = OSError("save cancelled")
    with pytest.raises(RuntimeError, match="save cancelled"):
        session.quit()
    assert session._app is None


def test_quit_failure_still_checks_if_process_exited(session, monkeypatch):
    session._app = Mock()
    session._app.Quit.side_effect = OSError("quit failed")
    session._pid = 1234
    wait = Mock(return_value=False)
    monkeypatch.setattr(core, "_wait_pid_gone", wait)

    with pytest.raises(RuntimeError, match="PID 1234 is still running"):
        session.quit()

    wait.assert_called_once_with(1234, timeout=10.0)


def test_unknown_pid_is_not_reported_as_verified_exit(session):
    session._app = Mock()
    assert "could not be verified" in session.quit()


def test_same_attached_database_is_not_closed(session, tmp_path):
    target = tmp_path / "example.accdb"
    target.touch()
    app = Mock()
    app.CurrentDb.return_value = SimpleNamespace(Name=str(target))
    session._app = app
    session._attached = True

    session._switch(str(target))

    app.CloseCurrentDatabase.assert_not_called()
    app.OpenCurrentDatabase.assert_not_called()
    assert session._db_open == str(target)


def test_attached_database_can_be_reopened_when_currentdb_is_unavailable(
    session, monkeypatch, tmp_path,
):
    target = tmp_path / "example.accdb"
    target.touch()
    app = Mock()
    current = [None]

    def current_db():
        if current[0] is None:
            raise RuntimeError("no open database")
        return SimpleNamespace(Name=current[0])

    app.CurrentDb.side_effect = current_db
    app.CurrentProject.FullName = ""
    app.hWndAccessApp = 0
    app.Forms.Count = 0
    app.OpenCurrentDatabase.side_effect = lambda path: current.__setitem__(0, path)
    session._app = app
    session._attached = True
    monkeypatch.delenv("MCP_ACCESS_EXCLUSIVE", raising=False)
    monkeypatch.setattr(core, "_lock_file_in_use", lambda path: False)
    monkeypatch.setattr(core, "_press_shift_bypass", lambda context: False)
    monkeypatch.setattr(session, "_suppress_recovery_dialog", Mock())

    session._switch(str(target))

    app.OpenCurrentDatabase.assert_called_once_with(str(target))


def test_attached_state_unknown_refuses_to_switch(session, monkeypatch, tmp_path):
    target = tmp_path / "example.accdb"
    target.touch()
    app = Mock()
    app.CurrentDb.side_effect = RuntimeError("COM failed")
    app.CurrentProject.FullName = Mock(side_effect=RuntimeError("COM failed"))
    session._app = app
    session._attached = True
    session._db_open = str(target)

    with pytest.raises(RuntimeError, match="Cannot safely determine"):
        session._switch(str(target))

    app.CloseCurrentDatabase.assert_not_called()
    app.OpenCurrentDatabase.assert_not_called()


def test_switch_leaves_users_database_and_launches_own(session, monkeypatch, tmp_path):
    target = tmp_path / "target.accdb"
    target.touch()
    app = Mock()
    # Live state differs from the cache: the user changed databases manually.
    app.CurrentDb.return_value = SimpleNamespace(Name=str(tmp_path / "user.accdb"))
    session._app = app
    session._attached = True
    session._db_open = str(target)

    class StopAfterLaunch(Exception):
        pass

    launch = Mock(side_effect=StopAfterLaunch)
    monkeypatch.setattr(session, "_launch", launch)
    with pytest.raises(StopAfterLaunch):
        session._switch(str(target))

    app.CloseCurrentDatabase.assert_not_called()
    app.Quit.assert_not_called()
    launch.assert_called_once_with(str(target))


def test_connect_checks_live_attached_database_even_when_cache_matches(
    session, monkeypatch, tmp_path,
):
    target = str(tmp_path / "target.accdb")
    app = Mock()
    app.CurrentDb.return_value = SimpleNamespace(Name=str(tmp_path / "user.accdb"))
    session._app = app
    session._attached = True
    session._db_open = target
    launch = Mock(side_effect=RuntimeError("test launch stopped"))
    monkeypatch.setattr(session, "_launch", launch)

    with pytest.raises(RuntimeError, match="test launch stopped"):
        session.connect(target)

    app.CloseCurrentDatabase.assert_not_called()
    launch.assert_called_once_with(target)


@pytest.mark.parametrize("attached", [False, True])
def test_opening_target_keeps_owned_switch_and_idle_attach_behaviour(
    session, monkeypatch, tmp_path, attached,
):
    target = str(tmp_path / "target.accdb")
    (tmp_path / "target.accdb").touch()
    app = Mock()
    current = [None if attached else str(tmp_path / "previous.accdb")]
    app.CurrentDb.side_effect = lambda: (
        SimpleNamespace(Name=current[0]) if current[0] else None
    )
    app.OpenCurrentDatabase.side_effect = lambda path: current.__setitem__(0, path)
    app.Forms.Count = 0
    app.hWndAccessApp = 0
    session._app = app
    session._attached = attached
    session._db_open = current[0]
    monkeypatch.delenv("MCP_ACCESS_EXCLUSIVE", raising=False)
    monkeypatch.setattr(core, "_lock_file_in_use", lambda path: False)
    monkeypatch.setattr(core, "_press_shift_bypass", lambda context: False)
    monkeypatch.setattr(session, "_suppress_recovery_dialog", Mock())

    session._switch(target)

    app.OpenCurrentDatabase.assert_called_once_with(target)
    assert app.CloseCurrentDatabase.call_count == (0 if attached else 1)
    assert session._db_open == target
    if attached:
        assert "remains open" in session.quit()
        assert current[0] == target
        app.Quit.assert_not_called()


def test_create_does_not_close_attached_database(session, monkeypatch, tmp_path):
    app = Mock()
    app.CurrentDb.return_value = SimpleNamespace(Name=str(tmp_path / "user.accdb"))
    session._app = app
    session._attached = True
    session._db_open = str(tmp_path / "user.accdb")
    target = str(tmp_path / "new.accdb")
    launch = Mock(side_effect=RuntimeError("test launch stopped"))
    monkeypatch.setattr(session, "_launch", launch)

    with pytest.raises(RuntimeError, match="test launch stopped"):
        database.ac_create_database(target)

    app.CloseCurrentDatabase.assert_not_called()
    app.NewCurrentDatabase.assert_not_called()
    launch.assert_called_once_with(target)


def test_shutdown_balances_com_on_worker_thread(session, monkeypatch):
    import pythoncom

    events = []
    monkeypatch.setattr(pythoncom, "CoInitialize",
                        lambda: events.append(("init", threading.get_ident())))
    monkeypatch.setattr(pythoncom, "CoUninitialize",
                        lambda: events.append(("uninit", threading.get_ident())))
    monkeypatch.setattr(session, "quit",
                        lambda: events.append(("quit", threading.get_ident())))
    with concurrent.futures.ThreadPoolExecutor(
        max_workers=1, initializer=core._com_thread_init,
    ) as executor:
        executor.submit(core.shutdown_com_thread).result()
        executor.submit(core.shutdown_com_thread).result()

    assert [name for name, _ in events] == ["init", "quit", "uninit"]
    assert len({ident for _, ident in events}) == 1
    assert events[0][1] != threading.get_ident()


def test_wait_pumps_sta_and_closes_process_handle(monkeypatch):
    import pythoncom
    import win32api
    import win32event

    handle = Mock()
    pump = Mock()
    monkeypatch.setattr(win32api, "OpenProcess", Mock(return_value=handle))
    monkeypatch.setattr(win32event, "WaitForSingleObject",
                        Mock(side_effect=[win32event.WAIT_TIMEOUT,
                                          win32event.WAIT_OBJECT_0]))
    monkeypatch.setattr(pythoncom, "PumpWaitingMessages", pump)
    monkeypatch.setattr(core.time, "sleep", Mock())

    assert core._wait_pid_gone(1234)
    pump.assert_called_once_with()
    handle.Close.assert_called_once_with()


def test_decompile_launch_failure_releases_shift_and_restores_attached_database(
    session, monkeypatch, tmp_path,
):
    target = tmp_path / "example.accdb"
    target.touch()
    current = [str(target)]
    app = Mock()

    def current_db():
        if current[0] is None:
            raise RuntimeError("database is closed")
        return SimpleNamespace(Name=current[0])

    def close_database():
        current[0] = None

    def open_database(path):
        current[0] = path

    app.CurrentDb.side_effect = current_db
    app.CurrentProject = SimpleNamespace(FullName="")
    app.CloseCurrentDatabase.side_effect = close_database
    app.OpenCurrentDatabase.side_effect = open_database
    app.hWndAccessApp = 0
    app.Forms.Count = 0
    session._app = app
    session._attached = True
    session._db_open = str(target)
    session._office_msaccess = "MSACCESS.EXE"
    monkeypatch.setattr(session, "_detect_office_install", Mock())
    monkeypatch.setattr(core.os.path, "exists", lambda path: True)
    monkeypatch.setattr(core, "_list_msaccess_pids", lambda: set())
    monkeypatch.setattr(core.subprocess, "Popen", Mock(side_effect=OSError("launch failed")))
    release = Mock()
    monkeypatch.setattr(core, "_release_shift", release)
    monkeypatch.setattr(core, "_press_shift_bypass", Mock(side_effect=[True, False]))
    monkeypatch.setattr(core, "_lock_file_in_use", lambda path: False)
    monkeypatch.setattr(session, "_suppress_recovery_dialog", Mock())
    monkeypatch.delenv("MCP_ACCESS_EXCLUSIVE", raising=False)

    with pytest.raises(RuntimeError, match="Could not start /decompile"):
        session._decompile(str(target))

    assert current[0] == str(target)
    release.assert_called_once_with()


def test_wait_reports_timeout_and_closes_handle(monkeypatch):
    import win32api
    import win32event

    handle = Mock()
    monkeypatch.setattr(win32api, "OpenProcess", Mock(return_value=handle))
    monkeypatch.setattr(win32event, "WaitForSingleObject",
                        Mock(return_value=win32event.WAIT_TIMEOUT))

    assert core._wait_pid_gone(1234, timeout=0) is False
    handle.Close.assert_called_once_with()


@pytest.mark.parametrize("error, expected", [(87, True), (5, None)])
def test_wait_does_not_treat_access_denied_as_process_exit(monkeypatch, error, expected):
    import pywintypes
    import win32api

    monkeypatch.setattr(win32api, "OpenProcess",
                        Mock(side_effect=pywintypes.error(error, "OpenProcess", "test")))
    if expected:
        assert core._wait_pid_gone(1234) is True
    else:
        with pytest.raises(RuntimeError, match="Could not verify"):
            core._wait_pid_gone(1234)


@pytest.mark.parametrize("window", [123, lambda: 123])
def test_pid_capture_accepts_property_or_method(monkeypatch, window):
    def get_pid(hwnd, pid):
        assert hwnd == 123
        pid._obj.value = 1234

    monkeypatch.setattr(core.ctypes, "windll", SimpleNamespace(
        user32=SimpleNamespace(GetWindowThreadProcessId=get_pid),
    ))
    assert core._get_com_pid(SimpleNamespace(hWndAccessApp=window)) == 1234


def test_shutdown_uninitializes_even_when_quit_fails(session, monkeypatch):
    import pythoncom

    uninit = Mock()
    monkeypatch.setattr(pythoncom, "CoInitialize", Mock())
    monkeypatch.setattr(pythoncom, "CoUninitialize", uninit)
    monkeypatch.setattr(session, "quit", Mock(side_effect=RuntimeError("quit failed")))
    with concurrent.futures.ThreadPoolExecutor(
        max_workers=1, initializer=core._com_thread_init,
    ) as executor:
        with pytest.raises(RuntimeError, match="quit failed"):
            executor.submit(core.shutdown_com_thread).result()
    uninit.assert_called_once_with()


@pytest.mark.parametrize("fail", [False, True])
def test_server_cleans_up_on_worker_after_stdio_exit(session, monkeypatch, fail):
    from contextlib import asynccontextmanager

    events = []

    @asynccontextmanager
    async def stdio():
        yield (object(), object())

    async def run(*args):
        if fail:
            raise RuntimeError("transport failed")

    def cleanup():
        events.append(threading.get_ident())

    executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
    monkeypatch.setattr(server, "_com_executor", executor)
    monkeypatch.setattr(server, "stdio_server", stdio)
    monkeypatch.setattr(server.server, "run", run)
    monkeypatch.setattr(core, "shutdown_com_thread", cleanup)
    monkeypatch.setattr(server, "shutdown_com_thread", cleanup, raising=False)

    if fail:
        with pytest.raises(RuntimeError, match="transport failed"):
            asyncio.run(server.main())
    else:
        asyncio.run(server.main())
    assert len(events) == 1
    assert events[0] != threading.get_ident()
    with pytest.raises(RuntimeError, match="shutdown"):
        executor.submit(lambda: None)
