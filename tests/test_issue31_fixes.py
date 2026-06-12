"""
Pure-Python unit tests for the v0.7.44 issue #31 follow-ups.

No COM / no Access needed — these exercise the COM-free logic:
  - access_eval_vba gains a `timeout` parameter (schema + dispatch parity
    with access_run_vba).
  - _save_all_modules never raises, even on a hostile app object.
  - _sweep_orphan_eval_modules removes only std modules carrying the
    _mcp_eval_wrapper marker.
  - _Session._last_dismissed plumbing exists for the dismissal note.

Run with:
    python -m pytest tests/test_issue31_fixes.py
    python tests/test_issue31_fixes.py
"""

import inspect
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access.code import _save_all_modules, ac_delete_object  # noqa: E402
from mcp_access.core import _Session  # noqa: E402
from mcp_access.tools import _TOOL_SCHEMA_INDEX, coerce_arguments  # noqa: E402
from mcp_access.vba_exec import (  # noqa: E402
    _sweep_orphan_eval_modules,
    ac_eval_vba,
    ac_run_vba,
)


# ---------------------------------------------------------------------------
# A. access_eval_vba timeout — schema, coercion, signature parity
# ---------------------------------------------------------------------------

def test_eval_schema_has_timeout():
    props = _TOOL_SCHEMA_INDEX["access_eval_vba"]["properties"]
    assert "timeout" in props
    # _fixup_schema widens integer to also accept string-serialising clients
    assert props["timeout"]["type"] == ["integer", "string"]


def test_eval_timeout_not_required():
    assert "timeout" not in _TOOL_SCHEMA_INDEX["access_eval_vba"]["required"]


def test_coerce_eval_timeout_string():
    args = {"db_path": "x.accdb", "expression": "Date()", "timeout": "30"}
    coerce_arguments("access_eval_vba", args)
    assert args["timeout"] == 30


def test_eval_signature_parity_with_run_vba():
    eval_params = inspect.signature(ac_eval_vba).parameters
    run_params = inspect.signature(ac_run_vba).parameters
    assert "timeout" in eval_params
    assert "timeout" in run_params
    assert eval_params["timeout"].default is None


# ---------------------------------------------------------------------------
# B. _save_all_modules — best-effort, never raises
# ---------------------------------------------------------------------------

class _HostileApp:
    """Every attribute access / call blows up — _save_all_modules must absorb it."""

    def __getattr__(self, name):
        raise RuntimeError(f"COM exploded reading {name}")


class _RunCommandFailsApp:
    """RunCommand(280) raises (error 2046 path); AllModules fallback is used."""

    def __init__(self):
        self.saved = []
        outer = self

        class _Mod:
            def __init__(self, name, loaded):
                self.Name = name
                self.IsLoaded = loaded

        class _Mods:
            Count = 2

            def Item(self, i):
                return [_Mod("modA", True), _Mod("modB", False)][i]

        class _Proj:
            AllModules = _Mods()

        class _DoCmd:
            def Save(self, ac_type, name):
                outer.saved.append((ac_type, name))

        self.CurrentProject = _Proj()
        self.DoCmd = _DoCmd()

    def RunCommand(self, cmd):
        raise RuntimeError("2046 not available now")


def test_save_all_modules_never_raises():
    _save_all_modules(_HostileApp())  # must not raise


def test_save_all_modules_fallback_saves_loaded_only():
    app = _RunCommandFailsApp()
    _save_all_modules(app)
    assert app.saved == [(5, "modA")]  # acModule=5, only IsLoaded modules


def test_delete_object_saves_before_delete():
    src = inspect.getsource(ac_delete_object)
    assert src.index("_save_all_modules(app)") < src.index("app.DoCmd.DeleteObject")


# ---------------------------------------------------------------------------
# C. _sweep_orphan_eval_modules — marker-targeted, std modules only
# ---------------------------------------------------------------------------

class _FakeCodeModule:
    def __init__(self, code):
        self._lines = code.split("\n")

    @property
    def CountOfLines(self):
        return len(self._lines)

    def Lines(self, start, count):
        return "\n".join(self._lines[start - 1:start - 1 + count])


class _FakeComp:
    def __init__(self, name, comp_type, code):
        self.Name = name
        self.Type = comp_type
        self.CodeModule = _FakeCodeModule(code)


class _FakeComps:
    def __init__(self, comps):
        self._comps = comps

    @property
    def Count(self):
        return len(self._comps)

    def Item(self, i):  # 1-based, like COM
        return self._comps[i - 1]

    def Remove(self, comp):
        self._comps.remove(comp)


class _FakeProj:
    def __init__(self, comps):
        self.VBComponents = _FakeComps(comps)


def test_sweep_removes_only_marked_std_modules():
    orphan1 = _FakeComp("Module1", 1, "Public Function _mcp_eval_wrapper() As Variant\n...")
    orphan2 = _FakeComp("Module7", 1, "Public Function _mcp_eval_wrapper() As Variant\n...")
    user_mod = _FakeComp("modUtils", 1, "Option Explicit\nPublic Sub Foo()\nEnd Sub")
    # Class module mentioning the marker must NOT be touched (Type != 1)
    user_cls = _FakeComp("clsX", 2, "' uses _mcp_eval_wrapper in a comment")
    proj = _FakeProj([orphan1, user_mod, orphan2, user_cls])

    removed = _sweep_orphan_eval_modules(proj)

    assert sorted(removed) == ["Module1", "Module7"]
    remaining = [proj.VBComponents.Item(i).Name
                 for i in range(1, proj.VBComponents.Count + 1)]
    assert remaining == ["modUtils", "clsX"]


def test_sweep_never_raises_on_hostile_project():
    assert _sweep_orphan_eval_modules(_HostileApp()) == []


# ---------------------------------------------------------------------------
# D. dismissal-note plumbing
# ---------------------------------------------------------------------------

def test_session_has_last_dismissed_slot():
    assert hasattr(_Session, "_last_dismissed")


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))
