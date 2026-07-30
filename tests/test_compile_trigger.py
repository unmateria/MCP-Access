"""
Pure-Python tests for the compile-trigger fix (false "NOT compiled").

Based on the tests contributed by @CaptainStormfield in PR #35.

No COM / no Access needed — these exercise `_ensure_code_pane` in
`mcp_access/compile.py` with fake VBE objects. The function must:
  - do NOTHING when a code pane of the current database's project is already
    active (a repeat compile must not keep opening code windows in the user's
    VBE, and enumerating a large project costs a COM round-trip per component);
  - otherwise prefer a STANDARD module's code pane (Type 1) over document
    modules, regardless of enumeration order (showing a standard module's pane
    never has Design-view side effects);
  - skip empty modules;
  - fall back to a document module when no standard module has code;
  - never raise: a project that can't be resolved or enumerated returns
    ok=False with a diagnostic detail, because the caller treats the pane
    activation as best-effort.

The end-to-end behaviour (Execute() succeeding after pane activation, the
honest trigger-failure report) requires a live Access instance.

Run with:
    python -m pytest tests/test_compile_trigger.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import mcp_access.compile as compile_mod  # noqa: E402
from mcp_access.core import _Session  # noqa: E402


# ---------------------------------------------------------------------------
# Fakes — just enough VBE surface for _ensure_code_pane
# ---------------------------------------------------------------------------

class _FakePane:
    def __init__(self, module=None):
        self.shown = False
        self.CodeModule = module

    def Show(self):
        self.shown = True


class _FakeCM:
    def __init__(self, count_of_lines, parent=None):
        self.CountOfLines = count_of_lines
        self.CodePane = _FakePane(self)
        self.Parent = parent


class _FakeComp:
    def __init__(self, name, comp_type, lines):
        self.Name = name
        self.Type = comp_type  # 1 = standard module, 100 = form/report doc
        self.CodeModule = _FakeCM(lines, parent=self)
        self.Collection = None  # set by _FakeProj


class _FakeProj:
    def __init__(self, name, comps, filename=""):
        self.Name = name
        self.FileName = filename
        self.VBComponents = comps
        for c in comps:
            c.Collection = self  # Collection.Parent -> project; self-referential
    # VBComponents.Parent is the project — flatten the chain onto ourselves.
    @property
    def Parent(self):
        return self


class _RaisingProj:
    """VBComponents access itself raises (broken reference / Trust Center)."""

    Name = "BrokenProj"
    FileName = ""

    @property
    def VBComponents(self):
        raise RuntimeError("VBComponents unavailable")


class _FakeVBE:
    def __init__(self, active_pane=None):
        self._active = active_pane

    @property
    def ActiveCodePane(self):
        if self._active is None:
            raise RuntimeError("no active code pane")
        return self._active


class _FakeApp:
    def __init__(self, active_pane=None):
        self.VBE = _FakeVBE(active_pane)


def _run_with_project(proj_or_exc, app=None, db_open=None):
    """Call _ensure_code_pane with _get_vb_project and _db_open stubbed out."""
    orig = compile_mod._get_vb_project
    orig_db = _Session._db_open
    if isinstance(proj_or_exc, Exception):
        def stub(app_):
            raise proj_or_exc
    else:
        def stub(app_):
            return proj_or_exc
    compile_mod._get_vb_project = stub
    _Session._db_open = db_open
    try:
        return compile_mod._ensure_code_pane(app if app is not None else _FakeApp())
    finally:
        compile_mod._get_vb_project = orig
        _Session._db_open = orig_db


# ---------------------------------------------------------------------------
# Fast path: an active pane of OUR project means touch nothing
# ---------------------------------------------------------------------------

def test_active_pane_of_our_project_short_circuits(tmp_path):
    db = tmp_path / "erp.accdb"
    db.write_text("x")
    std = _FakeComp("modUtils", 1, lines=10)
    proj = _FakeProj("MyProject", [std], filename=str(db))
    app = _FakeApp(active_pane=std.CodeModule.CodePane)

    result = _run_with_project(proj, app=app, db_open=str(db))

    assert result["ok"] is True
    assert result["detail"] == "code pane already active"
    # Nothing was re-shown: the user's VBE layout is left exactly as it was.
    assert std.CodeModule.CodePane.shown is False


def test_active_pane_of_another_project_is_ignored(tmp_path):
    """After a decompile/compact the active pane is typically acwzmain's —
    exactly the case that used to compile the wrong project."""
    db = tmp_path / "erp.accdb"
    db.write_text("x")
    wizard_mod = _FakeComp("wzmain", 1, lines=99)
    _FakeProj("acwzmain", [wizard_mod], filename=str(tmp_path / "acwzmain.accde"))

    ours = _FakeComp("modUtils", 1, lines=10)
    our_proj = _FakeProj("MyProject", [ours], filename=str(db))
    app = _FakeApp(active_pane=wizard_mod.CodeModule.CodePane)

    result = _run_with_project(our_proj, app=app, db_open=str(db))

    assert result["ok"] is True
    assert "modUtils" in result["detail"]
    assert ours.CodeModule.CodePane.shown is True


# ---------------------------------------------------------------------------
# Pane selection
# ---------------------------------------------------------------------------

def test_prefers_standard_module_over_document_module():
    form = _FakeComp("Form_frmMain", 100, lines=50)
    std = _FakeComp("modUtils", 1, lines=10)
    # Document module enumerates FIRST — the std module must still win.
    proj = _FakeProj("MyProject", [form, std])
    result = _run_with_project(proj)
    assert result["ok"] is True
    assert result["project"] == "MyProject"
    assert "modUtils" in result["detail"]
    assert std.CodeModule.CodePane.shown is True
    assert form.CodeModule.CodePane.shown is False


def test_skips_empty_modules():
    empty_std = _FakeComp("modEmpty", 1, lines=0)
    std = _FakeComp("modReal", 1, lines=3)
    proj = _FakeProj("P", [empty_std, std])
    result = _run_with_project(proj)
    assert result["ok"] is True
    assert "modReal" in result["detail"]
    assert empty_std.CodeModule.CodePane.shown is False


def test_falls_back_to_document_module_when_no_standard_has_code():
    empty_std = _FakeComp("modEmpty", 1, lines=0)
    form = _FakeComp("Form_frmMain", 100, lines=20)
    proj = _FakeProj("P", [empty_std, form])
    result = _run_with_project(proj)
    assert result["ok"] is True
    assert "Form_frmMain" in result["detail"]
    assert form.CodeModule.CodePane.shown is True


# ---------------------------------------------------------------------------
# Never fatal — the caller carries on with ok=False
# ---------------------------------------------------------------------------

def test_no_code_anywhere_reports_not_ok_with_project_name():
    proj = _FakeProj("EmptyProj", [_FakeComp("modEmpty", 1, lines=0)])
    result = _run_with_project(proj)
    assert result["ok"] is False
    assert result["project"] == "EmptyProj"
    assert "no component with code" in result["detail"]


def test_get_vb_project_raising_is_not_fatal():
    result = _run_with_project(RuntimeError("no project"))
    assert result["ok"] is False
    assert result["project"] is None
    assert "no VBProject" in result["detail"]


def test_vbcomponents_raising_is_not_fatal():
    result = _run_with_project(_RaisingProj())
    assert result["ok"] is False
    assert result["project"] == "BrokenProj"
    assert "not enumerable" in result["detail"]


def test_component_raising_is_skipped_not_fatal():
    class _BadComp:
        Name = "modBad"

        @property
        def Type(self):
            raise RuntimeError("component is broken")

    good = _FakeComp("modGood", 1, lines=5)
    proj = _FakeProj("P", [good])
    proj.VBComponents = [_BadComp(), good]
    result = _run_with_project(proj)
    assert result["ok"] is True
    assert "modGood" in result["detail"]


def test_broken_active_pane_does_not_prevent_opening_one(tmp_path):
    """ActiveCodePane raising (no pane open at all) is the NORMAL case on a
    fresh VBE — it must fall through to the enumeration, not bail out."""
    db = tmp_path / "erp.accdb"
    db.write_text("x")
    std = _FakeComp("modUtils", 1, lines=10)
    proj = _FakeProj("P", [std], filename=str(db))
    result = _run_with_project(proj, app=_FakeApp(active_pane=None), db_open=str(db))
    assert result["ok"] is True
    assert std.CodeModule.CodePane.shown is True
