"""
Pure-Python unit tests for the compile-trigger fix (false "NOT compiled").

No COM / no Access needed — these exercise `_ensure_code_pane` in
`mcp_access/compile.py` with fake VBE objects. The function must:
  - prefer a STANDARD module's code pane (Type 1) over document modules,
    regardless of enumeration order (showing a form module's pane is
    harmless, but a standard module never has Design-view side effects);
  - skip empty modules;
  - fall back to a document module when no standard module has code;
  - never raise: a project that can't be resolved or enumerated returns
    ok=False with a diagnostic detail, because the caller treats the pane
    activation as best-effort.

The end-to-end behaviour (Execute() succeeding after pane activation, the
honest trigger-failure report) requires a live Access instance and is
covered by manual verification against a real database — see the PR notes.

Run with:
    python -m pytest tests/test_compile_trigger.py
    python tests/test_compile_trigger.py            (plain — runs all asserts)
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import mcp_access.compile as compile_mod  # noqa: E402


# ---------------------------------------------------------------------------
# Fakes — just enough VBE surface for _ensure_code_pane
# ---------------------------------------------------------------------------

class _FakePane:
    def __init__(self):
        self.shown = False

    def Show(self):
        self.shown = True


class _FakeCM:
    def __init__(self, count_of_lines):
        self.CountOfLines = count_of_lines
        self.CodePane = _FakePane()


class _FakeComp:
    def __init__(self, name, comp_type, lines):
        self.Name = name
        self.Type = comp_type  # 1 = standard module, 100 = form/report doc
        self.CodeModule = _FakeCM(lines)


class _FakeProj:
    def __init__(self, name, comps):
        self.Name = name
        self.VBComponents = comps


class _RaisingProj:
    """VBComponents access itself raises (broken reference / Trust Center)."""

    Name = "BrokenProj"

    @property
    def VBComponents(self):
        raise RuntimeError("VBComponents unavailable")


def _run_with_project(proj_or_exc):
    """Call _ensure_code_pane with _get_vb_project stubbed out."""
    orig = compile_mod._get_vb_project
    if isinstance(proj_or_exc, Exception):
        def stub(app):
            raise proj_or_exc
    else:
        def stub(app):
            return proj_or_exc
    compile_mod._get_vb_project = stub
    try:
        return compile_mod._ensure_code_pane(app=object())
    finally:
        compile_mod._get_vb_project = orig


# ---------------------------------------------------------------------------
# Tests
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
        Name = "bad"

        @property
        def Type(self):
            raise RuntimeError("COM error")

    std = _FakeComp("modOK", 1, lines=5)
    proj = _FakeProj("P", [_BadComp(), std])
    result = _run_with_project(proj)
    assert result["ok"] is True
    assert "modOK" in result["detail"]


# ---------------------------------------------------------------------------
# Plain runner (no pytest dependency)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_") and callable(v)]
    failed = 0
    for fn in fns:
        try:
            fn()
            print(f"PASS {fn.__name__}")
        except AssertionError as e:
            failed += 1
            print(f"FAIL {fn.__name__}: {e}")
    print(f"\n{len(fns) - failed}/{len(fns)} passed")
    sys.exit(1 if failed else 0)
