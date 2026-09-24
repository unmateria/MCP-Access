"""
Pure-Python tests for databases that reference a library (PR #40, v0.7.63).

Access loads a referenced library's VBA project into the same
``VBE.VBProjects`` collection as the host. Measured on Access 2016: right
after opening a host that references a library, ``VBE.ActiveVBProject`` is
the LIBRARY, so anything reading it lists / edits the wrong project.

Covers:
  - ``_get_vb_project`` picks the host by ``CurrentProject.FullName`` first,
    ``_Session._db_open`` second, and never lets a stale ``_db_open`` that
    names the library win over the live host path;
  - no code path in the package reads ``VBE.ActiveVBProject`` any more;
  - the block checker closes ``Loop`` / ``Wend`` followed by a comment.

Run with:
    python -m pytest tests/test_referenced_library.py
"""

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access.core import _Session, _get_vb_project  # noqa: E402
from mcp_access.compile import _check_blocks_in_module  # noqa: E402

HOST = r"C:\data\frontEnd.accdb"
LIB = r"C:\data\library.accda"


class _FakeProject:
    def __init__(self, name, filename):
        self.Name = name
        self.FileName = filename


class _FakeProjects:
    def __init__(self, projects):
        self._projects = projects
        self.Count = len(projects)

    def __call__(self, index):
        return self._projects[index - 1]


class _FakeVBE:
    def __init__(self, projects):
        self.VBProjects = _FakeProjects(projects)


class _FakeCurrentProject:
    def __init__(self, full_name):
        self.FullName = full_name


class _FakeApp:
    def __init__(self, projects, current_path):
        self.VBE = _FakeVBE(projects)
        self._current_path = current_path

    @property
    def CurrentProject(self):
        if self._current_path is None:
            raise RuntimeError("CurrentProject unavailable")
        return _FakeCurrentProject(self._current_path)


def _resolve(projects, current_path, db_open):
    orig = _Session._db_open
    _Session._db_open = db_open
    try:
        return _get_vb_project(_FakeApp(projects, current_path)).Name
    finally:
        _Session._db_open = orig


LIB_FIRST = [_FakeProject("library", LIB), _FakeProject("frontEnd", HOST)]


def test_host_found_when_library_is_index_1():
    assert _resolve(LIB_FIRST, HOST, HOST) == "frontEnd"


def test_live_path_beats_stale_db_open_naming_the_library():
    # The library is enumerated first AND is what the stale cache names;
    # the live CurrentProject path must still decide.
    assert _resolve(LIB_FIRST, HOST, LIB) == "frontEnd"


def test_db_open_used_when_current_project_unavailable():
    assert _resolve(LIB_FIRST, None, HOST) == "frontEnd"


def test_path_comparison_ignores_case():
    assert _resolve(LIB_FIRST, HOST.upper(), None) == "frontEnd"


def test_falls_back_to_index_1_when_nothing_matches():
    assert _resolve(LIB_FIRST, r"C:\elsewhere\other.accdb", None) == "library"


def test_no_code_path_reads_active_vb_project():
    pkg = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "mcp_access")
    offenders = []
    for fname in os.listdir(pkg):
        if not fname.endswith(".py"):
            continue
        with open(os.path.join(pkg, fname), encoding="utf-8") as f:
            for n, line in enumerate(f, 1):
                if line.strip().startswith("#"):
                    continue
                if re.search(r"\.ActiveVBProject\b", line):
                    offenders.append(f"{fname}:{n}")
    assert offenders == [], (
        "VBE.ActiveVBProject can be a referenced library; use "
        f"_get_vb_project(app): {offenders}")


def _blocks(text):
    errors = []
    _check_blocks_in_module("mod", text.split("\n"), errors)
    return errors


def test_loop_with_trailing_comment_closes_do():
    code = """Public Sub Ok()
    Do
        i = i + 1
    Loop  ' end of pass
    Do While i < 10
        i = i + 1
    Loop 'no space before the comment
End Sub
"""
    assert _blocks(code) == []


def test_wend_with_trailing_comment_closes_while():
    code = """Public Sub Ok()
    While i < 10
        i = i + 1
    Wend  ' done
End Sub
"""
    assert _blocks(code) == []


def test_identifier_starting_with_loop_does_not_close_do():
    code = """Public Sub Bad()
    Do
        LoopCount = LoopCount + 1
End Sub
"""
    errors = _blocks(code)
    assert errors and "Do" in errors[0]["error"]
