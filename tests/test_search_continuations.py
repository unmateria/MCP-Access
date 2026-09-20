"""
Pure-Python unit tests for continuation-aware search output (v0.7.60).

Feature request by Tom van Stiphout (@TvanStiphout-Home): a grep-style search
returns the matched PHYSICAL line, so a VBA statement wrapped with ' _' hides
everything on its continuation lines — most damagingly the return type of an
API Declare, which is what a 64-bit LongPtr audit hinges on.

No COM / no Access needed: _join_continuations, _continuation_index and
_add_continuation are pure functions over module text.

Run with:
    python -m pytest tests/test_search_continuations.py
    python tests/test_search_continuations.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access.vbe import (  # noqa: E402
    _add_continuation, _continuation_index, _join_continuations,
)


# The declaration from the field report, verbatim in shape: the wrong `As Long`
# return sits on line 3 and never appeared in the search output.
DECLARE_SRC = '''Option Explicit

Private Declare PtrSafe Function apiSelectObject Lib "gdi32" Alias "SelectObject" _
        (ByVal hdc As LongPtr, _
         ByVal hObject As LongPtr) As Long

Public Sub Plain()
    Debug.Print 1
End Sub
'''


def _lines(src):
    return src.splitlines()


def test_join_continuations_reports_first_and_last_line():
    logical = _join_continuations(_lines(DECLARE_SRC))
    by_first = {first: (last, text) for first, last, text in logical}
    assert 3 in by_first, "the Declare must start a logical statement at line 3"
    last, text = by_first[3]
    assert last == 5, f"chain must end on line 5, got {last}"
    assert text.endswith("As Long"), text
    assert "apiSelectObject" in text


def test_single_line_statement_has_last_equal_first():
    logical = _join_continuations(_lines(DECLARE_SRC))
    for first, last, text in logical:
        if text.strip() == "Option Explicit":
            assert last == first
            break
    else:
        raise AssertionError("Option Explicit not found")


def test_index_skips_single_line_statements():
    index = _continuation_index(_lines(DECLARE_SRC))
    # Only the three physical lines of the Declare belong to a chain.
    assert sorted(index) == [3, 4, 5], sorted(index)


def test_match_on_first_line_is_enriched():
    index = _continuation_index(_lines(DECLARE_SRC))
    match = {"line": 3, "content": _lines(DECLARE_SRC)[2]}
    _add_continuation(match, index)
    assert match["statement_line"] == 3
    assert match["end_line"] == 5
    assert "As LongPtr" in match["content_full"]
    assert match["content_full"].endswith("As Long")
    # The physical line stays exactly as it was: backwards compatible.
    assert match["content"].endswith("_")


def test_match_on_continuation_line_resolves_whole_statement():
    """The nice-to-have: a hit on any physical line returns the statement."""
    index = _continuation_index(_lines(DECLARE_SRC))
    match = {"line": 4, "content": _lines(DECLARE_SRC)[3]}
    _add_continuation(match, index)
    assert match["line"] == 4, "the hit line must not be rewritten"
    assert match["statement_line"] == 3
    assert match["end_line"] == 5
    assert match["content_full"].endswith("As Long")


def test_single_line_match_is_untouched():
    """The 99% case must produce byte-identical output to pre-0.7.60."""
    index = _continuation_index(_lines(DECLARE_SRC))
    match = {"line": 8, "content": "    Debug.Print 1"}
    before = dict(match)
    _add_continuation(match, index)
    assert match == before, match


def test_chain_of_many_continuations():
    src = 'x = a _\n  & b _\n  & c _\n  & d\n'
    index = _continuation_index(src.splitlines())
    assert sorted(index) == [1, 2, 3, 4]
    match = {"line": 3, "content": "  & c _"}
    _add_continuation(match, index)
    assert match["statement_line"] == 1
    assert match["end_line"] == 4
    assert match["content_full"] == "x = a & b & c & d"


def test_underscore_inside_identifier_is_not_a_continuation():
    src = 'Dim my_var As Long\nDim other As Long\n'
    assert _continuation_index(src.splitlines()) == {}


def test_trailing_comment_after_continuation_still_chains():
    src = "Declare Sub Foo Lib \"x\" _ ' keep going\n    (ByVal a As Long)\n"
    index = _continuation_index(src.splitlines())
    assert sorted(index) == [1, 2]
    match = {"line": 1, "content": src.splitlines()[0]}
    _add_continuation(match, index)
    # content_full drops trailing comments by design (it is built for judging
    # code, not for round-tripping source); the raw line still carries it.
    assert "keep going" not in match["content_full"]
    assert "keep going" in match["content"]


def test_dangling_continuation_at_end_of_module_does_not_hang():
    src = 'x = a _\n'
    logical = _join_continuations(src.splitlines())
    assert logical == [(1, 1, "x = a _")], logical


if __name__ == "__main__":
    passed = failed = 0
    for _name, _fn in sorted(list(globals().items())):
        if _name.startswith("test_") and callable(_fn):
            try:
                _fn()
                passed += 1
                print(f"  PASS  {_name}")
            except AssertionError as exc:
                failed += 1
                print(f"  FAIL  {_name}: {exc}")
    print(f"\n{passed} passed, {failed} failed")
    sys.exit(1 if failed else 0)
