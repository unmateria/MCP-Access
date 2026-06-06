"""
Pure-Python unit tests for the v0.7.42 VBE fixes.

No COM / no Access needed — these exercise the COM-free logic:
  - _check_module_health Option-placement check (Fix B): a comment/blank header
    of any length must NOT trip the "Option statement misplaced" warning, but an
    Option line that follows real code still must.
  - _new_lines_to_code (Fix C): the new_lines→new_code alias normalisation.

Run with:
    python -m pytest tests/test_vbe_fixes.py     (if pytest installed)
    python tests/test_vbe_fixes.py                (plain — runs all asserts)

The blank-separator preservation (Fix A) and the new_lines end-to-end path are
covered by a live COM integration test (see the v0.7.42 changelog).
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access.vbe import _check_module_health  # noqa: E402
from mcp_access.dispatcher import _new_lines_to_code  # noqa: E402


class _FakeCM:
    """Minimal stand-in for a VBE CodeModule: just enough for _check_module_health."""

    def __init__(self, text: str):
        self._lines = text.splitlines()
        self.CountOfLines = len(self._lines)

    def Lines(self, start, count):  # 1-based, inclusive count — mirrors VBE
        return "\n".join(self._lines[start - 1:start - 1 + count])


def _option_warnings(text: str, name: str = "module:m"):
    return [w for w in _check_module_health(_FakeCM(text), name) if "Option" in w]


# ---------------------------------------------------------------------------
# Fix B — Option placement is comment-header-aware, not line-number-thresholded
# ---------------------------------------------------------------------------

def test_long_comment_header_no_warning():
    # _modTest case: a 6-line banner pushes Option Compare to line 7.
    text = (
        "' ------------------------------------------------------\n"
        "' Name:    _modTest\n"
        "' Kind:    Module\n"
        "' Purpose: Procedures for temporary developer testing only.\n"
        "' Out of scope: Production code.\n"
        "' Author:  dev\n"
        "' ------------------------------------------------------\n"
        "Option Compare Database\n"
        "Option Explicit\n"
        "\n"
        "Public Sub test1()\n"
        "    Dim m As Module\n"
        "End Sub\n"
    )
    assert _option_warnings(text) == []


def test_options_at_top_no_warning():
    text = "Option Compare Database\nOption Explicit\n\nSub Foo()\nEnd Sub\n"
    assert _option_warnings(text) == []


def test_multiple_option_family_lines_no_warning():
    # Option Base / Option Private must not be treated as "code" that would
    # then flag a following Option Explicit.
    text = "Option Compare Database\nOption Base 1\nOption Explicit\n\nSub Foo()\nEnd Sub\n"
    assert _option_warnings(text) == []


def test_option_after_code_warns():
    text = "Dim x As Long\nOption Explicit\n"
    w = _option_warnings(text)
    assert len(w) == 1 and "after executable code" in w[0]


def test_option_after_proc_warns():
    text = "Sub Foo()\nEnd Sub\nOption Explicit\n"
    assert any("after executable code" in w for w in _option_warnings(text))


# ---------------------------------------------------------------------------
# Fix C — new_lines → new_code alias normalisation
# ---------------------------------------------------------------------------

def test_new_lines_list_joins_with_newline():
    assert _new_lines_to_code(["", "Public Function X() As String"]) == \
        "\nPublic Function X() As String"
    assert _new_lines_to_code(["a", "b", "c"]) == "a\nb\nc"


def test_new_lines_empty_list_is_empty_string():
    assert _new_lines_to_code([]) == ""


def test_new_lines_json_string_is_parsed():
    # Some MCP clients serialise the array as a JSON-encoded string.
    assert _new_lines_to_code('["", "Public"]') == "\nPublic"


def test_new_lines_plain_string_passthrough():
    assert _new_lines_to_code("plain string") == "plain string"


def test_new_lines_none_is_none():
    assert _new_lines_to_code(None) is None


def test_new_lines_handles_none_items():
    assert _new_lines_to_code(["a", None, "b"]) == "a\n\nb"


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
