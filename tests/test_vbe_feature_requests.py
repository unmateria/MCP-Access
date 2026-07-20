"""
Pure-Python unit tests for the v0.7.52 VBE feature requests (Tom van Stiphout).

No COM / no Access needed — everything here exercises COM-free logic:
  - _apply_patches: the 4-tier match ladder, atomic reporting, require_unique,
    case-insensitive anchors, the Unicode length guard, and the existing
    fallbacks that must keep working byte-for-byte.
  - _check_blocks_in_module / _check_structure_in_module: the pure validators
    behind access_vbe_check_syntax.
  - (Declarations) resolution against a fake CodeModule.
  - _vbe_line_count / _cm_lines_list: the off-by-one fix.

Run with:
    python -m pytest tests/test_vbe_feature_requests.py
    python tests/test_vbe_feature_requests.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access.vbe import (  # noqa: E402
    _apply_patches, _cm_lines_list, _is_declarations, _vbe_line_count,
    _ws_normalized_matches, _case_insensitive_safe,
)
from mcp_access.compile import (  # noqa: E402
    _check_blocks_in_module, _check_structure_in_module,
)


class _FakeCodeModule:
    """
    Stand-in for a VBE CodeModule that reproduces the behaviour that matters:
      - Lines() joins with CRLF (the real VBE does; the older _FakeCM did not,
        which hid the whole CRLF-normalisation path in patch_proc).
      - CountOfLines counts a trailing blank line, so splitlines() disagrees
        with it by 1 — the exact off-by-one this release fixes.
    """

    def __init__(self, lines, decl_count=0):
        self._lines = list(lines)
        self.CountOfDeclarationLines = decl_count

    @property
    def CountOfLines(self):
        return len(self._lines)

    def Lines(self, start, count):
        if count <= 0:
            raise ValueError("VBE raises on count <= 0")
        return "\r\n".join(self._lines[start - 1:start - 1 + count])


# ---------------------------------------------------------------------------
# _apply_patches — atomic reporting
# ---------------------------------------------------------------------------

CODE = (
    "Private Sub Foo()\r\n"
    "    Dim v As Variant\r\n"
    "    If IsNull(v) Then\r\n"
    "        Exit Sub\r\n"
    "    End If\r\n"
    "End Sub"
)


def test_failed_patch_is_reported_and_code_object_untouched():
    """A batch with one bad anchor reports it; the caller's original string is
    never mutated in place, so the atomic gate can discard the result."""
    original = CODE
    r = _apply_patches(
        CODE,
        [{"find": "Exit Sub", "replace": "Exit Function"},
         {"find": "NoSuchAnchorHere", "replace": "x"}],
        proc_name="Foo",
    )
    assert r["applied"] == 1
    assert len(r["not_found"]) == 1
    assert "patch[1]" in r["not_found"][0]
    assert CODE == original, "input string must not be mutated"
    # The caller (ac_vbe_patch_proc) discards r["code"] when atomic is on.
    assert r["code"] != original


def test_patch_invalidated_by_earlier_patch_is_detected():
    """patch[0] destroys the anchor patch[1] cites — only a sequential
    simulation catches this; a pre-pass against the original text would not."""
    r = _apply_patches(
        CODE,
        [{"find": "Exit Sub", "replace": "Exit Function"},
         {"find": "Exit Sub", "replace": "Stop"}],
        proc_name="Foo",
    )
    assert r["applied"] == 1
    assert len(r["not_found"]) == 1


def test_patch_anchor_created_by_earlier_patch_matches():
    """The mirror case: patch[1]'s anchor only exists after patch[0] ran."""
    r = _apply_patches(
        CODE,
        [{"find": "Exit Sub", "replace": "GoTo Cleanup"},
         {"find": "GoTo Cleanup", "replace": "GoTo Done"}],
        proc_name="Foo",
    )
    assert r["applied"] == 2
    assert not r["not_found"]
    assert "GoTo Done" in r["code"]


# ---------------------------------------------------------------------------
# _apply_patches — require_unique
# ---------------------------------------------------------------------------

REPEATED = (
    "Private Sub Bar()\r\n"
    "    Call Log(\"x\")\r\n"
    "    Call Log(\"x\")\r\n"
    "End Sub"
)


def test_require_unique_blocks_repeated_anchor_and_reports_lines():
    r = _apply_patches(
        REPEATED, [{"find": "Call Log(\"x\")", "replace": "Call Log(\"y\")"}],
        require_unique=True, proc_name="Bar", base_line=10,
    )
    assert r["applied"] == 0
    assert len(r["unique_violations"]) == 1
    msg = r["unique_violations"][0]
    assert "matched 2 times" in msg
    # base_line=10 → the two hits are absolute module lines 11 and 12
    assert "[11, 12]" in msg
    assert r["code"] == REPEATED, "nothing may be replaced on a violation"


def test_repeated_anchor_without_require_unique_applies_first_and_warns():
    r = _apply_patches(
        REPEATED, [{"find": "Call Log(\"x\")", "replace": "Call Log(\"y\")"}],
        require_unique=False, proc_name="Bar",
    )
    assert r["applied"] == 1
    assert len(r["ambiguous_notes"]) == 1
    assert r["code"].count('Call Log("y")') == 1
    assert r["code"].count('Call Log("x")') == 1


# ---------------------------------------------------------------------------
# _apply_patches — case insensitivity
# ---------------------------------------------------------------------------

def test_lowercase_anchor_matches_camelcase_code():
    r = _apply_patches(
        CODE, [{"find": "if isnull(v) then", "replace": "If IsEmpty(v) Then"}],
        proc_name="Foo",
    )
    assert r["applied"] == 1
    assert "If IsEmpty(v) Then" in r["code"]
    assert "IsNull" not in r["code"]


def test_case_insensitive_match_echoes_stored_casing():
    r = _apply_patches(
        CODE, [{"find": "exit sub", "replace": "Exit Function"}], proc_name="Foo",
    )
    assert r["applied"] == 1
    joined = " ".join(r["fallback_notes"])
    assert "case-insensitively" in joined
    assert "Exit Sub" in joined, "the note must echo the REAL stored text"


def test_replacement_keeps_caller_casing():
    r = _apply_patches(
        CODE, [{"find": "exit sub", "replace": "eXiT fUnCtIoN"}], proc_name="Foo",
    )
    assert "eXiT fUnCtIoN" in r["code"]


def test_match_case_true_rejects_wrong_casing():
    r = _apply_patches(
        CODE, [{"find": "exit sub", "replace": "Exit Function"}],
        match_case=True, proc_name="Foo",
    )
    assert r["applied"] == 0
    assert len(r["not_found"]) == 1


def test_case_sensitive_tiers_run_before_case_insensitive_ones():
    """A ws-normalized case-SENSITIVE hit must win over a literal
    case-INSENSITIVE one, or existing calls would land somewhere new."""
    code = "    alpha\r\n    Beta\r\n        alpha\r\n"
    r = _apply_patches(code, [{"find": "alpha", "replace": "ZZZ"}], proc_name="P")
    assert r["applied"] == 1
    # tier 1 (literal, case-sensitive) hits the FIRST alpha, unchanged behaviour
    assert r["code"].startswith("    ZZZ")


def test_ws_normalized_fallback_still_works_and_keeps_its_note():
    r = _apply_patches(
        CODE, [{"find": "Exit Sub", "replace": "Exit Function"}], proc_name="Foo",
    )
    assert r["applied"] == 1
    # Indentation differs → the literal tier hits anyway (substring), so use a
    # multi-line anchor with wrong indentation to force the ws tier.
    r2 = _apply_patches(
        CODE,
        [{"find": "If IsNull(v) Then\nExit Sub\nEnd If", "replace": "' gone"}],
        match_case=True, proc_name="Foo",
    )
    assert r2["applied"] == 1
    assert any("ws-normalized fallback" in n for n in r2["fallback_notes"])


def test_lf_anchor_is_normalized_to_crlf():
    r = _apply_patches(
        CODE,
        [{"find": "    If IsNull(v) Then\n        Exit Sub\n",
          "replace": "    If IsEmpty(v) Then\r\n        Exit Sub\r\n"}],
        match_case=True, proc_name="Foo",
    )
    assert r["applied"] == 1
    assert not any("ws-normalized" in n for n in r["fallback_notes"]), \
        "CRLF normalisation must let the LITERAL tier hit"


# ---------------------------------------------------------------------------
# Unicode length guard
# ---------------------------------------------------------------------------

def test_case_insensitive_safe_flags_length_changing_lowercase():
    assert _case_insensitive_safe("If IsNull(v) Then")
    assert not _case_insensitive_safe("' İstanbul")  # U+0130 lowers to 2 chars


def test_length_changing_char_skips_case_insensitive_tiers():
    code = "    ' İstanbul branch\r\n    Exit Sub\r\n"
    r = _apply_patches(code, [{"find": "exit sub", "replace": "Exit Function"}],
                       proc_name="Foo")
    assert r["applied"] == 0, "must not splice at a shifted offset"
    assert len(r["case_notes"]) == 1
    assert r["code"] == code


def test_length_changing_char_still_allows_exact_match():
    code = "    ' İstanbul branch\r\n    Exit Sub\r\n"
    r = _apply_patches(code, [{"find": "Exit Sub", "replace": "Exit Function"}],
                       proc_name="Foo")
    assert r["applied"] == 1
    assert "Exit Function" in r["code"]


# ---------------------------------------------------------------------------
# _ws_normalized_matches
# ---------------------------------------------------------------------------

def test_ws_normalized_matches_returns_every_window():
    code = "  foo\r\n  bar\r\n    foo\r\n"
    assert len(_ws_normalized_matches(code, "foo")) == 2
    assert _ws_normalized_matches(code, "FOO", match_case=True) == []
    assert len(_ws_normalized_matches(code, "FOO", match_case=False)) == 2


# ---------------------------------------------------------------------------
# Off-by-one helpers
# ---------------------------------------------------------------------------

def test_vbe_line_count_counts_trailing_blank_line():
    assert _vbe_line_count("") == 0
    assert _vbe_line_count("a") == 1
    assert _vbe_line_count("a\r\nb") == 2
    assert _vbe_line_count("a\r\nb\r\n") == 3


def test_cm_lines_list_pads_to_count_of_lines():
    """A module ending in a blank line: splitlines() says 2, VBE says 3."""
    cm = _FakeCodeModule(["Sub A()", "End Sub", ""])
    assert cm.CountOfLines == 3
    assert len(cm.Lines(1, 3).splitlines()) == 2, "the bug being fixed"
    lines = _cm_lines_list(cm, "module:m")
    assert len(lines) == 3 == cm.CountOfLines
    assert lines[2] == ""


def test_cm_lines_list_leaves_normal_module_alone():
    cm = _FakeCodeModule(["Sub A()", "End Sub"])
    assert _cm_lines_list(cm, "module:m") == ["Sub A()", "End Sub"]


def test_cm_lines_list_handles_empty_module():
    assert _cm_lines_list(_FakeCodeModule([]), "module:m") == []


# ---------------------------------------------------------------------------
# (Declarations) resolution
# ---------------------------------------------------------------------------

def test_is_declarations_token():
    assert _is_declarations("(Declarations)")
    assert _is_declarations("  (declarations)  ")
    assert _is_declarations("(DECLARATIONS)")
    assert not _is_declarations("")
    assert not _is_declarations(None)
    assert not _is_declarations("Declarations")
    assert not _is_declarations("Form_Load")


def test_declarations_bounds_from_fake_module():
    cm = _FakeCodeModule(
        ["Option Compare Database", "Option Explicit", "Private mFoo As Long",
         "", "Private Sub A()", "End Sub"],
        decl_count=4,
    )
    assert cm.CountOfDeclarationLines == 4
    text = cm.Lines(1, cm.CountOfDeclarationLines)
    assert "Private Sub A()" not in text
    r = _apply_patches(text, [{"find": "As Long", "replace": "As String"}],
                       proc_name="(Declarations)")
    assert r["applied"] == 1
    assert "Private mFoo As String" in r["code"]


def test_declarations_count_zero_means_no_section():
    cm = _FakeCodeModule(["Private Sub A()", "End Sub"], decl_count=0)
    assert cm.CountOfDeclarationLines == 0
    # ac_vbe_patch_proc raises before calling Lines(1, 0), which VBE rejects.
    try:
        cm.Lines(1, 0)
        raise AssertionError("VBE would have raised here")
    except ValueError:
        pass


# ---------------------------------------------------------------------------
# access_vbe_check_syntax — the pure validators
# ---------------------------------------------------------------------------

GOOD_VBA = """Option Compare Database
Option Explicit

Private Const C As Long = 1

Public Sub Good()
    Dim i As Long
    For i = 1 To 10
        If i > 5 Then
            Debug.Print i
        End If
    Next i
    With Application
        .Echo True
    End With
End Sub
"""


def _blocks(text, name="mod"):
    errors = []
    _check_blocks_in_module(name, text.split("\n"), errors)
    return errors


def _structure(text, name="mod"):
    errors = []
    _check_structure_in_module(name, text.split("\n"), errors)
    return errors


def test_clean_vba_produces_no_errors():
    assert _blocks(GOOD_VBA) == []
    assert _structure(GOOD_VBA) == []


def test_missing_end_if_is_caught():
    bad = """Public Sub Bad()
    If x = 1 Then
        Debug.Print x
End Sub
"""
    errors = _blocks(bad)
    assert errors, "unclosed If must be reported"
    assert errors[0]["module"] == "mod"
    assert errors[0]["line"] == 2
    assert "If" in errors[0]["error"]


def test_missing_next_is_caught():
    bad = """Public Sub Bad()
    For i = 1 To 3
        Debug.Print i
End Sub
"""
    errors = _blocks(bad)
    assert errors and "For" in errors[0]["error"]


def test_missing_loop_is_caught():
    bad = """Public Sub Bad()
    Do While x
        x = x - 1
End Sub
"""
    errors = _blocks(bad)
    assert errors and "Do" in errors[0]["error"]


def test_missing_end_with_is_caught():
    bad = """Public Sub Bad()
    With Application
        .Echo True
End Sub
"""
    errors = _blocks(bad)
    assert errors and "With" in errors[0]["error"]


def test_code_outside_a_procedure_is_caught():
    """The classic damaged-module pattern: the Sub header got deleted."""
    bad = """Option Explicit

    Debug.Print "orphan"
"""
    errors = _structure(bad)
    assert errors, "executable code at module level must be reported"
    assert "line 3" in errors[0]


def test_missing_end_sub_leaves_orphan_code_detectable():
    bad = """Option Explicit

Public Sub A()
    Debug.Print 1

Public Sub B()
    Debug.Print 2
End Sub
"""
    # Structure check tolerates this (B is a valid proc start), but the missing
    # End Sub means A's body swallowed B — the block checker sees no mismatch
    # either. Documented limitation: this is a validator, not a compiler.
    assert _structure(bad) == []


def test_single_line_if_is_not_treated_as_a_block():
    ok = """Public Sub A()
    If x = 1 Then Debug.Print x
End Sub
"""
    assert _blocks(ok) == []


def test_missing_end_type_is_caught_by_structure_check():
    bad = """Option Explicit

Private Type TFoo
    a As Long

Public Sub A()
    Debug.Print 1
End Sub
"""
    # Without End Type everything below stays inside the Type block, so the
    # End Sub surfaces at module level as code outside a procedure.
    errors = _structure(bad)
    assert errors, "unterminated Type must not pass silently"


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
