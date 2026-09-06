"""
Pure-Python unit tests for the v0.7.43 usability fixes.

No COM / no Access needed — these exercise the COM-free logic:
  - _sql_effective_prefix / _is_destructive: unclosed comments must fail
    CLOSED (a DELETE hidden behind an unclosed /* must still be detected).
  - ac_vbe_replace_lines: omitting start_line in single mode raises an
    actionable error instead of the cryptic "start_line 0 out of range".

Run with:
    python -m pytest tests/test_usability_fixes.py
    python tests/test_usability_fixes.py
"""

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access.sql import _sql_effective_prefix, _is_destructive  # noqa: E402
from mcp_access.vbe import ac_vbe_replace_lines  # noqa: E402


# ---------------------------------------------------------------------------
# _sql_effective_prefix — comment stripping, incl. unclosed comments
# ---------------------------------------------------------------------------

def test_plain_statement():
    assert _sql_effective_prefix("SELECT * FROM t").startswith("SELECT")


def test_line_comment_then_delete():
    assert _sql_effective_prefix("-- note\nDELETE FROM t").startswith("DELETE")


def test_block_comment_then_delete():
    assert _sql_effective_prefix("/* note */ DELETE FROM t").startswith("DELETE")


def test_stacked_comments():
    sql = "-- a\n/* b */\n  -- c\nDROP TABLE t"
    assert _sql_effective_prefix(sql).startswith("DROP")


def test_unclosed_line_comment_is_inert():
    # A -- comment with no newline has nothing executable after it.
    assert _sql_effective_prefix("-- only a comment") == ""


def test_unclosed_block_comment_fails_closed():
    # The old code returned "" here, so the DELETE slipped past the
    # destructive guard. Must classify on the remaining text instead.
    sql = "/* unclosed\nDELETE FROM t"
    assert _sql_effective_prefix(sql).startswith("UNCLOSED")
    assert _is_destructive(sql) is False  # 'unclosed' is not a keyword...
    sql2 = "/*\nDELETE FROM t"
    assert _sql_effective_prefix(sql2).startswith("DELETE")
    assert _is_destructive(sql2) is True


def test_destructive_detection_still_works():
    assert _is_destructive("DELETE FROM t")
    assert _is_destructive("-- x\nDROP TABLE t")
    assert _is_destructive("SELECT a INTO t2 FROM t")
    assert not _is_destructive("SELECT a FROM t")


# ---------------------------------------------------------------------------
# ac_vbe_replace_lines — omitted start_line raises an actionable error
# (validation runs before any COM connection, so no Access needed)
# ---------------------------------------------------------------------------

def test_replace_lines_requires_start_line():
    with pytest.raises(ValueError, match="start_line is required"):
        ac_vbe_replace_lines("fake.accdb", "module", "mod1", new_code="x")


def test_replace_lines_negative_start_line():
    with pytest.raises(ValueError, match="start_line is required"):
        ac_vbe_replace_lines("fake.accdb", "module", "mod1",
                             start_line=0, count=1, new_code="x")


def test_access_tips_schema_lists_every_topic():
    """The tool description enumerates the topics by hand and had drifted twice
    (layout and design_vbe were missing). Keep it honest automatically."""
    from mcp_access.tips import _TIPS
    from mcp_access.tools import TOOLS
    desc = [t for t in TOOLS if t.name == "access_tips"][0].description
    listed = desc.split("Topics: ")[1].split(".")[0].split(", ")
    assert set(listed) == set(_TIPS), set(listed) ^ set(_TIPS)


def test_control_type_numbers_in_tips_match_the_real_map():
    """The 'controls' topic used to hand out numbers belonging to other controls
    (e.g. it called 106 a ComboBox — 106 is CheckBox)."""
    import re
    from mcp_access.tips import _TIPS
    from mcp_access.constants import CTRL_TYPE
    for num, name in re.findall(r"(\d{3})=([A-Za-z]+)", _TIPS["controls"]):
        if int(num) in CTRL_TYPE:
            assert CTRL_TYPE[int(num)] == name, f"{num} is {CTRL_TYPE[int(num)]}, not {name}"


def test_coord_keeps_the_string_minus_one_as_automatic():
    """coerce_prop maps "-1" to True (in Access -1 IS True) and int(True) is 1,
    so a client that serialises every argument as a string used to ask for the
    automatic position and get coordinate 1."""
    from mcp_access.controls import _coord
    from mcp_access.helpers import coerce_prop
    assert int(coerce_prop("-1")) == 1          # the trap this exists to avoid
    assert _coord("-1") == -1
    assert _coord(-1) == -1


def test_coord_handles_the_other_shapes_a_client_may_send():
    from mcp_access.controls import _coord
    assert _coord("1200") == 1200 and _coord(1200) == 1200
    assert _coord("2.0") == 2
    assert _coord(True) == -1 and _coord(False) == 0   # Access booleans: -1/0
    for junk in ("", "  ", "abc", None, [], {}):
        assert _coord(junk) == -1, junk


def test_geometry_args_no_longer_go_through_coerce_prop():
    """Structural guard: the four CreateControl geometry arguments and the
    snap_to_grid loop must use _coord."""
    import inspect
    from mcp_access import controls
    src = inspect.getsource(controls)
    assert "int(coerce_prop(_pop_ci" not in src
    assert src.count("_coord(_pop_ci") == 4


if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for fn in fns:
        fn()
        print(f"PASS {fn.__name__}")
    print(f"\n{len(fns)} tests passed")
