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


if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for fn in fns:
        fn()
        print(f"PASS {fn.__name__}")
    print(f"\n{len(fns)} tests passed")
