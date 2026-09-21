"""
Pure-Python unit tests for wrapped property values in a SaveAsText export.

Covers:
  * join_wrapped_value / _parse_controls re-assembling a value Access split
    across 2 and 3 physical lines (the v0.7.61 bug: only the first fragment
    was returned, as a syntactically plausible half-expression);
  * a control WITHOUT any wrap producing exactly the v0.7.60 dict (no new
    keys, same values) — the non-regression guarantee;
  * decode_access_escapes surfacing caption_text only when there IS an escape;
  * _scan_control_properties: name resolved when `Name =` comes AFTER the
    property citing it, scope='form' for form-level properties, and nothing
    read from inside a nested `ConditionalFormat = Begin` block;
  * a search term that falls ACROSS the split being found;
  * ac_list_controls `fields` filtering (and rejecting an unknown field).

No COM / no Access needed. Run with:

    python -m pytest tests/test_control_property_wrap.py
    python tests/test_control_property_wrap.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access import controls as controls_mod                      # noqa: E402
from mcp_access.controls import (                                    # noqa: E402
    _parse_controls, _scan_control_properties, ac_list_controls,
)
from mcp_access.constants import CTRL_TYPE_BY_NAME                   # noqa: E402
from mcp_access.helpers import (                                     # noqa: E402
    join_wrapped_value, decode_access_escapes, text_matches,
)


# Shaped like a real Access 2016 export: long quoted values continued on their
# own indented lines, `Name =` after the property that cites it, a nested
# ConditionalFormat block, and a tab control with a page.
FORM_TEXT = "\n".join([
    "Version =21",
    "VersionRequired =20",
    "Begin Form",
    '    RecordSource ="qryOrdersOpen"',
    '    Filter ="[Status]=1 AND [Region]=\\04228\\042"',
    "    Width =9200",
    "    Begin",                                   # defaults block — no controls
    "        Begin TextBox",
    "            Width =1701",
    "        End",
    "    End",
    "    Begin Section",
    "        Height =5952",
    '        Name ="Detail"',
    "        Begin",
    "            Begin TextBox",
    "                Left =1200",
    "                Top =600",
    "                Width =2400",
    "                Height =300",
    # Split across TWO lines. Reading only the first gives "...*1" — an
    # expression that parses and multiplies by one instead of a hundred.
    '                ControlSource ="=FormatPercent(((Nz([Paid])+Nz([Credit])-Nz([Refund]))*1"',
    '                    "00)/Nz([Total])/100)"',
    '                Name ="txtMargin"',           # name AFTER the property
    "                ConditionalFormat = Begin",
    '                    ControlSource ="NOT_A_REAL_PROPERTY"',
    "                End",
    "            End",
    "            Begin ComboBox",
    "                Left =1200",
    "                Top =1000",
    # Split across THREE lines.
    '                RowSource ="SELECT tblCustomer.ID, tblCustomer.Name FROM tblCustomer W"',
    '                    "HERE tblCustomer.Active=True ORDER BY tblCustomer.Name, tblCu"',
    '                    "stomer.City"',
    '                Name ="cboCustomer"',
    "            End",
    "            Begin TextBox",
    "                Left =3600",
    "                Top =600",
    "                Width =1200",
    "                Height =300",
    '                ControlSource ="Total"',
    '                Name ="txtTotal"',
    "            End",
    "            Begin CommandButton",
    "                Left =6000",
    '                Caption ="Print\\015\\012invoice"',
    '                Name ="btPrint"',
    "            End",
    "            Begin Tab",
    '                Name ="tabDetails"',
    "                Begin",
    "                    Begin Page",
    '                        Name ="pagLines"',
    "                        Begin",
    "                            Begin ListBox",
    '                                RowSource ="SELECT * FROM tblLine"',
    '                                Name ="lstLines"',
    "                            End",
    "                        End",
    "                    End",
    "                End",
    "            End",
    "        End",
    "    End",
    "End",
])

FULL_MARGIN = ("=FormatPercent(((Nz([Paid])+Nz([Credit])-Nz([Refund]))*100)"
               "/Nz([Total])/100)")
FULL_ROWSOURCE = ("SELECT tblCustomer.ID, tblCustomer.Name FROM tblCustomer "
                  "WHERE tblCustomer.Active=True ORDER BY tblCustomer.Name, "
                  "tblCustomer.City")


def _by_name(parsed):
    return {c["name"]: c for c in parsed["controls"] if c["name"]}


# ---------------------------------------------------------------------------
# 1. Wrapped values come back whole
# ---------------------------------------------------------------------------

def test_two_line_wrap_is_joined():
    ctrl = _by_name(_parse_controls(FORM_TEXT))["txtMargin"]
    assert ctrl["control_source"] == FULL_MARGIN, ctrl["control_source"]
    # The half-value that used to be returned must NOT be what we get.
    assert not ctrl["control_source"].endswith("*1")


def test_three_line_wrap_is_joined():
    ctrl = _by_name(_parse_controls(FORM_TEXT))["cboCustomer"]
    scan = _scan_control_properties(FORM_TEXT, frozenset({"RowSource"}))
    got = [e for e in scan if e["control"] == "cboCustomer"]
    assert len(got) == 1, got
    assert got[0]["value"] == FULL_ROWSOURCE, got[0]["value"]
    assert ctrl["name"] == "cboCustomer"


def test_join_wrapped_value_without_continuation_is_a_plain_strip():
    lines = ['ControlSource ="Total"', '    Name ="txtTotal"']
    assert join_wrapped_value(lines, 0, '"Total"') == ("Total", 0)
    # A non-quoted value (a number) is untouched too.
    assert join_wrapped_value(["Left =1200"], 0, "1200") == ("1200", 0)


# ---------------------------------------------------------------------------
# 2. Non-regression: a control with no wrap is unchanged
# ---------------------------------------------------------------------------

def test_control_without_wrap_is_byte_identical_to_v0_7_60():
    ctrl = _by_name(_parse_controls(FORM_TEXT))["txtTotal"]
    # Exactly the keys v0.7.60 produced for a plain control: no caption_text,
    # no control_source_text, nothing else added for the 99% case.
    assert set(ctrl) == {
        "name", "control_type", "type_name", "caption", "control_source",
        "left", "top", "width", "height", "visible",
        "start_line", "end_line", "raw_block",
    }, sorted(ctrl)
    assert ctrl["name"] == "txtTotal"
    assert ctrl["control_type"] == CTRL_TYPE_BY_NAME["textbox"]
    assert ctrl["type_name"] == "TextBox"
    assert ctrl["caption"] == ""
    assert ctrl["control_source"] == "Total"
    assert (ctrl["left"], ctrl["top"], ctrl["width"], ctrl["height"]) == \
           ("3600", "600", "1200", "300")
    assert ctrl["visible"] == ""
    assert ctrl["raw_block"].startswith("            Begin TextBox")
    assert ctrl["raw_block"].rstrip().endswith("End")


# ---------------------------------------------------------------------------
# 3. Octal escapes
# ---------------------------------------------------------------------------

def test_caption_with_escapes_gains_caption_text():
    ctrl = _by_name(_parse_controls(FORM_TEXT))["btPrint"]
    assert ctrl["caption"] == "Print\\015\\012invoice"     # raw stays raw
    assert ctrl["caption_text"] == "Print\r\ninvoice"


def test_caption_without_escapes_adds_no_field():
    ctrl = _by_name(_parse_controls(FORM_TEXT))["txtTotal"]
    assert "caption_text" not in ctrl
    assert "control_source_text" not in ctrl


def test_decode_access_escapes_leaves_non_octal_alone():
    assert decode_access_escapes("a\\042b") == 'a"b'
    assert decode_access_escapes("path\\099x") == "path\\099x"
    assert decode_access_escapes("nothing here") == "nothing here"


# ---------------------------------------------------------------------------
# 4. _scan_control_properties
# ---------------------------------------------------------------------------

def test_scan_resolves_name_declared_after_the_property():
    scan = _scan_control_properties(FORM_TEXT, frozenset({"ControlSource"}))
    margin = [e for e in scan if e["value"] == FULL_MARGIN]
    assert len(margin) == 1, margin
    assert margin[0]["control"] == "txtMargin"
    assert margin[0]["control_type"] == "TextBox"
    assert margin[0]["scope"] == "control"
    # line is the physical line where the assignment starts (1-based)
    assert FORM_TEXT.splitlines()[margin[0]["line"] - 1].strip().startswith(
        "ControlSource =")


def test_scan_marks_form_level_properties():
    scan = _scan_control_properties(FORM_TEXT, frozenset({"RecordSource", "Filter"}))
    by_prop = {e["property"]: e for e in scan}
    assert by_prop["RecordSource"]["scope"] == "form"
    assert by_prop["RecordSource"]["control"] == ""
    assert by_prop["RecordSource"]["value"] == "qryOrdersOpen"
    assert by_prop["Filter"]["scope"] == "form"


def test_scan_ignores_nested_conditional_format_block():
    scan = _scan_control_properties(FORM_TEXT, None)
    assert not [e for e in scan if e["value"] == "NOT_A_REAL_PROPERTY"], scan


def test_scan_skips_the_defaults_block():
    # The leading `Begin` holds prototype controls with no Name — those are
    # not real controls and must not show up as nameless hits.
    scan = _scan_control_properties(FORM_TEXT, None)
    assert not [e for e in scan
                if e["scope"] == "control" and not e["control"]], scan


def test_scan_finds_controls_inside_a_tab_page():
    scan = _scan_control_properties(FORM_TEXT, frozenset({"RowSource"}))
    names = {e["control"] for e in scan}
    assert "lstLines" in names, names


# ---------------------------------------------------------------------------
# 5. A term that falls across the split
# ---------------------------------------------------------------------------

def test_search_term_crossing_the_split_is_found():
    needle = "*100)/Nz([Total])"
    # The physical-line scan the old code did cannot see it...
    assert not any(needle in line for line in FORM_TEXT.splitlines())
    # ...the joined value can.
    scan = _scan_control_properties(FORM_TEXT, None)
    hits = [e for e in scan
            if text_matches(needle, e["value"], False, False)]
    assert len(hits) == 1, hits
    assert hits[0]["control"] == "txtMargin"
    assert hits[0]["property"] == "ControlSource"


# ---------------------------------------------------------------------------
# 6. ac_list_controls(fields=...)
# ---------------------------------------------------------------------------

class _FakeParsed:
    """Swap _get_parsed_controls for the parsed fixture — no COM, no Access."""

    def __enter__(self):
        self._orig = controls_mod._get_parsed_controls
        parsed = _parse_controls(FORM_TEXT)
        controls_mod._get_parsed_controls = lambda *_a, **_k: parsed
        return parsed

    def __exit__(self, *_exc):
        controls_mod._get_parsed_controls = self._orig
        return False


def test_fields_filters_and_always_keeps_name():
    with _FakeParsed():
        out = ac_list_controls("x.accdb", "form", "frmOrders",
                               fields=["control_source"])
    assert out["count"] > 0
    for entry in out["controls"]:
        assert set(entry) == {"name", "control_source"}, entry
    margin = [c for c in out["controls"] if c["name"] == "txtMargin"][0]
    assert margin["control_source"] == FULL_MARGIN


def test_fields_none_returns_the_full_entry():
    with _FakeParsed():
        out = ac_list_controls("x.accdb", "form", "frmOrders")
    total = [c for c in out["controls"] if c["name"] == "txtTotal"][0]
    assert "left" in total and "start_line" in total
    assert "raw_block" not in total


def test_unknown_field_raises_and_lists_the_valid_ones():
    try:
        ac_list_controls("x.accdb", "form", "frmOrders", fields=["ctrl_source"])
    except ValueError as exc:
        assert "ctrl_source" in str(exc)
        assert "control_source" in str(exc)
    else:
        raise AssertionError("an unknown field must raise ValueError")


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
