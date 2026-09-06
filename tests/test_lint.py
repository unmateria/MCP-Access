"""
Pure-Python unit tests for the UI design lint engine (mcp_access.lint).

No COM / no Access needed — the rule functions operate on hand-built models
and on the real SaveAsText export grammar. Run with:

    python -m pytest tests/test_lint.py            (if pytest installed)
    python tests/test_lint.py                       (plain — runs all asserts)
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access import lint as L  # noqa: E402
from mcp_access.constants import LINT_RULES  # noqa: E402


# ---------------------------------------------------------------------------
# Colour decoding + contrast
# ---------------------------------------------------------------------------

def test_decode_bgr():
    assert L._decode_bgr(16777215) == {"r": 255, "g": 255, "b": 255,
                                       "system_theme": False, "valid": True}
    assert L._decode_bgr(0)["r"] == 0 and L._decode_bgr(0)["valid"]
    assert L._decode_bgr(255) == {"r": 255, "g": 0, "b": 0,
                                  "system_theme": False, "valid": True}
    assert L._decode_bgr(16711680) == {"r": 0, "g": 0, "b": 255,
                                       "system_theme": False, "valid": True}
    # System/theme colour: high bit set → not RGB-decodable
    sysc = L._decode_bgr(-2147483633)
    assert sysc["system_theme"] is True and sysc["valid"] is False
    assert L._decode_bgr(0x80000005)["system_theme"] is True


def test_contrast_ratio():
    black = L._decode_bgr(0)
    white = L._decode_bgr(16777215)
    assert round(L._contrast_ratio(black, white), 1) == 21.0
    assert round(L._contrast_ratio(white, white), 2) == 1.0
    # mid-grey on white is below the 4.5 AA threshold
    grey = L._decode_bgr(0x00B0B0B0)
    assert L._contrast_ratio(grey, white) < 4.5


# ---------------------------------------------------------------------------
# Style extraction (ignores nested ConditionalFormat colours)
# ---------------------------------------------------------------------------

def test_extract_style_ignores_nested():
    raw = "\n".join([
        "Begin TextBox",
        '    ForeColor =0',
        "    BackStyle =1",
        '    BackColor =16777215',
        '    FontName ="Calibri"',
        "    FontSize =11",
        "    ConditionalFormat = Begin",
        "        ForeColor =255",       # must NOT leak into top-level style
        "        BackColor =255",
        "    End",
        "    GUID = Begin",
        "        0xdeadbeef",
        "    End",
        "End",
    ])
    style = L._extract_style(raw)
    assert style["ForeColor"] == "0"          # top-level black, not the 255 inside CF
    assert style["BackColor"] == "16777215"
    assert style["FontName"] == "Calibri"
    assert style["BackStyle"] == "1"


# ---------------------------------------------------------------------------
# Geometry parsing
# ---------------------------------------------------------------------------

_FORM_TEXT = "\n".join([
    "Version =21",
    "Begin Form",
    "    Width =9360",
    "    BackColor =16777215",
    "    Begin FormHeader",
    "        Height =1440",
    '        Name ="FormHeader"',
    "    End",
    "    Begin Section",
    "        Height =5760",
    '        Name ="Detail"',
    "        BackColor =16777215",
    "        Begin Label",
    '            Name ="lbl1"',
    "        End",
    "    End",
    "    Begin FormFooter",
    "        Height =720",
    '        Name ="FormFooter"',
    "    End",
    "End",
])


def test_parse_geometry():
    geo = L._parse_geometry(_FORM_TEXT)
    assert geo["form_width"] == 9360
    assert geo["form_backcolor"] == 16777215
    names = {s["name"]: s["height"] for s in geo["sections"]}
    assert names == {"FormHeader": 1440, "Detail": 5760, "FormFooter": 720}
    # lbl1 is on line 13 (1-based) → inside Detail's range
    assert L._assign_section(13, geo["sections"]) == "Detail"


# ---------------------------------------------------------------------------
# Helpers to build a minimal model for rule tests
# ---------------------------------------------------------------------------

def _ctrl(name, type_name="TextBox", left=0, top=0, width=1000, height=300,
          section="Detail", parent="", visible="", caption="",
          control_source="", **style):
    return {"name": name, "type_name": type_name, "left": left, "top": top,
            "width": width, "height": height, "section": section,
            "parent": parent, "visible": visible, "caption": caption,
            "control_source": control_source, "style": {k: str(v) for k, v in style.items()}}


def _model(controls, form_width=9360, sections=None):
    sections = sections or [{"name": "Detail", "height": 5760, "backcolor": 16777215,
                             "begin_idx": 0, "end_idx": 9999}]
    return {
        "geometry": {"form_width": form_width, "form_backcolor": 16777215,
                     "object_kind": "Form", "sections": sections},
        "sections_by_name": {s["name"]: s for s in sections},
        "controls": controls,
    }


# ---------------------------------------------------------------------------
# Rules
# ---------------------------------------------------------------------------

def test_contrast_white_on_white():
    m = _model([_ctrl("lblBad", type_name="Label", caption="Hi",
                      ForeColor=16777215, BackStyle=1, BackColor=16777215)])
    v, theme, cf = L._rule_contrast(m)
    assert len(v) == 1 and v[0]["severity"] == "error" and v[0]["rule"] == "contrast"


def test_contrast_ok_black_on_white():
    m = _model([_ctrl("lblOk", type_name="Label", caption="Hi",
                      ForeColor=0, BackStyle=1, BackColor=16777215)])
    v, _, _ = L._rule_contrast(m)
    assert v == []


def test_contrast_skips_theme_colour():
    m = _model([_ctrl("lblTheme", type_name="Label", caption="Hi",
                      ForeColor=-2147483633, BackStyle=1, BackColor=16777215)])
    v, theme, cf = L._rule_contrast(m)
    assert v == [] and "lblTheme" in theme


def test_contrast_skips_conditional_format():
    # A control with conditional formatting can't be statically contrast-checked
    # (CF colours are binary in the export) — flagged in cf list, not a violation.
    c = _ctrl("lblCF", type_name="Label", caption="Hi",
              ForeColor=16777215, BackStyle=1, BackColor=16777215)
    c["has_conditional_format"] = True
    m = _model([c])
    v, theme, cf = L._rule_contrast(m)
    assert v == [] and "lblCF" in cf


def test_contrast_transparent_resolves_to_section_bg():
    # Transparent label (BackStyle=0) with white text over a white Detail bg → flagged.
    m = _model([_ctrl("lblT", type_name="Label", caption="Hi",
                      ForeColor=16777215, BackStyle=0)])
    v, _, _ = L._rule_contrast(m)
    assert len(v) == 1


def test_overlap_same_section():
    m = _model([
        _ctrl("a", left=0, top=0, width=1000, height=300),
        _ctrl("b", left=500, top=100, width=1000, height=300),
    ])
    v = L._rule_overlap(m)
    assert len(v) == 1 and v[0]["rule"] == "overlap"


def test_no_overlap_across_tab_pages():
    # Same coords but different parent (different Tab pages) → never overlap.
    m = _model([
        _ctrl("a", left=0, top=0, width=1000, height=300, parent="pageOne"),
        _ctrl("b", left=0, top=0, width=1000, height=300, parent="pageTwo"),
    ])
    assert L._rule_overlap(m) == []


def test_overlap_skips_transparent_layering():
    m = _model([
        _ctrl("a", type_name="Label", left=0, top=0, width=1000, height=300,
              BackStyle=0, BorderStyle=0),
        _ctrl("b", type_name="Label", left=100, top=50, width=1000, height=300,
              BackStyle=0, BorderStyle=0),
    ])
    assert L._rule_overlap(m) == []


def test_overlap_skips_transparent_button():
    # Transparent button stacked on a colored label = custom-button pattern.
    lbl = _ctrl("tileLbl", type_name="Label", left=0, top=0, width=1500, height=800,
                BackStyle=1, BackColor=2500000)
    btn = _ctrl("tileBtn", type_name="CommandButton", left=0, top=0, width=1500, height=800,
                Transparent=-1)
    assert L._rule_overlap(_model([lbl, btn])) == []


def test_out_of_bounds_horizontal():
    m = _model([_ctrl("wide", left=8400, top=100, width=1200, height=300)],
               form_width=9360)
    v = L._rule_out_of_bounds(m)
    assert len(v) == 1 and v[0]["measured"]["axis"] == "x" and v[0]["severity"] == "error"


def test_out_of_bounds_vertical():
    m = _model([_ctrl("tall", left=0, top=5600, width=500, height=400)])
    v = L._rule_out_of_bounds(m)
    assert len(v) == 1 and v[0]["measured"]["axis"] == "y"


def test_out_of_bounds_skips_parented():
    m = _model([_ctrl("inpage", left=8400, top=100, width=4000, height=300,
                      parent="pageOne")], form_width=9360)
    assert L._rule_out_of_bounds(m) == []


def test_truncation_heuristic():
    m = _model([_ctrl("lblLong", type_name="Label", width=400,
                      caption="A very very long caption that will not fit",
                      FontSize=11)])
    v = L._rule_truncation(m, None)
    assert len(v) == 1 and v[0]["rule"] == "truncation"


def test_truncation_multiline_label_not_flagged():
    # A tall label wraps text across lines; capacity = width x lines.
    long = "Some moderately long hint text that wraps nicely here"
    tall = _model([_ctrl("lblTall", type_name="Label", width=2000, height=1200,
                         caption=long, FontSize=11)])
    short = _model([_ctrl("lblShort", type_name="Label", width=2000, height=240,
                          caption=long, FontSize=11)])
    assert L._rule_truncation(tall, None) == []
    assert len(L._rule_truncation(short, None)) == 1


def test_truncation_skips_autosize_label():
    m = _model([_ctrl("lblAuto", type_name="Label", width=400,
                      caption="A very very long caption that will not fit",
                      FontSize=11, AutoSize=1)])
    assert L._rule_truncation(m, None) == []


def test_truncation_skips_bound_textbox():
    m = _model([_ctrl("txtData", type_name="TextBox", width=200,
                      control_source="SomeField")])
    assert L._rule_truncation(m, None) == []


def test_sibling_inconsistency_lone_outlier():
    # 4 buttons share 360; one lone 600 is the real oddball.
    m = _model([
        _ctrl("b1", type_name="CommandButton", top=0,    height=360, FontSize=11),
        _ctrl("b2", type_name="CommandButton", top=400,  height=360, FontSize=11),
        _ctrl("b3", type_name="CommandButton", top=800,  height=360, FontSize=11),
        _ctrl("b4", type_name="CommandButton", top=1200, height=600, FontSize=11),
    ])
    v = L._rule_sibling_inconsistency(m)
    assert len(v) == 1 and v[0]["control"] == "b4"


def test_sibling_two_legit_sizes_not_flagged():
    # A row of tall main buttons AND a row of short inline buttons: both shared
    # (>=2 each) → clustering accepts both, nothing flagged.
    m = _model([
        _ctrl("big1", type_name="CommandButton", top=0,   height=400, FontSize=11),
        _ctrl("big2", type_name="CommandButton", top=500, height=400, FontSize=11),
        _ctrl("big3", type_name="CommandButton", top=1000, height=400, FontSize=11),
        _ctrl("inl1", type_name="CommandButton", top=0,   height=300, FontSize=11),
        _ctrl("inl2", type_name="CommandButton", top=400, height=300, FontSize=11),
    ])
    assert L._rule_sibling_inconsistency(m) == []


def test_sibling_ignores_huge_multiline_outlier():
    # A memo box 4x the group height is a different control class, not an inconsistency.
    m = _model([
        _ctrl("t1", type_name="TextBox", top=0,    height=300, FontSize=11),
        _ctrl("t2", type_name="TextBox", top=400,  height=300, FontSize=11),
        _ctrl("t3", type_name="TextBox", top=800,  height=300, FontSize=11),
        _ctrl("tNotes", type_name="TextBox", top=1200, height=1440, FontSize=11),
    ])
    assert L._rule_sibling_inconsistency(m) == []


def test_zero_size_explicit_vs_absent():
    # Explicit height 0 → flagged; absent height (None) → inherits default, NOT flagged.
    m = _model([
        _ctrl("explicitZero", height=0),
        {**_ctrl("absentH"), "height": None},
    ])
    v = L._rule_invisible_or_zero_size(m)
    flagged = {x["control"] for x in v if x["severity"] == "warning"}
    assert flagged == {"explicitZero"}


def test_caption_lines_multiline():
    # SaveAsText encodes line breaks as literal \015\012.
    assert L._caption_lines(">\015\012") == [">"]
    assert L._caption_lines("IMPRIME\015\012ETIQUETAS") == ["IMPRIME", "ETIQUETAS"]
    assert L._caption_lines("Plain") == ["Plain"]


def test_truncation_multiline_button_not_flagged():
    # A 2-line button caption: only the longest line must fit the width.
    m = _model([_ctrl("btTwo", type_name="CommandButton", width=1400, height=600,
                      caption="IMPRIME\015\012ETIQUETAS", FontSize=11)])
    assert L._rule_truncation(m, None) == []
    # Same text on one line in the same width → too wide → flagged.
    m2 = _model([_ctrl("btOne", type_name="CommandButton", width=900, height=300,
                       caption="IMPRIME ETIQUETAS", FontSize=11)])
    assert len(L._rule_truncation(m2, None)) == 1


def test_strip_accelerator():
    assert L._strip_accelerator("&Info") == "Info"
    assert L._strip_accelerator("Save && Exit") == "Save & Exit"
    assert L._strip_accelerator("Plain") == "Plain"


def test_summary_verdict():
    assert L._summarize([{"severity": "error", "rule": "x"}])["verdict"] == "FAIL"
    assert L._summarize([{"severity": "warning", "rule": "x"}])["verdict"] == "REVIEW"
    assert L._summarize([])["verdict"] == "PASS"


def test_normalize_rules():
    assert L._normalize_rules(None) == list(LINT_RULES)
    assert L._normalize_rules(["contrast", "bogus"]) == ["contrast"]
    assert L._normalize_rules(["nope"]) == list(LINT_RULES)


# ---------------------------------------------------------------------------
# Plain runner (no pytest required)
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# tab_parent_hint — the unparented-control-on-a-tab warning
# ---------------------------------------------------------------------------

def _tab_model(**kw):
    """A tab control at (195,600)-(9200,5600) with two pages on it."""
    ctrls = [
        _ctrl("tabDetails", "Tab", left=195, top=600, width=9005, height=5000),
        _ctrl("pagGeneral", "Page", left=270, top=1065, width=8850, height=4455,
              parent="tabDetails"),
        _ctrl("pagLines", "Page", left=270, top=1065, width=8850, height=4455,
              parent="tabDetails"),
        _ctrl("btPrint", "CommandButton", left=1200, top=1500,
              width=2000, height=400, **kw),
    ]
    return _model(ctrls)


def test_tab_hint_fires_for_an_unparented_control_inside_a_tab():
    msg = L.tab_parent_hint(_tab_model(), "btPrint")
    assert msg is not None
    assert "tabDetails" in msg
    # The trap is that CreateControl wants the PAGE name — say so, and name them.
    assert "PAGE name" in msg
    assert "pagGeneral" in msg and "pagLines" in msg


def test_tab_hint_silent_when_the_control_is_parented():
    assert L.tab_parent_hint(_tab_model(parent="pagGeneral"), "btPrint") is None


def test_tab_hint_silent_when_the_control_is_outside_the_tab():
    m = _model([
        _ctrl("tabDetails", "Tab", left=195, top=600, width=9005, height=5000),
        _ctrl("btPrint", "CommandButton", left=200, top=6000, width=2000, height=400),
    ])
    assert L.tab_parent_hint(m, "btPrint") is None


def test_tab_hint_silent_when_only_partly_over_the_tab():
    """Straddling the edge is an overlap, and _rule_overlap already covers it."""
    m = _model([
        _ctrl("tabDetails", "Tab", left=195, top=600, width=9005, height=5000),
        _ctrl("btPrint", "CommandButton", left=8000, top=1500, width=3000, height=400),
    ])
    assert L.tab_parent_hint(m, "btPrint") is None


def test_tab_hint_ignores_a_tab_in_another_section():
    m = _model([
        _ctrl("tabDetails", "Tab", left=195, top=600, width=9005, height=5000,
              section="Detail"),
        _ctrl("btPrint", "CommandButton", left=1200, top=1500, width=2000,
              height=400, section="FormHeader"),
    ])
    assert L.tab_parent_hint(m, "btPrint") is None


def test_tab_hint_silent_without_full_geometry():
    """An absent dimension inherits the form default — unknown, not zero."""
    m = _model([
        _ctrl("tabDetails", "Tab", left=195, top=600, width=9005, height=5000),
        _ctrl("btPrint", "CommandButton", left=1200, top=1500, width=None, height=400),
    ])
    assert L.tab_parent_hint(m, "btPrint") is None
    m2 = _model([
        _ctrl("tabDetails", "Tab", left=195, top=600, width=None, height=5000),
        _ctrl("btPrint", "CommandButton", left=1200, top=1500, width=2000, height=400),
    ])
    assert L.tab_parent_hint(m2, "btPrint") is None


def test_tab_hint_handles_a_missing_control_and_no_tabs():
    assert L.tab_parent_hint(_tab_model(), "doesNotExist") is None
    assert L.tab_parent_hint(_model([_ctrl("btPrint", "CommandButton")]),
                             "btPrint") is None


def test_tab_hint_matches_the_name_case_insensitively():
    assert L.tab_parent_hint(_tab_model(), "BTPRINT") is not None


def test_tab_hint_reports_no_pages_gracefully():
    m = _model([
        _ctrl("tabDetails", "Tab", left=195, top=600, width=9005, height=5000),
        _ctrl("btPrint", "CommandButton", left=1200, top=1500, width=2000, height=400),
    ])
    msg = L.tab_parent_hint(m, "btPrint")
    assert msg is not None and "Pages on it" not in msg


if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_") and callable(v)]
    passed = 0
    for fn in fns:
        try:
            fn()
            print(f"PASS  {fn.__name__}")
            passed += 1
        except AssertionError as e:
            print(f"FAIL  {fn.__name__}: {e}")
        except Exception as e:
            print(f"ERROR {fn.__name__}: {type(e).__name__}: {e}")
    print(f"\n{passed}/{len(fns)} passed")
    sys.exit(0 if passed == len(fns) else 1)
