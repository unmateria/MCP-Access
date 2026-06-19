"""
Pure-Python tests for the deterministic auto-layout planner (build_form) and the
v0.7.45 layout-quality lint rules. No COM / no Access — they exercise the pure
geometry function and the rule functions against hand-built models.

    python -m pytest tests/test_build_form_layout.py
    python tests/test_build_form_layout.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access import design_defaults as D          # noqa: E402
from mcp_access.build_form import _plan_layout        # noqa: E402
from mcp_access import lint as L                       # noqa: E402


# ---------------------------------------------------------------------------
# design_defaults
# ---------------------------------------------------------------------------

def test_bgr_palette():
    assert D.bgr(255, 255, 255) == 16777215
    assert D.bgr(0, 0, 0) == 0
    assert D.bgr(255, 0, 0) == 255          # pure red → b=0,g=0,r=255
    assert D.bgr(0, 0, 255) == 16711680     # pure blue
    # Documented palette values (must match tips('layout') and the writeup).
    assert D.PALETTE["form_bg"] == 16119285
    assert D.PALETTE["field_border"] == 13421772
    assert D.PALETTE["text"] == 3355443
    assert D.PALETTE["accent"] == 15426341


def test_snap():
    assert D.snap(0) == 0
    assert D.snap(113) == 120
    assert D.snap(89) == 60
    assert D.snap(30) == 60 or D.snap(30) == 0   # 30 is exactly half a dot
    assert D.snap("240") == 240
    assert D.snap(None) == 0
    assert all(D.snap(x) % D.GRID == 0 for x in (1, 57, 301, 999, 4801))


# ---------------------------------------------------------------------------
# _plan_layout — single column
# ---------------------------------------------------------------------------

def test_plan_single_counts_and_grid():
    plan = _plan_layout(
        fields=["Nombre", {"field": "Provincia", "control": "combobox"},
                {"field": "Notas", "control": "memo"}],
        actions=[{"caption": "Guardar", "on_click": "[Event Procedure]"}, "Cerrar"],
        title="Ficha", layout="single", theme="light",
    )
    ctrls = plan["controls"]
    # 3 fields → 3 labels + 3 fields; + 1 title + 2 buttons = 9
    assert len(ctrls) == 9
    roles = [c["role"] for c in ctrls]
    assert roles.count("label") == 3
    assert roles.count("field") == 3
    assert roles.count("title") == 1
    assert roles.count("button") == 2
    # Everything snapped to the grid.
    for c in ctrls:
        for k in ("left", "top", "width", "height"):
            assert c[k] % D.GRID == 0, (c["name"], k, c[k])
    # Tabbable = fields + buttons.
    assert sum(1 for c in ctrls if c.get("tab")) == 5


def test_plan_single_width_formula():
    plan = _plan_layout(fields=["A", "B"], actions=[], title=None,
                        layout="single", theme="light")
    # 240 + 1800 + 120 + 2400 + 240
    assert plan["form_width"] == 4800
    assert not plan["need_header"] and not plan["need_footer"]


def test_plan_memo_taller():
    plan = _plan_layout(fields=[{"field": "Obs", "control": "memo"}],
                        actions=[], title=None, layout="single", theme="light")
    field = next(c for c in plan["controls"] if c["role"] == "field")
    assert field["height"] == D.MEMO_H


def test_plan_no_overlap_same_section():
    plan = _plan_layout(
        fields=["A", "B", "C", "D"],
        actions=["Ok", "Cancel"], title="T", layout="single", theme="light")
    by_sec: dict = {}
    for c in plan["controls"]:
        by_sec.setdefault(c["section"], []).append(c)
    for ctrls in by_sec.values():
        for i in range(len(ctrls)):
            for j in range(i + 1, len(ctrls)):
                a, b = ctrls[i], ctrls[j]
                ox = min(a["left"] + a["width"], b["left"] + b["width"]) - max(a["left"], b["left"])
                oy = min(a["top"] + a["height"], b["top"] + b["height"]) - max(a["top"], b["top"])
                assert not (ox > 4 and oy > 4), f"{a['name']} overlaps {b['name']}"


def test_plan_fields_within_bounds():
    plan = _plan_layout(fields=["A", "B", "C"], actions=["X"], title="T",
                        layout="single", theme="light")
    fw = plan["form_width"]
    for c in plan["controls"]:
        assert c["left"] >= 0 and c["top"] >= 0
        assert c["left"] + c["width"] <= fw, (c["name"], c["left"] + c["width"], fw)


def test_plan_binding_and_unbound():
    plan = _plan_layout(fields=["Nombre", {"field": "Total libre"}],
                        actions=[], title=None, layout="single", theme="light")
    fields = [c for c in plan["controls"] if c["role"] == "field"]
    assert fields[0]["props"].get("ControlSource") == "Nombre"
    # A name with a space isn't a plain column → left unbound.
    assert "ControlSource" not in fields[1]["props"]


def test_plan_two_column_rows():
    plan = _plan_layout(fields=["A", "B", "C"], actions=[], title=None,
                        layout="two-column", theme="light")
    fields = [c for c in plan["controls"] if c["role"] == "field"]
    # A and B share a row (same top), C drops to the next row, two left columns.
    tops = sorted(set(c["top"] for c in fields))
    assert len(tops) == 2
    lefts = sorted(set(c["left"] for c in fields))
    assert len(lefts) == 2  # two field columns


def test_plan_plain_theme_no_colours():
    plan = _plan_layout(fields=["A"], actions=[], title=None,
                        layout="single", theme="plain")
    field = next(c for c in plan["controls"] if c["role"] == "field")
    assert "BackColor" not in field["props"]
    assert "ForeColor" not in field["props"]


# ---------------------------------------------------------------------------
# Layout-quality lint rules (v0.7.45)
# ---------------------------------------------------------------------------

def _ctrl(name, type_name, left, top, width, height, **style):
    return {
        "name": name, "type_name": type_name,
        "left": left, "top": top, "width": width, "height": height,
        "visible": "", "caption": style.pop("caption", ""),
        "control_source": "", "parent": "",
        "section": style.pop("section", "Detail"),
        "section_kind": style.pop("section_kind", ""),
        "style": style, "has_picture": False, "has_conditional_format": False,
    }


def _model(controls, form_width=4800):
    return {
        "geometry": {"form_width": form_width, "form_backcolor": None,
                     "object_kind": "Form", "sections": []},
        "sections_by_name": {},
        "controls": controls,
    }


def test_grid_alignment():
    m = _model([_ctrl("ok", "TextBox", 240, 240, 2400, 300),
                _ctrl("bad", "TextBox", 113, 250, 2400, 300)])
    v = L._rule_grid_alignment(m)
    flagged = {x["control"] for x in v}
    assert "bad" in flagged and "ok" not in flagged
    assert all(x["severity"] == "info" for x in v)


def test_edge_margin():
    m = _model([_ctrl("hug", "TextBox", 0, 240, 2400, 300),
                _ctrl("fine", "TextBox", 240, 660, 2400, 300)])
    v = L._rule_edge_margin(m)
    flagged = {x["control"] for x in v}
    assert "hug" in flagged and "fine" not in flagged


def test_spacing_consistency():
    # 4 controls in one column; gaps 120,120,660 → the 4th is the outlier.
    col = [_ctrl(f"c{i}", "TextBox", 240, top, 2400, 300)
           for i, top in enumerate((240, 660, 1080, 2100))]
    v = L._rule_spacing_consistency(_model(col))
    assert any(x["control"] == "c3" for x in v)
    assert all(x["severity"] == "info" for x in v)


def test_spacing_consistency_even_passes():
    col = [_ctrl(f"c{i}", "TextBox", 240, 240 + i * 420, 2400, 300)
           for i in range(5)]
    assert L._rule_spacing_consistency(_model(col)) == []


def test_hierarchy_button_smaller_than_body():
    m = _model([
        _ctrl("t1", "TextBox", 240, 240, 2400, 300, FontSize="11"),
        _ctrl("t2", "TextBox", 240, 660, 2400, 300, FontSize="11"),
        _ctrl("btn", "CommandButton", 240, 1080, 1500, 360, FontSize="8"),
    ])
    v = L._rule_hierarchy(m)
    assert any(x["control"] == "btn" for x in v)


def test_hierarchy_no_explicit_font_no_flag():
    m = _model([
        _ctrl("t1", "TextBox", 240, 240, 2400, 300, FontSize="11"),
        _ctrl("btn", "CommandButton", 240, 1080, 1500, 360),  # inherits
    ])
    assert L._rule_hierarchy(m) == []


def test_new_rules_registered():
    from mcp_access.constants import LINT_RULES
    for r in ("grid_alignment", "spacing_consistency", "edge_margin", "hierarchy",
              "generic_font"):
        assert r in LINT_RULES
        assert r in L._RULE_FUNCS and L._RULE_FUNCS[r] is not None


# ---------------------------------------------------------------------------
# Design directions (v0.7.46): type scale, spacing/density, palette anti-drift
# and WCAG contrast — all pure, verified against the lint's own colour maths.
# ---------------------------------------------------------------------------

# The human-readable source of truth for every direction colour. The palette in
# design_defaults is built with bgr() from exactly these hexes; the anti-drift
# test recomputes bgr(hex) and asserts they still match.
_DIRECTION_HEX = {
    "despacho": {"form_bg": "#FBFAF7", "field_bg": "#FFFFFF", "text": "#1A1A2E",
                 "accent": "#0F766E", "field_border": "#94A3B8"},
    "panel":    {"form_bg": "#F1F5F9", "field_bg": "#FFFFFF", "text": "#1E293B",
                 "accent": "#1E293B", "field_border": "#CBD5E1", "accent2": "#B45309"},
    "archivo":  {"form_bg": "#F4F1EC", "field_bg": "#FFFFFF", "text": "#23211E",
                 "accent": "#7C4A33", "field_border": "#DAD3C6"},
}


def _hex_rgb(h):
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def test_type_scale_values():
    assert D.type_scale(11, 1.25) == {"caption": 9, "body": 11, "subhead": 14,
                                      "title": 17, "display": 21}
    assert D.type_scale(11, 1.2) == {"caption": 9, "body": 11, "subhead": 13,
                                     "title": 16, "display": 19}
    # Every step is a whole number of points.
    for v in D.type_scale(13, 1.25).values():
        assert isinstance(v, int)


def test_space_and_density_on_grid():
    assert all(v % D.GRID == 0 for v in D.SPACE.values())
    assert all(v % D.GRID == 0 for d in D.DENSITY.values() for v in d.values())
    # Legacy aliases keep their historic values (so old layouts are unchanged).
    assert (D.MARGIN_X, D.MARGIN_Y, D.GAP_LABEL, D.ROW_GAP, D.COL_GAP) == (
        240, 240, 120, 120, 360)
    # compact density == the historic light spacing.
    assert D.DENSITY["compact"] == {"margin": 240, "margin_y": 240, "row_gap": 120}


def test_directions_palette_anti_drift():
    """Each declared palette Long must equal bgr() recomputed from its hex."""
    for name, dr in D.DIRECTIONS.items():
        for key, val in dr["palette"].items():
            expected = D.bgr(*_hex_rgb(_DIRECTION_HEX[name][key]))
            assert val == expected, (name, key, val, expected)


def test_directions_contrast_wcag():
    """Verified with the lint's own WCAG maths: body text, the white band title,
    and the accent used as text on the paper all clear the threshold."""
    white = L._decode_bgr(D.bgr(255, 255, 255))
    for name, dr in D.DIRECTIONS.items():
        pal = dr["palette"]
        text = L._decode_bgr(pal["text"])
        field_bg = L._decode_bgr(pal["field_bg"])
        accent = L._decode_bgr(pal["accent"])
        paper = L._decode_bgr(pal["form_bg"])
        assert L._contrast_ratio(text, field_bg) >= 4.5, name      # body text
        assert L._contrast_ratio(white, accent) >= 4.5, name       # white title on band
        assert L._contrast_ratio(accent, paper) >= 3.0, name       # accent-as-text on paper
    # Panel's optional amber primary accent reads on white.
    amber = L._decode_bgr(D.DIRECTIONS["panel"]["palette"]["accent2"])
    assert L._contrast_ratio(white, amber) >= 4.5


def test_light_theme_regression():
    """theme='light' is byte-for-byte unchanged by the directions work."""
    plan = _plan_layout(
        fields=["Nombre", {"field": "Provincia", "control": "combobox"},
                {"field": "Notas", "control": "memo"}],
        actions=[{"caption": "Guardar", "on_click": "[Event Procedure]"}, "Cerrar"],
        title="Ficha", layout="single", theme="light")
    assert plan["form_width"] == 4800
    assert len(plan["controls"]) == 9
    field = next(c for c in plan["controls"] if c["role"] == "field")
    title = next(c for c in plan["controls"] if c["role"] == "title")
    assert field["props"]["FontName"] == "Calibri"
    assert title["props"]["FontName"] == "Calibri"   # no title_font for light
    # No direction palette/header band leaked in.
    assert not any(c["name"] == "recHeaderBand" for c in plan["controls"])


def test_direction_plan_uses_palette_and_title_font():
    for name in ("despacho", "panel", "archivo"):
        plan = _plan_layout(fields=["Nombre", "Provincia"],
                            actions=[{"caption": "Guardar cambios"}], title="Ficha",
                            layout="single", theme=name)
        pal = D.DIRECTIONS[name]["palette"]
        fonts = D.DIRECTIONS[name]["fonts"]
        title = next(c for c in plan["controls"] if c["role"] == "title")
        field = next(c for c in plan["controls"] if c["role"] == "field")
        band = next(c for c in plan["controls"] if c["name"] == "recHeaderBand")
        # Title in the display typeface, white, on its own accent band.
        assert title["props"]["FontName"] == fonts["display"]
        assert title["props"]["ForeColor"] == D.bgr(255, 255, 255)
        assert title["props"]["BackColor"] == pal["accent"]
        assert band["props"]["BackColor"] == pal["accent"]
        # Fields in the body typeface on the field colour.
        assert field["props"]["FontName"] == fonts["body"]
        assert field["props"]["BackColor"] == pal["field_bg"]
        assert field["props"]["BorderColor"] == pal["field_border"]
        # Title font is larger than the field font (clean hierarchy).
        assert title["props"]["FontSize"] > field["props"]["FontSize"]
        # Canvas is the direction's paper.
        assert plan["canvas"] == pal["form_bg"]


def test_panel_has_card_others_dont():
    assert any(c["name"] == "recCard"
               for c in _plan_layout(["A"], [], None, "single", "panel")["controls"])
    for name in ("despacho", "archivo"):
        assert not any(c["name"] == "recCard"
                       for c in _plan_layout(["A"], [], None, "single", name)["controls"])


def test_generic_font_rule():
    m = _model([_ctrl("a", "Label", 240, 240, 1800, 300, FontName="Arial"),
                _ctrl("b", "Label", 240, 660, 1800, 300, FontName="Segoe UI"),
                _ctrl("c", "TextBox", 240, 1080, 2400, 300, FontName="Times New Roman")])
    flagged = {x["control"] for x in L._rule_generic_font(m)}
    assert flagged == {"a", "c"}
    assert all(x["severity"] == "info" for x in L._rule_generic_font(m))
    # A direction's fonts never trip it.
    for name in ("despacho", "panel", "archivo"):
        for f in D.DIRECTIONS[name]["fonts"].values():
            assert f.strip().lower() not in L._GENERIC_FONTS, (name, f)


def test_type_hierarchy_header_title():
    # Title in the header at 9pt, body at 11pt → the title fails to lead.
    m = _model([
        _ctrl("t1", "TextBox", 240, 660, 2400, 300, FontSize="11"),
        _ctrl("t2", "TextBox", 240, 1080, 2400, 300, FontSize="11"),
        _ctrl("titulo", "Label", 240, 120, 3000, 400, FontSize="9",
              section="FormHeader", section_kind="FormHeader", caption="Ficha"),
    ])
    flagged = [x for x in L._rule_hierarchy(m) if x["control"] == "titulo"]
    assert flagged and flagged[0]["severity"] == "info"
    # A 17pt title clears it.
    m2 = _model([
        _ctrl("t1", "TextBox", 240, 660, 2400, 300, FontSize="11"),
        _ctrl("big", "Label", 240, 120, 3000, 400, FontSize="17",
              section="FormHeader", section_kind="FormHeader", caption="Ficha"),
    ])
    assert not any(x["control"] == "big" for x in L._rule_hierarchy(m2))


if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for fn in fns:
        fn()
        print(f"  ok  {fn.__name__}")
    print(f"\n{len(fns)} tests passed.")
