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
        "control_source": "", "parent": "", "section": "Detail",
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
    for r in ("grid_alignment", "spacing_consistency", "edge_margin", "hierarchy"):
        assert r in LINT_RULES
        assert r in L._RULE_FUNCS and L._RULE_FUNCS[r] is not None


if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for fn in fns:
        fn()
        print(f"  ok  {fn.__name__}")
    print(f"\n{len(fns)} tests passed.")
