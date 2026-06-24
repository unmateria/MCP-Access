"""
UI design lint for Access forms/reports.

Deterministic, rules-based validation that turns invisible-to-an-LLM layout
defects into structured JSON the agent can act on:

  - contrast               WCAG 2.1 contrast ratio (white-on-white etc.)
  - overlap                axis-aligned rectangle intersection of siblings
  - out_of_bounds          control spills outside its section / form width
  - truncation             caption wider than the control that holds it
  - sibling_inconsistency  one button/label differs from the group norm
  - misalignment           control off the column its siblings share (info)
  - invisible_or_zero_size zero-area or hidden controls

The engine is pure-Python and reads the SaveAsText export that
``ac_get_code`` already produces (controls + sections + colours live in the
text), so a full lint needs **one** COM round-trip (the export) and never
opens Design view.

Text width for the truncation rule is estimated with a conservative
proportional-font heuristic by default. ``measure="wizhook"`` opts into exact
measurement via Access' undocumented ``WizHook.TwipsFromFont`` (requires a
compiled VBA project; falls back to the heuristic on any failure).

The same engine powers the *embedded* lint that the design-mutation tools
(`ac_set_control_props`, `ac_set_multiple_controls`, `ac_create_control`)
attach to their result — see ``lint_compact`` — so a broken layout surfaces
deterministically on every edit, without the agent having to ask.
"""

import math
import re
from typing import Any, Optional

from .constants import CTRL_TYPE, LINT_RULES, WIZHOOK_KEY
from .controls import _get_parsed_controls
from .design_defaults import GRID as _GRID, MARGIN_X as _MARGIN_X, snap as _snap

# Section begin tokens in the SaveAsText export. Detail is `Begin Section`;
# the others are named blocks. Each carries `Name =` and `Height =`.
_SECTION_BEGIN_RE = re.compile(
    r"^Begin\s+(FormHeader|FormFooter|PageHeader|PageFooter|"
    r"ReportHeader|ReportFooter|GroupHeader\d*|GroupFooter\d*|Section)\s*$",
    re.IGNORECASE,
)
_TOP_BEGIN_RE = re.compile(r"^Begin\s+(?:Form|Report)\s*$", re.IGNORECASE)
_BEGIN_ANY_RE = re.compile(r"^Begin\b")
_BEGIN_PROP_RE = re.compile(r"^\w+\s*=\s*Begin\s*$")

# Style properties we pull out of each control's raw_block.
_STYLE_KEYS = {
    "ForeColor", "BackColor", "BorderColor", "BackStyle", "BorderStyle",
    "FontName", "FontSize", "FontWeight", "FontItalic", "FontUnderline",
    "TextAlign", "AutoSize", "CanGrow", "CanShrink", "Transparent",
}

# Control types that hold static, design-time text (Caption) we can measure
# for truncation. Bound TextBox/ComboBox show runtime data — unknowable here.
_TEXT_CAPTION_TYPES = {"Label", "CommandButton", "ToggleButton"}

# Types excluded from contrast (no text, or themed/complex backgrounds).
_NO_CONTRAST_TYPES = {
    "Line", "Rectangle", "Image", "Subform", "PageBreak", "Page",
    "OptionGroup", "CommandButton", "ToggleButton", "WebBrowser",
    "CustomControl", "BoundObjectFrame", "ObjectFrame", "NavigationControl",
}

# Types that legitimately have a zero extent on one axis.
_ONE_AXIS_TYPES = {"Line", "PageBreak"}

# Container/layout types exempt from geometry rules: tab Pages stack on top of
# one another by design (so they "overlap"), and containers span large areas —
# none of that is a layout bug.
_LAYOUT_EXEMPT_TYPES = {"Page", "OptionGroup"}

_DEFAULT_PADDING_TWIPS = 120  # internal control padding allowance


# ---------------------------------------------------------------------------
# Small numeric / colour helpers (pure, unit-tested)
# ---------------------------------------------------------------------------

def _twips(val: Any, default: int = 0) -> int:
    """Coerce an export string like ``"1748"`` to int twips. Empty -> default."""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return int(val)
    s = str(val).strip()
    if not s:
        return default
    try:
        return int(s)
    except ValueError:
        try:
            return int(float(s))
        except ValueError:
            return default


def _twips_opt(val: Any) -> Optional[int]:
    """Like _twips but returns None when the property is ABSENT/empty.

    An omitted Left/Top/Width/Height in the export means the control inherits
    the form's default control dimensions — NOT zero. Distinguishing absent
    from an explicit 0 avoids false "zero-size"/bounds violations on controls
    that simply didn't serialise a dimension.
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return int(val)
    s = str(val).strip()
    if not s:
        return None
    try:
        return int(s)
    except ValueError:
        try:
            return int(float(s))
        except ValueError:
            return None


def _decode_bgr(raw: Any) -> dict:
    """Decode an Access colour Long (BGR order) to ``{r,g,b}``.

    System / theme colours have the high bit (0x80000000) set — these adapt to
    the Windows theme and CANNOT be decoded as RGB, so we flag them instead.
    """
    try:
        n = int(raw)
    except (TypeError, ValueError):
        return {"system_theme": False, "valid": False}
    v = n & 0xFFFFFFFF
    if v & 0x80000000:
        return {"system_theme": True, "valid": False, "raw": n}
    return {
        "r": v & 0xFF,
        "g": (v >> 8) & 0xFF,
        "b": (v >> 16) & 0xFF,
        "system_theme": False,
        "valid": True,
    }


def _relative_luminance(r: int, g: int, b: int) -> float:
    """WCAG 2.1 relative luminance from 8-bit sRGB channels."""
    def _lin(c: int) -> float:
        cs = c / 255.0
        return cs / 12.92 if cs <= 0.03928 else ((cs + 0.055) / 1.055) ** 2.4
    return 0.2126 * _lin(r) + 0.7152 * _lin(g) + 0.0722 * _lin(b)


def _contrast_ratio(rgb1: dict, rgb2: dict) -> float:
    """WCAG 2.1 contrast ratio between two decoded colours (1.0 .. 21.0)."""
    l1 = _relative_luminance(rgb1["r"], rgb1["g"], rgb1["b"])
    l2 = _relative_luminance(rgb2["r"], rgb2["g"], rgb2["b"])
    hi, lo = (l1, l2) if l1 >= l2 else (l2, l1)
    return (hi + 0.05) / (lo + 0.05)


def _heuristic_text_width_twips(text: str, font_size_pt: float,
                                bold: bool, font_name: str = "") -> int:
    """Conservative proportional-font width estimate, in twips.

    1 pt = 20 twips. Average glyph advance for a proportional UI font is
    ~0.50x the point size; monospace fonts ~0.60x. Deliberately *under*-counts
    a touch so truncation (a warning) is only raised when text clearly overflows.
    """
    if not text:
        return 0
    name = (font_name or "").lower()
    mono = any(k in name for k in ("courier", "consolas", "mono", "lucida console"))
    # Average glyph advance as a fraction of the point size. Narrow UI fonts
    # (Calibri/Tahoma/Segoe — the Access defaults) sit around 0.45; 0.46 keeps
    # the estimate from over-reporting truncation on labels that actually fit.
    ratio = 0.58 if mono else 0.46
    width = len(text) * font_size_pt * 20.0 * ratio
    if bold:
        width *= 1.06
    return int(width)


# ---------------------------------------------------------------------------
# Style + geometry extraction from the export text
# ---------------------------------------------------------------------------

def _extract_style(raw_block: str) -> dict:
    """Pull top-level style properties from a control's ``Begin..End`` block.

    Mirrors ``_parse_controls`` depth tracking so colours inside nested
    ``ConditionalFormat = Begin .. End`` (or GUID/NameMap) blocks are ignored —
    only properties at the control's own level (depth 1) are read.
    """
    style: dict = {}
    lines = raw_block.splitlines()
    depth = 0
    for idx, line in enumerate(lines):
        s = line.strip()
        if idx == 0:
            # The opening "Begin <Type>" line.
            depth = 1
            continue
        if depth == 1:
            m = re.match(r"^(\w+)\s*=(.*)", s)
            if m and m.group(1) in _STYLE_KEYS:
                val = m.group(2).strip().strip('"')
                if val and val != "Begin":
                    style[m.group(1)] = val
        if _BEGIN_ANY_RE.match(s) or _BEGIN_PROP_RE.match(s):
            depth += 1
        elif s == "End":
            depth -= 1
            if depth == 0:
                break
    return style


def _parse_geometry(form_text: str) -> dict:
    """Extract form width/back-colour and per-section height/back-colour/range.

    Returns::

        {
          "form_width": int,
          "form_backcolor": int|None,
          "object_kind": "Form"|"Report",
          "sections": [
            {"name": str, "height": int, "backcolor": int|None,
             "begin_idx": int, "end_idx": int}, ...
          ],
        }

    Section ``begin_idx``/``end_idx`` are 0-based line indices used to assign
    each control (by its 1-based ``start_line``) to the section it lives in.
    """
    lines = form_text.splitlines()
    geo: dict = {
        "form_width": 0, "form_backcolor": None,
        "object_kind": "Form", "sections": [],
    }

    # Locate Begin Form/Report and read top-level props before the first section.
    form_begin = -1
    for i, line in enumerate(lines):
        m = _TOP_BEGIN_RE.match(line.strip())
        if m:
            form_begin = i
            geo["object_kind"] = "Report" if line.strip().lower().endswith("report") else "Form"
            break
    if form_begin == -1:
        return geo

    # Form-level Width / BackColor: scan top-level props (depth 1) until the
    # first section block opens.
    depth = 0
    i = form_begin
    while i < len(lines):
        s = lines[i].strip()
        if i == form_begin:
            depth = 1
            i += 1
            continue
        if _SECTION_BEGIN_RE.match(s):
            break
        if depth == 1:
            m = re.match(r"^(\w+)\s*=\s*(-?\d+)\s*$", s)
            if m:
                if m.group(1) == "Width" and not geo["form_width"]:
                    geo["form_width"] = int(m.group(2))
                elif m.group(1) == "BackColor" and geo["form_backcolor"] is None:
                    geo["form_backcolor"] = int(m.group(2))
        if _BEGIN_ANY_RE.match(s) or _BEGIN_PROP_RE.match(s):
            depth += 1
        elif s == "End":
            depth -= 1
        i += 1

    # Walk section blocks, capturing each block's range + Height + BackColor + Name.
    i = form_begin + 1
    while i < len(lines):
        s = lines[i].strip()
        sm = _SECTION_BEGIN_RE.match(s)
        if sm:
            begin_idx = i
            # ``kind`` is the Begin token (FormHeader/FormFooter/Section/...) —
            # always present, unlike the optional ``Name =`` property. It lets a
            # rule tell a header from the detail body even when sections export
            # without a Name. Additive: never used by the pre-v0.7.46 rules.
            sec = {"name": "", "kind": sm.group(1), "height": 0, "backcolor": None,
                   "begin_idx": begin_idx, "end_idx": begin_idx}
            d = 1
            j = i + 1
            while j < len(lines):
                ss = lines[j].strip()
                if d == 1:
                    mh = re.match(r"^Height\s*=\s*(-?\d+)\s*$", ss)
                    mb = re.match(r"^BackColor\s*=\s*(-?\d+)\s*$", ss)
                    mn = re.match(r'^Name\s*=\s*"(.*)"\s*$', ss)
                    if mh and not sec["height"]:
                        sec["height"] = int(mh.group(1))
                    elif mb and sec["backcolor"] is None:
                        sec["backcolor"] = int(mb.group(1))
                    elif mn and not sec["name"]:
                        sec["name"] = mn.group(1)
                if _BEGIN_ANY_RE.match(ss) or _BEGIN_PROP_RE.match(ss):
                    d += 1
                elif ss == "End":
                    d -= 1
                    if d == 0:
                        sec["end_idx"] = j
                        break
                j += 1
            geo["sections"].append(sec)
            i = sec["end_idx"] + 1
            continue
        i += 1

    return geo


def _assign_section(start_line_1based: int, sections: list) -> str:
    """Return the name of the section whose line range contains the control."""
    idx = start_line_1based - 1
    best = ""
    for sec in sections:
        if sec["begin_idx"] < idx <= sec["end_idx"]:
            best = sec["name"] or best
    return best


def _assign_section_kind(start_line_1based: int, sections: list) -> str:
    """Return the *kind* (Begin token) of the section containing the control —
    e.g. ``FormHeader``/``FormFooter``/``Section``. Robust where ``Name`` is
    absent, used to distinguish the header band from the detail body."""
    idx = start_line_1based - 1
    best = ""
    for sec in sections:
        if sec["begin_idx"] < idx <= sec["end_idx"]:
            best = sec.get("kind") or best
    return best


# ---------------------------------------------------------------------------
# Build the lint model (one export, no Design view)
# ---------------------------------------------------------------------------

def _build_model(db_path: str, object_type: str, object_name: str) -> dict:
    """Parse controls + geometry from a single SaveAsText export.

    Reuses ``ac_get_code`` (binary sections stripped) and the cached
    ``_parse_controls`` result, then layers style/section onto each control.
    """
    from .code import ac_get_code  # lazy import (circular)
    text = ac_get_code(db_path, object_type, object_name)
    geo = _parse_geometry(text)
    parsed = _get_parsed_controls(db_path, object_type, object_name)

    sec_by_name = {s["name"]: s for s in geo["sections"]}
    controls: list[dict] = []
    for c in parsed["controls"]:
        name = (c.get("name") or "").strip()
        if not name:
            continue  # defaults block / unnamed — never a real control
        raw_block = c.get("raw_block", "")
        style = _extract_style(raw_block)
        type_name = c.get("type_name") or CTRL_TYPE.get(c.get("control_type", -1), "")
        # An icon/image button shows its Picture, not its Caption — so its
        # Caption must not be measured for truncation (it's often a leftover
        # default like "Comando111" or an unrelated label).
        has_picture = bool(re.search(r"^\s*Picture(Data)?\s*=\s*\S",
                                     raw_block, re.M))
        controls.append({
            "name": name,
            "type_name": type_name,
            # None means "absent in export" → inherits default, NOT zero.
            "left": _twips_opt(c.get("left")),
            "top": _twips_opt(c.get("top")),
            "width": _twips_opt(c.get("width")),
            "height": _twips_opt(c.get("height")),
            "visible": str(c.get("visible", "")).strip(),
            "caption": c.get("caption", ""),
            "control_source": c.get("control_source", ""),
            "parent": c.get("parent", ""),
            "section": _assign_section(c.get("start_line", 0), geo["sections"]),
            "section_kind": _assign_section_kind(c.get("start_line", 0), geo["sections"]),
            "style": style,
            "has_picture": has_picture,
            # Conditional formatting overrides ForeColor/BackColor at runtime,
            # and that data is BINARY in the export (unreadable here). So the
            # static contrast we can compute may not reflect what the user sees.
            "has_conditional_format": bool(c.get("format_conditions")),
        })
    return {"geometry": geo, "sections_by_name": sec_by_name, "controls": controls}


# ---------------------------------------------------------------------------
# Effective colour resolution
# ---------------------------------------------------------------------------

def _backstyle(style: dict, type_name: str) -> int:
    """Effective BackStyle: explicit if present, else per-type default.

    Labels default to Transparent (0); data controls default to Normal (1).
    """
    if "BackStyle" in style:
        return _twips(style["BackStyle"], 1)
    return 0 if type_name == "Label" else 1


def _effective_background(ctrl: dict, model: dict) -> Optional[dict]:
    """Resolve the colour a control's text sits on. None if not determinable."""
    style = ctrl["style"]
    type_name = ctrl["type_name"]
    bs = _backstyle(style, type_name)
    if bs == 0:
        # Transparent → whatever is behind it: section, then form.
        sec = model["sections_by_name"].get(ctrl["section"])
        if sec and sec.get("backcolor") is not None:
            return _decode_bgr(sec["backcolor"])
        fb = model["geometry"].get("form_backcolor")
        if fb is not None:
            return _decode_bgr(fb)
        return None
    if "BackColor" in style:
        return _decode_bgr(style["BackColor"])
    # Opaque control with no explicit BackColor: Access omits BackColor when it
    # equals the default (white). So an opaque text control with no BackColor
    # line renders on white — this is exactly how white-on-white slips through.
    if type_name in ("Label", "TextBox", "ComboBox", "ListBox"):
        return _decode_bgr(16777215)
    return None


# ---------------------------------------------------------------------------
# Geometry helpers
# ---------------------------------------------------------------------------

def _has_full_geom(c: dict) -> bool:
    """True only if left/top/width/height are all present and the control has
    a positive extent — the precondition for any geometry-based rule."""
    return (c["left"] is not None and c["top"] is not None
            and c["width"] is not None and c["height"] is not None
            and c["width"] > 0 and c["height"] > 0)


def _rect(c: dict) -> tuple:
    return (c["left"], c["top"], c["left"] + c["width"], c["top"] + c["height"])


def _overlap_area(a: dict, b: dict) -> tuple:
    ax1, ay1, ax2, ay2 = _rect(a)
    bx1, by1, bx2, by2 = _rect(b)
    ox = min(ax2, bx2) - max(ax1, bx1)
    oy = min(ay2, by2) - max(ay1, by1)
    return (ox, oy)


def _group_key(c: dict) -> tuple:
    return (c["section"], c.get("parent", ""))


def _is_transparent_no_border(c: dict) -> bool:
    style = c["style"]
    bs = _backstyle(style, c["type_name"])
    border = _twips(style.get("BorderStyle"), 0)
    return bs == 0 and border == 0


def _is_truthy(val) -> bool:
    """Access exports booleans as -1/1/True/Yes/NotDefault (vs absent/0/False)."""
    if val is None:
        return False
    return str(val).strip().lower() in ("-1", "1", "true", "yes", "notdefault")


def _is_transparent_button(c: dict) -> bool:
    """A Transparent command/toggle button is an invisible click layer placed
    over a styled Label/Rectangle (the standard Access custom-button pattern) —
    overlapping it is intentional, not a layout bug."""
    return (c["type_name"] in ("CommandButton", "ToggleButton")
            and _is_truthy(c["style"].get("Transparent")))


# ---------------------------------------------------------------------------
# Rules — each returns a list of violation dicts
# ---------------------------------------------------------------------------

def _v(rule, severity, control, detail, **extra) -> dict:
    out = {"rule": rule, "severity": severity, "control": control["name"],
           "control_type": control.get("type_name", ""),
           "section": control.get("section", ""), "detail": detail}
    if control.get("parent"):
        out["parent"] = control["parent"]
    out.update(extra)
    return out


def _rule_contrast(model: dict) -> tuple:
    """White-on-white & low-contrast text. Returns (violations, theme/cf notes)."""
    violations, theme_skipped, cf_skipped = [], [], []
    for c in model["controls"]:
        if c["type_name"] in _NO_CONTRAST_TYPES:
            continue
        # Needs text to be a contrast problem.
        if not c["caption"] and not c["control_source"]:
            continue
        # Conditional formatting can override the colours at runtime, and that
        # data is binary in the export — the static contrast is unreliable, so
        # don't assert a violation. Note it for the agent to verify manually.
        if c.get("has_conditional_format"):
            cf_skipped.append(c["name"])
            continue
        style = c["style"]
        if "ForeColor" not in style:
            continue
        fore = _decode_bgr(style["ForeColor"])
        if fore.get("system_theme"):
            theme_skipped.append(c["name"])
            continue
        if not fore.get("valid"):
            continue
        back = _effective_background(c, model)
        if back is None:
            continue
        if back.get("system_theme"):
            theme_skipped.append(c["name"])
            continue
        if not back.get("valid"):
            continue
        ratio = _contrast_ratio(fore, back)
        font_size = float(_twips(style.get("FontSize"), 11) or 11)
        bold = _twips(style.get("FontWeight"), 400) >= 700
        large = font_size >= 18 or (font_size >= 14 and bold)
        threshold = 3.0 if large else 4.5
        if ratio + 1e-9 < threshold:
            severity = "error" if ratio < 3.0 else "warning"
            violations.append(_v(
                "contrast", severity, c,
                f"Text contrast ratio {ratio:.2f}:1 is below the "
                f"{threshold:.1f}:1 minimum for {'large' if large else 'normal'} text.",
                measured={"fore": [fore["r"], fore["g"], fore["b"]],
                          "back": [back["r"], back["g"], back["b"]],
                          "ratio": round(ratio, 2), "font_size": font_size},
                expected={"min_ratio": threshold},
                suggested_fix=(
                    "Darken ForeColor (e.g. 0 = black) or change BackColor — "
                    "foreground and background are nearly identical."
                    if ratio < 1.5 else
                    "Increase contrast: pick a darker text or lighter background colour."),
            ))
    return violations, theme_skipped, cf_skipped


def _rule_overlap(model: dict) -> list:
    violations = []
    groups: dict = {}
    for c in model["controls"]:
        if not _has_full_geom(c) or c["type_name"] in _LAYOUT_EXEMPT_TYPES:
            continue
        groups.setdefault(_group_key(c), []).append(c)
    seen = set()
    for ctrls in groups.values():
        for i in range(len(ctrls)):
            for j in range(i + 1, len(ctrls)):
                a, b = ctrls[i], ctrls[j]
                ox, oy = _overlap_area(a, b)
                if ox > 4 and oy > 4:
                    # Both transparent & border-less → intentional layering.
                    if _is_transparent_no_border(a) and _is_transparent_no_border(b):
                        continue
                    # A Transparent button over a styled control is the standard
                    # custom-button pattern (invisible click layer) — not a bug.
                    if _is_transparent_button(a) or _is_transparent_button(b):
                        continue
                    # A frame (Rectangle) that *contains* the other isn't an overlap bug.
                    if _containment(a, b):
                        continue
                    key = tuple(sorted((a["name"], b["name"])))
                    if key in seen:
                        continue
                    seen.add(key)
                    area = min(a, b, key=lambda c: c["width"] * c["height"])
                    violations.append(_v(
                        "overlap", "warning", area,
                        f"'{a['name']}' and '{b['name']}' overlap by "
                        f"{ox}x{oy} twips in section '{a['section']}'.",
                        measured={"a": a["name"], "b": b["name"],
                                  "overlap_twips": [ox, oy]},
                        suggested_fix="Move or resize one control so their rectangles don't intersect.",
                    ))
    return violations


def _containment(a: dict, b: dict) -> bool:
    """True if one control is a frame fully containing the other (±20 twips)."""
    def inside(outer, inner):
        ox1, oy1, ox2, oy2 = _rect(outer)
        ix1, iy1, ix2, iy2 = _rect(inner)
        return (ox1 - 20 <= ix1 and oy1 - 20 <= iy1
                and ox2 + 20 >= ix2 and oy2 + 20 >= iy2)
    frames = {"Rectangle", "OptionGroup"}
    if a["type_name"] in frames and inside(a, b):
        return True
    if b["type_name"] in frames and inside(b, a):
        return True
    return False


def _rule_out_of_bounds(model: dict) -> list:
    violations = []
    form_width = model["geometry"].get("form_width", 0)
    for c in model["controls"]:
        # Parented controls (inside a Page/OptionGroup) use container-relative
        # coords — skip to avoid false positives.
        if c.get("parent") or c["type_name"] in _LAYOUT_EXEMPT_TYPES:
            continue
        if not _has_full_geom(c):
            continue
        left, top, right, bottom = _rect(c)
        if left < 0 or top < 0:
            violations.append(_v(
                "out_of_bounds", "error", c,
                f"Control has a negative position (left={left}, top={top}).",
                measured={"left": left, "top": top},
                suggested_fix="Set Left and Top to >= 0.",
            ))
            continue
        if form_width and right > form_width:
            violations.append(_v(
                "out_of_bounds", "error", c,
                f"Right edge {right} twips exceeds form width {form_width} by "
                f"{right - form_width}.",
                measured={"left": left, "width": c["width"], "right": right,
                          "bound": form_width, "axis": "x"},
                suggested_fix=f"Reduce Left to <= {max(0, form_width - c['width'])} "
                              f"or Width to <= {max(0, form_width - left)}.",
            ))
            continue
        sec = model["sections_by_name"].get(c["section"])
        sec_h = sec.get("height", 0) if sec else 0
        cangrow = _twips(c["style"].get("CanGrow"), 0)
        if sec_h and bottom > sec_h and not cangrow:
            violations.append(_v(
                "out_of_bounds", "error", c,
                f"Bottom edge {bottom} twips exceeds section '{c['section']}' "
                f"height {sec_h} by {bottom - sec_h}.",
                measured={"top": top, "height": c["height"], "bottom": bottom,
                          "bound": sec_h, "axis": "y"},
                suggested_fix=f"Reduce Top to <= {max(0, sec_h - c['height'])}, "
                              f"shrink Height, or grow the section.",
            ))
    return violations


def _strip_accelerator(caption: str) -> str:
    """Drop a single '&' accelerator marker; keep '&&' as a literal ampersand."""
    if "&" not in caption:
        return caption
    return caption.replace("&&", "\x00").replace("&", "").replace("\x00", "&")


# SaveAsText encodes embedded line breaks in captions as the literal escape
# sequences \015\012 (CR LF), \015 or \012 — a multi-line button/label caption.
_CAPTION_BREAKS = ("\\015\\012", "\\012\\015", "\\015", "\\012", "\r\n", "\n", "\r")


def _caption_lines(caption: str) -> list:
    """Split a caption into its display lines (explicit breaks), accelerator-
    stripped, empties dropped. ``">\\015\\012"`` → ``[">"]`` not one long blob."""
    if not caption:
        return []
    s = caption
    for tok in _CAPTION_BREAKS:
        s = s.replace(tok, "\n")
    return [ln for ln in (_strip_accelerator(p).strip() for p in s.split("\n")) if ln]


def _rule_truncation(model: dict, widths: Optional[dict]) -> list:
    violations = []
    for c in model["controls"]:
        if c["type_name"] not in _TEXT_CAPTION_TYPES:
            continue
        # Icon buttons display their Picture, not the Caption — don't measure.
        if c["type_name"] in ("CommandButton", "ToggleButton") and c.get("has_picture"):
            continue
        lines_text = _caption_lines(c["caption"])
        if not lines_text or not c["width"] or c["width"] <= 0:
            continue
        style = c["style"]
        # Labels that auto-size never truncate; multi-line that can grow neither.
        if _twips(style.get("AutoSize"), 0) == 1:
            continue
        if _twips(style.get("CanGrow"), 0) == 1:
            continue

        font_size = float(_twips(style.get("FontSize"), 11) or 11)
        bold = _twips(style.get("FontWeight"), 400) >= 700
        fontname = style.get("FontName", "")

        def _w(text: str) -> int:
            if widths and c["name"] in widths:
                return widths[c["name"]]   # WizHook measured the longest line
            return _heuristic_text_width_twips(text, font_size, bold, fontname)

        method = "wizhook" if (widths and c["name"] in widths) else "heuristic"
        avail = c["width"] - _DEFAULT_PADDING_TWIPS
        if avail <= 0:
            continue

        # Both Labels AND CommandButtons auto-wrap their caption across lines.
        # The box can DISPLAY `cap_lines` lines (from its height); the text
        # NEEDS however many lines it wraps into. Truncated only if it needs
        # more than fit. The 0.9 slack absorbs heuristic width error so labels
        # whose text *just* exceeds one line aren't falsely flagged.
        # Line box ≈ font point size × 1.2, in twips (1pt = 20 twips). A 540-twip
        # button shows 2 lines of 11pt text, so use round() (lenient) not floor —
        # under-counting display lines is the main truncation false-positive.
        line_h = font_size * 20 * 1.2
        cap_lines = max(1, round((c["height"] or 0) / line_h)) if c["height"] else 1
        # How many display lines each explicit line wraps into. The heuristic
        # over-estimates width by ~15-20% for narrow UI fonts (Calibri/Tahoma),
        # so a line only counts as overflowing when it exceeds the width by
        # >15% — that absorbs heuristic error (a header label that really fits
        # one line) while still catching genuinely long text (>1.15×).
        tol = 1.25 if method == "heuristic" else 1.02

        def _lines_for(tw: int) -> int:
            r = tw / avail
            return 1 if r <= tol else math.ceil(r)

        needed = sum(_lines_for(_w(ln)) for ln in lines_text)
        if needed > cap_lines:
            total_tw = sum(_w(ln) for ln in lines_text)
            violations.append(_v(
                "truncation", "warning", c,
                f"Caption text wraps to ~{needed} lines but only {cap_lines} fit "
                f"the control height ({method}).",
                measured={"needed_lines": needed, "fit_lines": cap_lines,
                          "text_width": total_tw, "control_width": c["width"],
                          "control_height": c["height"], "method": method},
                suggested_fix="Increase Height (to show more wrapped lines), widen "
                              "the control, or shorten the caption.",
            ))
    return violations


def _mode(values: list):
    """Most common value (ties → smallest), or None for empty."""
    if not values:
        return None
    counts: dict = {}
    for v in values:
        counts[v] = counts.get(v, 0) + 1
    best = max(counts.items(), key=lambda kv: (kv[1], -_sortable(kv[0])))
    return best[0]


def _sortable(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _accepted_clusters(values: list, tol: float) -> list:
    """Centroids of clusters that ≥2 controls share — the 'accepted' norms.

    A form legitimately uses more than one size (e.g. tall main buttons AND a
    row of short inline buttons). Only a value NOBODY else shares (a singleton
    far from every shared cluster) is a real inconsistency. Returns the list of
    (centroid, size) for clusters of size >= 2, largest first.
    """
    if not values:
        return []
    clusters: list = []  # list of [members]
    for v in sorted(values):
        if clusters and v - clusters[-1][-1] <= tol:
            clusters[-1].append(v)
        else:
            clusters.append([v])
    shared = [(sum(c) / len(c), len(c)) for c in clusters if len(c) >= 2]
    shared.sort(key=lambda cs: -cs[1])
    return shared


def _near_any(value: float, centroids: list, tol: float) -> bool:
    return any(abs(value - cen) <= tol for cen, _ in centroids)


def _rule_sibling_inconsistency(model: dict) -> list:
    violations = []
    # Group same-type controls within the same section+parent.
    groups: dict = {}
    for c in model["controls"]:
        if c["type_name"] not in ("CommandButton", "Label", "TextBox", "ComboBox"):
            continue
        groups.setdefault((c["type_name"],) + _group_key(c), []).append(c)

    for key, ctrls in groups.items():
        body = [c for c in ctrls if c["height"] is not None]
        if len(body) < 4:   # need a real population before talking about a "norm"
            continue

        heights = [c["height"] for c in body]
        sizes = [round(float(_twips(c["style"].get("FontSize"), 11) or 11), 1)
                 for c in body]
        fonts = [c["style"].get("FontName", "") for c in body
                 if c["style"].get("FontName")]

        # Accepted norms = values shared by >=2 controls (clustered). Two
        # legitimate sizes => two clusters => neither is flagged.
        h_clusters = _accepted_clusters(heights, tol=40)
        s_clusters = _accepted_clusters(sizes, tol=0.6)
        dominant_h = h_clusters[0][0] if h_clusters else None
        # Font-name norms: names used by >=2 controls.
        fn_counts: dict = {}
        for fn in fonts:
            fn_counts[fn] = fn_counts.get(fn, 0) + 1
        accepted_fonts = {fn for fn, n in fn_counts.items() if n >= 2}

        for c in body:
            issues = []
            h = c["height"]
            # Flag height only if there IS a shared norm, this control matches
            # none of the shared clusters, and it isn't a different class
            # (>2x the dominant size → memo/image box, intentional).
            if (h_clusters and not _near_any(h, h_clusters, 40)
                    and dominant_h and h <= dominant_h * 2):
                issues.append(f"height {h} vs group {int(dominant_h)}")
            fs = round(float(_twips(c["style"].get("FontSize"), 11) or 11), 1)
            if s_clusters and not _near_any(fs, s_clusters, 0.6):
                issues.append(f"font size {fs} vs group {s_clusters[0][0]:g}")
            fn = c["style"].get("FontName", "")
            if accepted_fonts and fn and fn not in accepted_fonts:
                issues.append(f"font '{fn}' vs group {sorted(accepted_fonts)}")
            if issues:
                violations.append(_v(
                    "sibling_inconsistency", "warning", c,
                    f"{c['type_name']} differs from its peers: " + "; ".join(issues) + ".",
                    measured={"height": h, "font_size": fs, "font_name": fn,
                              "group_height": int(dominant_h) if dominant_h else None,
                              "group_font_size": s_clusters[0][0] if s_clusters else None,
                              "accepted_fonts": sorted(accepted_fonts)},
                    suggested_fix="Match a peer value so controls look uniform "
                                  "(or this is an intentional one-off).",
                ))
    return violations


def _rule_misalignment(model: dict) -> list:
    violations = []
    groups: dict = {}
    for c in model["controls"]:
        if not _has_full_geom(c) or c["type_name"] in _LAYOUT_EXEMPT_TYPES:
            continue
        groups.setdefault(_group_key(c), []).append(c)
    for ctrls in groups.values():
        if len(ctrls) < 3:
            continue
        for axis, field in (("left", "left"), ("top", "top")):
            # Cluster shared edges; flag a near-miss to a cluster of >=3.
            vals = sorted(set(c[field] for c in ctrls))
            clusters = _cluster(sorted(c[field] for c in ctrls), tol=60)
            big = [cl for cl in clusters if len(cl) >= 3]
            for c in ctrls:
                for cl in big:
                    centroid = sum(cl) // len(cl)
                    d = abs(c[field] - centroid)
                    if 15 < d <= 60 and c[field] not in cl:
                        violations.append(_v(
                            "misalignment", "info", c,
                            f"{field.capitalize()} {c[field]} is {d} twips off the "
                            f"column at {centroid} shared by {len(cl)} controls.",
                            measured={field: c[field], "column": centroid, "off_by": d},
                            suggested_fix=f"Set {field.capitalize()} to {centroid} to align.",
                        ))
                        break
    return violations


def _cluster(sorted_vals: list, tol: int) -> list:
    clusters: list = []
    cur: list = []
    for v in sorted_vals:
        if cur and v - cur[-1] > tol:
            clusters.append(cur)
            cur = []
        cur.append(v)
    if cur:
        clusters.append(cur)
    return clusters


def _rule_invisible_or_zero_size(model: dict) -> list:
    violations = []
    for c in model["controls"]:
        hidden = c["visible"].lower() in ("0", "false", "no")
        # Only an EXPLICIT dimension of 0 is "zero-size". An absent dimension
        # (None) inherits the form default — not a defect.
        w, h = c["width"], c["height"]
        zero = (w is not None and w <= 0) or (h is not None and h <= 0)
        if c["type_name"] in _ONE_AXIS_TYPES:
            zero = (w is not None and w <= 0) and (h is not None and h <= 0)
        if zero and hidden:
            continue  # intentional: hidden + collapsed
        if zero:
            violations.append(_v(
                "invisible_or_zero_size", "warning", c,
                f"Control has no rendered area (width={c['width']}, height={c['height']}).",
                measured={"width": c["width"], "height": c["height"]},
                suggested_fix="Give the control a positive Width and Height.",
            ))
        elif hidden:
            violations.append(_v(
                "invisible_or_zero_size", "info", c,
                "Control is set to Visible = False.",
                measured={"visible": False},
                suggested_fix="Confirm this control is meant to be hidden at startup.",
            ))
    return violations


# ---------------------------------------------------------------------------
# Layout-quality rules (v0.7.45). All emit at most `info` — they enrich the
# full report but never flip the verdict nor reach the compact embedded lint.
# Deterministic and conservative: each one keys off the canonical design grid
# / margin (mcp_access.design_defaults) the build_form engine produces, so a
# form laid out by that engine passes them clean.
# ---------------------------------------------------------------------------

def _has_pos(c: dict) -> bool:
    """left/top present (the precondition for position-only rules)."""
    return c["left"] is not None and c["top"] is not None


def _rule_grid_alignment(model: dict) -> list:
    """Controls whose Left/Top don't sit on the 60-twip design grid (info)."""
    violations = []
    for c in model["controls"]:
        if c.get("parent") or c["type_name"] in _LAYOUT_EXEMPT_TYPES:
            continue
        if not _has_pos(c):
            continue
        left, top = c["left"], c["top"]
        off_l, off_t = left % _GRID, top % _GRID
        if not off_l and not off_t:
            continue
        axes = []
        if off_l:
            axes.append(f"Left {left}→{_snap(left)}")
        if off_t:
            axes.append(f"Top {top}→{_snap(top)}")
        violations.append(_v(
            "grid_alignment", "info", c,
            f"Off the {_GRID}-twip design grid ({', '.join(axes)}).",
            measured={"left": left, "top": top, "grid": _GRID},
            suggested_fix=f"Snap Left/Top to the nearest multiple of {_GRID} "
                          f"(e.g. Left={_snap(left)}, Top={_snap(top)}).",
        ))
    return violations


def _rule_edge_margin(model: dict) -> list:
    """Top-level controls hugging the form's left/right edge (info)."""
    violations = []
    form_width = model["geometry"].get("form_width", 0)
    for c in model["controls"]:
        if c.get("parent") or c["type_name"] in _LAYOUT_EXEMPT_TYPES:
            continue
        if not _has_full_geom(c):
            continue
        left, _t, right, _b = _rect(c)
        if 0 <= left < _GRID:
            violations.append(_v(
                "edge_margin", "info", c,
                f"Control hugs the left edge (Left={left}); no breathing room.",
                measured={"left": left, "recommended_margin": _MARGIN_X},
                suggested_fix=f"Indent Left to ~{_MARGIN_X} twips for a margin.",
            ))
            continue
        if form_width and 0 < (form_width - right) < _GRID:
            violations.append(_v(
                "edge_margin", "info", c,
                f"Right edge {right} nearly touches the form width {form_width}.",
                measured={"right": right, "form_width": form_width,
                          "recommended_margin": _MARGIN_X},
                suggested_fix=f"Leave ~{_MARGIN_X} twips between the control and "
                              f"the form's right edge.",
            ))
    return violations


def _median(values: list) -> float:
    s = sorted(values)
    n = len(s)
    if not n:
        return 0.0
    mid = n // 2
    return s[mid] if n % 2 else (s[mid - 1] + s[mid]) / 2.0


def _rule_spacing_consistency(model: dict) -> list:
    """Uneven vertical gaps between controls stacked in the same column (info).

    Groups by section+parent, clusters by Left (tol 60) to recover visual
    columns, then within a column of >=4 controls flags any inter-control gap
    that deviates sharply from the column's median gap. Needs a real population
    and a positive median, so a form with consistent ROW_STRIDE never trips.
    """
    violations = []
    groups: dict = {}
    for c in model["controls"]:
        if not _has_full_geom(c) or c["type_name"] in _LAYOUT_EXEMPT_TYPES:
            continue
        groups.setdefault(_group_key(c), []).append(c)

    for ctrls in groups.values():
        # Recover columns by clustering Left edges (tol = one grid dot).
        by_left = sorted(ctrls, key=lambda c: c["left"])
        col: list = []
        columns: list = []
        for c in by_left:
            if col and c["left"] - col[-1]["left"] > _GRID:
                columns.append(col)
                col = []
            col.append(c)
        if col:
            columns.append(col)

        for column in columns:
            if len(column) < 4:
                continue
            ordered = sorted(column, key=lambda c: c["top"])
            pairs = []
            for a, b in zip(ordered, ordered[1:]):
                gap = b["top"] - (a["top"] + a["height"])
                if gap >= 0:                 # negatives are overlaps (own rule)
                    pairs.append((gap, b))
            gaps = [g for g, _ in pairs]
            if len(gaps) < 3:
                continue
            med = _median(gaps)
            if med <= 0:
                continue
            for gap, below in pairs:
                wide = gap > med * 1.8 and (gap - med) > _GRID
                tight = gap < med * 0.45 and (med - gap) > _GRID
                if wide or tight:
                    violations.append(_v(
                        "spacing_consistency", "info", below,
                        f"Vertical gap above this control is {gap} twips vs the "
                        f"column median of {int(med)} — spacing looks uneven.",
                        measured={"gap": gap, "median_gap": int(med)},
                        suggested_fix=f"Even out the spacing (≈{int(med)} twips "
                                      f"between controls in this column).",
                    ))
    return violations


def _header_title(model: dict) -> Optional[dict]:
    """The form's title: the largest-area Label in a header section. None if the
    object has no header or no titled label there."""
    cands = [c for c in model["controls"]
             if c["type_name"] == "Label"
             and c.get("section_kind") in ("FormHeader", "PageHeader", "ReportHeader")
             and _has_full_geom(c)]
    if not cands:
        return None
    return max(cands, key=lambda c: c["width"] * c["height"])


def _rule_hierarchy(model: dict) -> list:
    """Type hierarchy inversions, info-only (two narrow, objective checks):

    1. A clickable action (CommandButton) whose caption font is strictly smaller
       than the form's dominant body (TextBox/Label) size by more than a point.
    2. The header title (largest Label in the header) whose FontSize is not
       larger than the body text — a title that fails to lead the hierarchy.

    Both fire only on an *explicit* FontSize so an inherited default is never
    flagged; a form built by access_build_form (title 16-17pt > body 11pt) passes.
    """
    body_sizes = [round(float(_twips(c["style"].get("FontSize"), 11) or 11), 1)
                  for c in model["controls"]
                  if c["type_name"] in ("TextBox", "Label")]
    if not body_sizes:
        return []
    body = _mode(body_sizes)
    if body is None:
        return []
    violations = []
    for c in model["controls"]:
        if c["type_name"] != "CommandButton":
            continue
        if "FontSize" not in c["style"]:
            continue           # inherits — no explicit inversion to flag
        fs = round(float(_twips(c["style"]["FontSize"], 11) or 11), 1)
        if fs < body - 1:
            violations.append(_v(
                "hierarchy", "info", c,
                f"Action button text ({fs}pt) is smaller than the body text "
                f"({body:g}pt) — actions should read at least as large.",
                measured={"button_font": fs, "body_font": body},
                suggested_fix=f"Raise this button's FontSize to >= {body:g}pt.",
            ))

    title = _header_title(model)
    if title is not None and "FontSize" in title["style"]:
        ts = round(float(_twips(title["style"]["FontSize"], 11) or 11), 1)
        if ts <= body:
            violations.append(_v(
                "hierarchy", "info", title,
                f"Header title text ({ts:g}pt) is not larger than the body text "
                f"({body:g}pt) — the title should set a clear visual hierarchy.",
                measured={"title_font": ts, "body_font": body},
                suggested_fix=f"Raise the title FontSize above {body:g}pt "
                              f"(a build_form direction uses {int(body) + 5}-{int(body) + 6}pt).",
            ))
    return violations


# Fonts that read as a generic AI/office default — a closed list so the rule
# never second-guesses a deliberate typeface. Lower-cased for comparison.
_GENERIC_FONTS = {
    "arial", "roboto", "inter", "times new roman", "ms sans serif",
}


def _rule_generic_font(model: dict) -> list:
    """Controls set in a generic/templated default typeface (info).

    Conservative by design: fires only on a short closed list of fonts that read
    as "AI default" (Arial / Roboto / Inter / Times New Roman / MS Sans Serif),
    never on a font the designer might have chosen on purpose. A build_form
    direction (Segoe UI / Constantia / Cambria / Corbel) passes clean.
    """
    violations = []
    for c in model["controls"]:
        fn = c["style"].get("FontName", "")
        if fn and fn.strip().lower() in _GENERIC_FONTS:
            violations.append(_v(
                "generic_font", "info", c,
                f"Font '{fn}' reads as a generic default — a typeface with more "
                f"character makes the form feel intentional.",
                measured={"font_name": fn},
                suggested_fix="Pick a distinctive font (e.g. Segoe UI, Constantia, "
                              "Cambria, Corbel) — see access_tips('design') or a "
                              "build_form direction (despacho/panel/archivo).",
            ))
    return violations


_RULE_FUNCS = {
    "contrast": None,  # handled specially (returns notes)
    "overlap": _rule_overlap,
    "out_of_bounds": _rule_out_of_bounds,
    "truncation": None,  # needs widths
    "sibling_inconsistency": _rule_sibling_inconsistency,
    "misalignment": _rule_misalignment,
    "invisible_or_zero_size": _rule_invisible_or_zero_size,
    "grid_alignment": _rule_grid_alignment,
    "spacing_consistency": _rule_spacing_consistency,
    "edge_margin": _rule_edge_margin,
    "hierarchy": _rule_hierarchy,
    "generic_font": _rule_generic_font,
}

_SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}


# ---------------------------------------------------------------------------
# WizHook exact text measurement (opt-in)
# ---------------------------------------------------------------------------

def _measure_text_batch(db_path: str, items: list) -> dict:
    """Measure many captions in ONE COM round-trip via WizHook.TwipsFromFont.

    items: [{"name","text","font","size","weight","italic"}]. Returns
    {control_name: width_twips}. Raises on failure (caller falls back to the
    heuristic). Requires a compiled VBA project; uses a temp standard module.
    """
    from .core import _Session, _get_vb_project
    from .vba_exec import _invoke_app_run

    app = _Session.connect(db_path)
    proj = _get_vb_project(app)
    comp = None
    try:
        comp = proj.VBComponents.Add(1)  # vbext_ct_StdModule
        temp_name = comp.Name
        body = [
            "Public Function _mcp_measure() As String",
            "    On Error Resume Next",
            f"    WizHook.Key = {WIZHOOK_KEY}",
            "    Dim s As String, w As Long, h As Long",
        ]
        for it in items:
            text = (it["text"] or "").replace('"', '""')
            font = (it["font"] or "Calibri").replace('"', '""')
            size = int(it["size"]) or 11
            weight = int(it["weight"]) or 400
            ital = "True" if it["italic"] else "False"
            nm = it["name"].replace('"', '""')
            body.append("    w = 0: h = 0")
            body.append(
                f'    WizHook.TwipsFromFont "{font}", {size}, {weight}, '
                f'{ital}, False, 0, "{text}", 0, w, h')
            body.append(f'    s = s & "{nm}" & Chr(1) & w & Chr(2)')
        body.append("    _mcp_measure = s")
        body.append("End Function")
        comp.CodeModule.InsertLines(1, "\r\n".join(body) + "\r\n")

        raw = _invoke_app_run(app, f"{temp_name}._mcp_measure", [])
        result: dict = {}
        if raw:
            for entry in str(raw).split("\x02"):
                if "\x01" in entry:
                    nm, w = entry.split("\x01", 1)
                    try:
                        result[nm] = int(w)
                    except ValueError:
                        pass
        if not result:
            raise RuntimeError("WizHook returned no measurements")
        return result
    finally:
        if comp is not None:
            try:
                proj.VBComponents.Remove(comp)
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def _run_rules(model: dict, db_path: str, rules: list, measure: str) -> dict:
    violations: list = []
    notes: list = []
    measure_method = "none"

    # Truncation widths (optional WizHook batch).
    widths = None
    if "truncation" in rules and measure in ("auto", "wizhook"):
        items = []
        for c in model["controls"]:
            cap_lines = _caption_lines(c["caption"])
            if (c["type_name"] in _TEXT_CAPTION_TYPES and cap_lines
                    and (c["width"] or 0) > 0
                    and not (c["type_name"] in ("CommandButton", "ToggleButton")
                             and c.get("has_picture"))):
                st = c["style"]
                longest = max(cap_lines, key=len)
                items.append({
                    "name": c["name"], "text": longest,
                    "font": st.get("FontName", "Calibri"),
                    "size": _twips(st.get("FontSize"), 11) or 11,
                    "weight": _twips(st.get("FontWeight"), 400) or 400,
                    "italic": _twips(st.get("FontItalic"), 0) == 1,
                })
        if items:
            try:
                widths = _measure_text_batch(db_path, items)
                measure_method = "wizhook"
            except Exception as exc:
                if measure == "wizhook":
                    notes.append(f"WizHook measurement failed ({exc}); used heuristic.")
                measure_method = "heuristic"
        else:
            measure_method = "heuristic"
    elif "truncation" in rules:
        measure_method = "heuristic"

    for rule in rules:
        if rule == "contrast":
            v, theme, cf = _rule_contrast(model)
            violations.extend(v)
            if theme:
                uniq = sorted(set(theme))
                notes.append(
                    f"{len(uniq)} control(s) use theme/system colours; contrast not "
                    f"statically checkable: {uniq[:10]}")
            if cf:
                uniq = sorted(set(cf))
                notes.append(
                    f"{len(uniq)} control(s) have conditional formatting; runtime "
                    f"colours may differ, contrast not statically checkable: {uniq[:10]}")
        elif rule == "truncation":
            violations.extend(_rule_truncation(model, widths))
        else:
            fn = _RULE_FUNCS.get(rule)
            if fn:
                violations.extend(fn(model))

    return {"violations": violations, "notes": notes, "measure_method": measure_method}


def _summarize(violations: list) -> dict:
    by_sev = {"error": 0, "warning": 0, "info": 0}
    by_rule: dict = {}
    for v in violations:
        by_sev[v["severity"]] = by_sev.get(v["severity"], 0) + 1
        by_rule[v["rule"]] = by_rule.get(v["rule"], 0) + 1
    verdict = "FAIL" if by_sev["error"] else ("REVIEW" if by_sev["warning"] else "PASS")
    return {**by_sev, "total": len(violations), "by_rule": by_rule, "verdict": verdict}


def _sort_violations(violations: list) -> list:
    return sorted(violations, key=lambda v: (
        _SEVERITY_ORDER.get(v["severity"], 9), v["rule"], v["control"]))


def _normalize_rules(rules) -> list:
    if not rules:
        return list(LINT_RULES)
    valid = []
    for r in rules:
        rl = str(r).strip().lower()
        if rl in LINT_RULES:
            valid.append(rl)
    return valid or list(LINT_RULES)


# ---------------------------------------------------------------------------
# Public: full lint tool
# ---------------------------------------------------------------------------

def ac_lint_form(
    db_path: str, object_type: str, object_name: str, *,
    rules=None, measure: str = "auto", include_screenshot: bool = False,
    max_violations: int = 200,
) -> dict:
    """Deterministic UI lint of a form/report. Returns structured violations."""
    if object_type not in ("form", "report"):
        raise ValueError("ac_lint_form only accepts object_type 'form' or 'report'")
    rules = _normalize_rules(rules)

    model = _build_model(db_path, object_type, object_name)
    run = _run_rules(model, db_path, rules, measure)
    violations = _sort_violations(run["violations"])
    summary = _summarize(violations)

    # Cap: keep all errors/warnings, drop info first.
    truncated = 0
    if len(violations) > max_violations:
        keep = [v for v in violations if v["severity"] != "info"][:max_violations]
        if len(keep) < max_violations:
            keep += [v for v in violations if v["severity"] == "info"][
                :max_violations - len(keep)]
        truncated = len(violations) - len(keep)
        violations = keep

    geo = model["geometry"]
    result: dict = {
        "object_type": object_type,
        "object_name": object_name,
        "measure_method": run["measure_method"],
        "geometry": {
            "form_width": geo.get("form_width", 0),
            "sections": {s["name"]: s["height"] for s in geo["sections"] if s["name"]},
        },
        "summary": summary,
        "violations": violations,
        "notes": run["notes"],
        "screenshot_path": None,
    }
    if truncated:
        result["truncated"] = truncated

    if include_screenshot:
        try:
            from .ui import ac_screenshot
            shot = ac_screenshot(db_path, object_type=object_type,
                                 object_name=object_name)
            result["screenshot_path"] = shot.get("path")
        except Exception as exc:
            result["notes"].append(f"Screenshot failed: {exc}")

    return result


# ---------------------------------------------------------------------------
# Public: compact lint embedded in design-mutation tool results
# ---------------------------------------------------------------------------

def _violation_controls(v: dict) -> set:
    """All control names a violation refers to (its own control + an overlap's pair)."""
    names = {v.get("control")}
    measured = v.get("measured")
    if isinstance(measured, dict):
        names.update(n for n in (measured.get("a"), measured.get("b")) if n)
    return {n for n in names if n}


def lint_compact(db_path: str, object_type: str, object_name: str,
                 max_items: int = 25,
                 focus_controls: Optional[set] = None) -> Optional[dict]:
    """Fast, embeddable lint for the design-mutation tools.

    Heuristic measurement (no WizHook/Run dependency — works on uncompiled
    projects), errors+warnings only, capped. Returns None on any failure so it
    can NEVER break the host mutation (the edit is the primary operation).

    focus_controls -- if set, the returned ``violations`` list is filtered to
    only those touching these controls (the ones the mutation just changed), so
    pre-existing issues on unrelated controls of a big inherited form don't
    drown out the result. The ``error``/``warning``/``info`` counts stay
    whole-form, so the caller still sees the form has other issues.
    """
    if object_type not in ("form", "report"):
        return None
    try:
        model = _build_model(db_path, object_type, object_name)
        run = _run_rules(model, db_path, list(LINT_RULES), measure="heuristic")
        actionable = [v for v in _sort_violations(run["violations"])
                      if v["severity"] in ("error", "warning")]
        summary = _summarize(run["violations"])
        focused = actionable
        filtered_out = 0
        if focus_controls:
            want = {c.casefold() for c in focus_controls}
            focused = [v for v in actionable
                       if {n.casefold() for n in _violation_controls(v)} & want]
            filtered_out = len(actionable) - len(focused)
        out = {
            "verdict": summary["verdict"],
            "error": summary["error"],
            "warning": summary["warning"],
            "info": summary["info"],
            "violations": focused[:max_items],
        }
        if filtered_out:
            out["hint"] = (
                f"{filtered_out} more error(s)/warning(s) on other controls hidden "
                "(pre-existing). Pass full_lint=true or run access_lint_form to see them."
            )
        elif focused:
            out["hint"] = "Run access_lint_form for the full report and suggested fixes."
        else:
            out["hint"] = "No errors or warnings detected."
        return out
    except Exception:
        return None
