"""
Canonical design tokens for Access form layout.

The LLM is poor at emitting absolute twip coordinates blind — it overlaps
controls, spills out of bounds and invents incoherent colours/sizes. The fix
that works in practice (mirrored by v0/Vercel-style "design systems" and the
project's own lint philosophy) is to take the arithmetic away from the model:
give it a *closed* set of canonical sizes, spacings, fonts and colours, then
compute the geometry deterministically.

This module is the single source of truth for those tokens. ``build_form``
computes every coordinate from here; the ``lint`` rules ``grid_alignment`` /
``edge_margin`` reference :data:`GRID` / :data:`MARGIN_X`; ``tips('layout')``
documents the same numbers. Change a value here and the whole stack follows.

Units
-----
Everything is in **twips** (the Access form unit). 1 inch = 1440 twips,
1 cm = 567 twips, 1 point = 20 twips, 1 px@96dpi = 15 twips. Access' native
design grid is 24 subdivisions per inch → **60 twips** per grid dot, which is
why every value below is a multiple of 60.

Colours
-------
Access stores ForeColor/BackColor as a **BGR** Long (``b*65536 + g*256 + r``),
NOT RGB. :func:`bgr` builds one from r,g,b so the palette stays readable.
"""

# ---------------------------------------------------------------------------
# Unit grid
# ---------------------------------------------------------------------------
TWIPS_PER_INCH = 1440
TWIPS_PER_CM = 567
TWIPS_PER_POINT = 20

# Access' native design grid (24 dots/inch). All layout values snap to this.
GRID = 60

# ---------------------------------------------------------------------------
# Spacing scale (twips, all multiples of GRID)
# ---------------------------------------------------------------------------
# A small, closed spacing scale — the rhythm every margin and gap is drawn from,
# named like a CSS/design-system space scale so the intent is legible. Every
# value is a multiple of GRID (60) so layouts stay snapped.
SPACE: dict[str, int] = {
    "xs": 60, "sm": 120, "md": 240, "lg": 360, "xl": 480, "2xl": 600,
}

# Legacy spacing constants are now aliases into SPACE — same values, so existing
# layouts and tests are byte-identical. Prefer SPACE / DENSITY in new code.
MARGIN_X = SPACE["md"]      # 240 — left/right form margin
MARGIN_Y = SPACE["md"]      # 240 — top/bottom margin inside a section
GAP_LABEL = SPACE["sm"]     # 120 — horizontal gap between a label and its field
ROW_GAP = SPACE["sm"]       # 120 — vertical gap between rows
COL_GAP = SPACE["lg"]       # 360 — gap between columns in a two-column layout

# ---------------------------------------------------------------------------
# Standard control sizes (twips)
# ---------------------------------------------------------------------------
ROW_H = 300             # one-line control height (textbox/combo/label row)
LABEL_W = 1800          # default label width
FIELD_W = 2400          # default data-control width
MEMO_H = 960            # multi-line / memo control height (~3 lines)
CHECKBOX_W = 300        # the box itself (its label rides in the label column)

BUTTON_W = 1500         # comfortable click target (>= 1 inch wide)
BUTTON_H = 360
BUTTON_GAP = 120

ROW_STRIDE = ROW_H + ROW_GAP   # 420 — vertical advance per single-line row

# Section heights
HEADER_H = 720          # form-header band (holds the title)
FOOTER_H = 600          # form-footer band (holds the action buttons)

# ---------------------------------------------------------------------------
# Typography
# ---------------------------------------------------------------------------
BASE_FONT = "Calibri"
TITLE_FONT_SIZE = 16    # header title (rendered bold)
LABEL_FONT_SIZE = 10    # field labels
FIELD_FONT_SIZE = 11    # data controls
FONT_WEIGHT_NORMAL = 400
FONT_WEIGHT_BOLD = 700


def type_scale(base_pt: int = 11, ratio: float = 1.25) -> dict[str, int]:
    """A modular type scale rounded to whole points.

    One ``ratio`` (1.2 ≈ minor third, 1.25 ≈ major third) generates a coherent
    family of sizes from a single ``base_pt`` body size — the same idea as a CSS
    type scale, where every step is the previous one multiplied (or divided) by
    the ratio. ``type_scale(11, 1.25)`` → caption 9 / body 11 / subhead 14 /
    title 17 / display 21; ``type_scale(11, 1.2)`` → 9 / 11 / 13 / 16 / 19.

    A design direction picks the ratio; :func:`build_form._resolve_theme` maps
    ``title``→form title, ``caption``→labels, ``body``→data fields.
    """
    base = float(base_pt)
    return {
        "caption": round(base / ratio),
        "body":    round(base),
        "subhead": round(base * ratio),
        "title":   round(base * ratio ** 2),
        "display": round(base * ratio ** 3),
    }


# ---------------------------------------------------------------------------
# Colour helpers + closed palette
# ---------------------------------------------------------------------------

def bgr(r: int, g: int, b: int) -> int:
    """Build an Access colour Long from r,g,b (Access stores BGR order)."""
    return (b << 16) | (g << 8) | r


# A small, closed light-theme palette. The point is that the model picks FROM
# this list instead of inventing hex values that clash or fail contrast.
PALETTE: dict[str, int] = {
    "form_bg":      bgr(245, 245, 245),   # #F5F5F5 light grey canvas
    "field_bg":     bgr(255, 255, 255),   # #FFFFFF white data fields
    "field_border": bgr(204, 204, 204),   # #CCCCCC subtle borders
    "text":         bgr(51, 51, 51),      # #333333 body text + header title (not pure black)
    "accent":       bgr(37, 99, 235),     # #2563EB accent for highlights (not the header band:
                                          # Access themes the header gradient over a literal BackColor)
}

# Built-in colour aliases the model may pass by name (resolved by build_form).
COLOR_ALIASES: dict[str, int] = {
    "white": bgr(255, 255, 255),
    "black": bgr(0, 0, 0),
    "red":   bgr(255, 0, 0),
    "green": bgr(0, 128, 0),
    "blue":  bgr(0, 0, 255),
    "grey":  bgr(128, 128, 128),
    "gray":  bgr(128, 128, 128),
    **PALETTE,
}


# ---------------------------------------------------------------------------
# Density — margin / gap rhythm only (never control sizes)
# ---------------------------------------------------------------------------
# A direction picks a density; it scales the *air* (form margin, section margin,
# row gap) without ever touching a control's own dimensions, so the content
# stays the same legible size and only the breathing room changes. ``compact``
# reproduces the historic ``light`` spacing exactly.
DENSITY: dict[str, dict] = {
    "compact":     {"margin": 240, "margin_y": 240, "row_gap": 120},
    "comfortable": {"margin": 360, "margin_y": 480, "row_gap": 180},
    "spacious":    {"margin": 480, "margin_y": 600, "row_gap": 360},
}


# ---------------------------------------------------------------------------
# Curated design directions
# ---------------------------------------------------------------------------
# Three coherent, hand-tuned looks that translate real design-system thinking
# (a typeface with character ≠ Arial/Inter; a dominant + accent palette with
# verified WCAG contrast ≠ clichés; an intentional type scale; a spacing
# rhythm; containment) into what *native Access* can actually render.
#
# The honest ceiling: Access has no gradients, shadows, rounded corners, blur,
# animation or fine letter-spacing — none of that exists. What carries the look
# is the translatable layer: font, type scale + weights, a cohesive palette
# that passes contrast, spacing rhythm, density and containment. The result is
# tasteful *within the flat medium*, not web-grade.
#
# Palette keys are the canonical ones build_form already reads
# (form_bg / field_bg / field_border / text / accent); ``accent`` doubles as the
# header-band colour. Colours are built with bgr() straight from the hex so they
# cannot drift from the design intent (a test recomputes bgr(hex) and compares).
# Fonts are confirmed-installed Windows 11 / Office typefaces; the comment notes
# a fallback, but Access takes a single FontName so the primary is used.
DIRECTIONS: dict[str, dict] = {
    # Despacho — character: a serif title over warm paper, fields on the paper
    # (no card). Major-third scale, comfortable air.
    "despacho": {
        "fonts": {"body": "Segoe UI", "display": "Constantia"},  # fb: Tahoma / Georgia
        "scale": type_scale(11, 1.25),
        "density": "comfortable",
        "card": False,
        "palette": {
            "form_bg":      bgr(0xFB, 0xFA, 0xF7),   # #FBFAF7 warm paper
            "field_bg":     bgr(0xFF, 0xFF, 0xFF),   # #FFFFFF white fields
            "text":         bgr(0x1A, 0x1A, 0x2E),   # #1A1A2E ink   (16.3:1 on paper)
            "accent":       bgr(0x0F, 0x76, 0x6E),   # #0F766E teal band (white 5.47:1)
            "field_border": bgr(0x94, 0xA3, 0xB8),   # #94A3B8 slate
        },
    },
    # Panel — sober office: a semibold sans title and a white card floating on a
    # cool canvas. Minor-third scale, comfortable air.
    "panel": {
        "fonts": {"body": "Segoe UI", "display": "Segoe UI Semibold"},
        "scale": type_scale(11, 1.2),
        "density": "comfortable",
        "card": True,
        "palette": {
            "form_bg":      bgr(0xF1, 0xF5, 0xF9),   # #F1F5F9 cool canvas
            "field_bg":     bgr(0xFF, 0xFF, 0xFF),   # #FFFFFF card
            "text":         bgr(0x1E, 0x29, 0x3B),   # #1E293B slate  (14.6:1)
            "accent":       bgr(0x1E, 0x29, 0x3B),   # #1E293B slate band (white 14.6:1)
            "field_border": bgr(0xCB, 0xD5, 0xE1),   # #CBD5E1
            # Amber primary accent (white 5.02:1). Unused while flat_buttons is
            # False (buttons are native); kept for an opt-in primary button.
            "accent2":      bgr(0xB4, 0x53, 0x09),   # #B45309 amber
        },
    },
    # Archivo — warm editorial, generous space: a serif title, no card, the
    # widest rhythm. Major-third scale, spacious air.
    "archivo": {
        "fonts": {"body": "Corbel", "display": "Cambria"},
        "scale": type_scale(11, 1.25),
        "density": "spacious",
        "card": False,
        "palette": {
            "form_bg":      bgr(0xF4, 0xF1, 0xEC),   # #F4F1EC warm paper
            "field_bg":     bgr(0xFF, 0xFF, 0xFF),   # #FFFFFF white fields
            "text":         bgr(0x23, 0x21, 0x1E),   # #23211E ink   (14.3:1 on paper)
            "accent":       bgr(0x7C, 0x4A, 0x33),   # #7C4A33 clay band (white 7.28:1)
            "field_border": bgr(0xDA, 0xD3, 0xC6),   # #DAD3C6
        },
    },
}

# Common direction traits (every curated direction shares these). Kept here so
# _resolve_theme stays a thin expander and the per-direction dicts only carry
# what actually varies.
DIRECTION_COMMON: dict = {
    "chrome_off": True, "header_band": True, "flat_buttons": False,
}


# ---------------------------------------------------------------------------
# Grid snapping
# ---------------------------------------------------------------------------

def snap(value, grid: int = GRID) -> int:
    """Round a twip value to the nearest grid dot (default 60). None-safe-ish:
    coerces ints/floats/numeric strings; non-numeric falls back to 0."""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return 0
    if grid <= 0:
        return int(round(v))
    return int(round(v / grid) * grid)


def snap_up(value, grid: int = GRID) -> int:
    """Round a twip value UP to the next grid dot. Use for heights that must not
    clip their content (a label/band that has to fully contain its text)."""
    import math
    try:
        v = float(value)
    except (TypeError, ValueError):
        return 0
    if grid <= 0:
        return int(math.ceil(v))
    return int(math.ceil(v / grid) * grid)


def line_height(font_size_pt: float) -> int:
    """Twips a single line of text needs (1pt = 20 twips, ~1.5x leading), rounded
    up to the grid — the minimum height for a label/band not to clip the text."""
    return snap_up(float(font_size_pt) * 20 * 1.5)
