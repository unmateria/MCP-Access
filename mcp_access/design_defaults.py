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
# Margins & spacing (all multiples of GRID)
# ---------------------------------------------------------------------------
MARGIN_X = 240          # left/right form margin
MARGIN_Y = 240          # top/bottom margin inside a section
GAP_LABEL = 120         # horizontal gap between a label and its field
ROW_GAP = 120           # vertical gap between rows
COL_GAP = 360           # gap between columns in a two-column layout

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
