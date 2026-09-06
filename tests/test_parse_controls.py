"""
Pure-Python unit tests for the SaveAsText control parser (_parse_controls).

No COM / no Access needed. The fixture below is a trimmed copy of a real
Access 2016 export of a form holding a tab control with two pages, so the
nesting, the anonymous wrapper blocks and the omitted `ControlType =` are
exactly what Access writes. Run with:

    python -m pytest tests/test_parse_controls.py
    python tests/test_parse_controls.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access.controls import _parse_controls  # noqa: E402


# A tab control exports as `Begin Tab`, its pages live inside an anonymous
# `Begin` wrapper, and each page wraps its own children the same way. Note the
# leading defaults block: same tokens, no `Name =`. Modern exports omit
# `ControlType =` entirely — the type has to come from the Begin token.
FORM_TEXT = "\n".join([
    "Version =21",
    "VersionRequired =20",
    "Begin Form",
    "    Width =9200",
    "    Begin",                       # defaults block — NOT controls
    "        Begin CommandButton",
    "            Width =1701",
    "        End",
    "        Begin Tab",
    "            Width =5103",
    "        End",
    "        Begin Page",
    "            Width =1701",
    "        End",
    "    End",
    "    Begin Section",
    "        Height =5952",
    '        Name ="Detail"',
    "        Begin",
    "            Begin Tab",
    "                Left =195",
    "                Top =600",
    "                Width =9005",
    "                Height =5000",
    '                Name ="tabDetails"',
    "                Begin",
    "                    Begin Page",
    "                        Left =270",
    "                        Top =1065",
    "                        Width =8850",
    "                        Height =4455",
    '                        Name ="pagGeneral"',
    "                        Begin",
    "                            Begin CommandButton",
    "                                Left =1200",
    "                                Top =1500",
    "                                Width =2000",
    "                                Height =400",
    '                                Name ="btInside"',
    "                                GUID = Begin",
    "                                    0x0102",
    "                                End",
    "                            End",
    "                        End",
    "                    End",
    "                    Begin Page",
    "                        Left =270",
    "                        Top =1065",
    "                        Width =8850",
    "                        Height =4455",
    '                        Name ="pagLines"',
    "                    End",
    "                End",
    "            End",
    "            Begin CommandButton",
    "                Left =1200",
    "                Top =2500",
    "                Width =2000",
    "                Height =400",
    '                Name ="btOrphan"',
    "            End",
    "            Begin Label",
    "                Left =200",
    "                Top =6000",
    "                Width =1000",
    "                Height =300",
    '                Name ="lblOutside"',
    "            End",
    "        End",
    "    End",
    "End",
])


def _named(text=FORM_TEXT):
    """Controls that carry a Name, keyed by name — what the tools expose."""
    return {c["name"]: c for c in _parse_controls(text)["controls"] if c["name"]}


def test_tab_control_is_enumerated():
    """It used to be invisible: 123 was missing from CTRL_TYPE, so a form's tab
    control had no geometry anywhere in the package."""
    ctrls = _named()
    assert "tabDetails" in ctrls
    tab = ctrls["tabDetails"]
    assert tab["type_name"] == "Tab"
    assert (tab["left"], tab["top"]) == ("195", "600")
    assert (tab["width"], tab["height"]) == ("9005", "5000")


def test_pages_declare_their_tab_as_parent():
    ctrls = _named()
    assert ctrls["pagGeneral"]["parent"] == "tabDetails"
    assert ctrls["pagLines"]["parent"] == "tabDetails"


def test_children_of_a_page_still_enumerate():
    """The regression guard for making Tab a container: a recognised
    non-container would skip to the end of its block and swallow every page
    and everything on them."""
    ctrls = _named()
    assert "btInside" in ctrls, "page children were swallowed — is Tab still in CONTAINER_TYPES?"
    assert ctrls["btInside"]["parent"] == "pagGeneral"
    assert ctrls["btInside"]["left"] == "1200"


def test_nested_property_block_does_not_close_the_control_early():
    """`GUID = Begin` opens a block closed by its own End. Counting only plain
    `Begin <Type>` used to close btInside there and lose everything after it."""
    ctrls = _named()
    assert "pagLines" in ctrls          # comes after the GUID block
    assert "btOrphan" in ctrls


def test_controls_outside_the_tab_have_no_parent():
    ctrls = _named()
    assert "parent" not in ctrls["btOrphan"]
    assert "parent" not in ctrls["lblOutside"]
    assert "parent" not in ctrls["tabDetails"]


def test_control_type_is_resolved_without_an_explicit_ControlType():
    """Access omits `ControlType =` when it equals the default, which left
    control_type at -1 while type_name was right — a number that could not be
    fed back to ac_create_control."""
    ctrls = _named()
    assert ctrls["tabDetails"]["control_type"] == 123
    assert ctrls["pagGeneral"]["control_type"] == 118
    assert ctrls["btOrphan"]["control_type"] == 104
    assert ctrls["lblOutside"]["control_type"] == 100
    assert all(c["control_type"] != -1 for c in ctrls.values())


def test_explicit_ControlType_still_wins():
    text = FORM_TEXT.replace('                Name ="btOrphan"',
                             '                Name ="btOrphan"\n'
                             "                ControlType =104")
    assert _named(text)["btOrphan"]["control_type"] == 104


def test_defaults_block_contributes_no_named_controls():
    """The block before the first section holds unnamed prototypes of the very
    same types — they must never reach the caller."""
    assert set(_named()) == {"tabDetails", "pagGeneral", "pagLines",
                             "btInside", "btOrphan", "lblOutside"}


if __name__ == "__main__":
    for _name, _fn in sorted(globals().items()):
        if _name.startswith("test_") and callable(_fn):
            _fn()
            print("ok", _name)
    print("all passed")
