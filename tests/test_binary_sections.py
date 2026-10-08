"""
Pure-Python unit tests for putting binary blocks back into a stripped
form/report export (access_set_code after access_get_code).

Covers:
  * the v0.7.63 bug: RecSrcDt / GUID / NameMap were appended after the
    sections block, where LoadFromText only accepts `End`
    ("Expected: 'End'. Found: RecSrcDt") — every round trip failed;
  * strip + inject reproducing the export exactly, the per-section and
    per-control GUIDs included (they used to be dropped, and the last one
    overwrote the form's own GUID);
  * each block anchored after the property that preceded it, including a
    wrapped value and a non-binary `ConditionalFormat = Begin` block;
  * fallbacks: anchor property gone, owner renamed, owner moved deeper,
    no sections block at all.

The sample mirrors the layout of a real Access 2016 / 365 export (measured
in the v0.7.64 verification); names and hex are invented.

No COM / no Access needed. Run with:

    python -m pytest tests/test_binary_sections.py
    python tests/test_binary_sections.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access.helpers import (                                     # noqa: E402
    strip_binary_sections, extract_binary_blocks, inject_binary_blocks,
)


EXPORT = """\
Version =21
VersionRequired =20
Begin Form
    DefaultView =0
    Width =6994
    DatasheetGridlinesColor =15263976
    RecSrcDt = Begin
        0x1111111111111111
    End
    GUID = Begin
        0xf0f0f0f0f0f0f0f0f0f0f0f0f0f0f0f0
    End
    NameMap = Begin
        0x0acc0e5500000000aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa ,
        0x0000
    End
    RecordSource ="tblOrders"
    NoSaveCTIWhenDisabled =1
    Begin
        Begin TextBox
            FontSize =11
        End
        Begin FormHeader
            Height =1134
            Name ="FormHeader"
            GUID = Begin
                0xa1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1
            End
            BackThemeColorIndex =1
        End
        Begin Section
            Height =5952
            Name ="Detail"
            GUID = Begin
                0xd1d1d1d1d1d1d1d1d1d1d1d1d1d1d1d1
            End
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Left =1800
                    Name ="txtTotal"
                    ControlSource ="=FormatPercent(((Nz([a])+Nz([b])-Nz([c]))*1"
                        "00)/Nz([Total])/100)"
                    GUID = Begin
                        0xc1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1
                    End

                    LayoutCachedLeft =1800
                    Begin
                        Begin Label
                            Name ="lblTotal"
                            Caption ="Total"
                            GUID = Begin
                                0xc2c2c2c2c2c2c2c2c2c2c2c2c2c2c2c2
                            End
                        End
                    End
                End
                Begin ComboBox
                    Name ="cboStatus"
                    ConditionalFormat = Begin
                        0x01000000
                    End
                    GUID = Begin
                        0xc3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3
                    End
                End
            End
        End
        Begin FormFooter
            Height =1134
            Name ="FormFooter"
            GUID = Begin
                0xb1b1b1b1b1b1b1b1b1b1b1b1b1b1b1b1
            End
        End
    End
End
"""


def _restore(code: str, original: str = EXPORT) -> str:
    return inject_binary_blocks(code, extract_binary_blocks(original))


def test_round_trip_is_identical():
    assert _restore(strip_binary_sections(EXPORT)) == EXPORT


def test_nothing_but_end_after_the_sections_block():
    # The v0.7.63 bug: blocks landed between the sections block's End and
    # the form's End.  Access accepts nothing there.
    lines = _restore(strip_binary_sections(EXPORT)).splitlines()
    assert lines[-3:] == ["        End", "    End", "End"]


def test_form_guid_is_not_overwritten_by_a_section_guid():
    blocks = extract_binary_blocks(EXPORT)
    assert "0xf0f0" in blocks[""]["GUID"][1]
    assert "0xb1b1" in blocks["FormFooter"]["GUID"][1]
    assert set(blocks) == {"", "FormHeader", "Detail", "FormFooter",
                           "txtTotal", "lblTotal", "cboStatus"}


def test_anchors():
    blocks = extract_binary_blocks(EXPORT)
    assert [a for a, _ in blocks[""].values()] == ["DatasheetGridlinesColor"] * 3
    assert blocks["txtTotal"]["GUID"][0] == "ControlSource"
    assert blocks["cboStatus"]["GUID"][0] == "ConditionalFormat"


def test_block_after_a_wrapped_value_follows_its_last_line():
    out = _restore(strip_binary_sections(EXPORT))
    assert '"00)/Nz([Total])/100)"\n                    GUID = Begin\n' in out


def test_missing_anchor_falls_back_before_the_child_block():
    code = strip_binary_sections(EXPORT).replace("    DatasheetGridlinesColor =15263976\n", "")
    lines = _restore(code).splitlines()
    begin = lines.index("    Begin")
    assert lines[begin - 1] == "    End"                       # NameMap's End
    assert lines.index("    NoSaveCTIWhenDisabled =1") < lines.index("    RecSrcDt = Begin")
    assert lines[-3:] == ["        End", "    End", "End"]


def test_renamed_control_loses_only_its_own_guid():
    code = strip_binary_sections(EXPORT).replace('Name ="cboStatus"', 'Name ="cboState"')
    out = _restore(code)
    assert "0xc3c3" not in out
    assert out.count("GUID = Begin") == EXPORT.count("GUID = Begin") - 1


def test_control_moved_deeper_is_reindented():
    original = (
        "Begin Form\n"
        "    Begin\n"
        "        Begin Section\n"
        '            Name ="Detail"\n'
        "            Begin\n"
        "                Begin TextBox\n"
        '                    Name ="txtA"\n'
        "                    GUID = Begin\n"
        "                        0xaaaa\n"
        "                    End\n"
        "                End\n"
        "            End\n"
        "        End\n"
        "    End\n"
        "End\n"
    )
    moved = (
        "Begin Form\n"
        "    Begin\n"
        "        Begin Section\n"
        '            Name ="Detail"\n'
        "            Begin\n"
        "                Begin Tab\n"
        '                    Name ="tabMain"\n'
        "                    Begin\n"
        "                        Begin Page\n"
        '                            Name ="pagOne"\n'
        "                            Begin\n"
        "                                Begin TextBox\n"
        '                                    Name ="txtA"\n'
        "                                End\n"
        "                            End\n"
        "                        End\n"
        "                    End\n"
        "                End\n"
        "            End\n"
        "        End\n"
        "    End\n"
        "End\n"
    )
    out = _restore(moved, original)
    assert (
        '                                    Name ="txtA"\n'
        "                                    GUID = Begin\n"
        "                                        0xaaaa\n"
        "                                    End\n"
    ) in out


def test_no_sections_block_injects_before_the_outer_end():
    original = (
        "Begin Report\n"
        "    Width =100\n"
        "    PrtMip = Begin\n"
        "        0x6801\n"
        "    End\n"
        "End\n"
    )
    out = _restore("Begin Report\n    Width =100\nEnd\n", original)
    assert out == original


def test_owner_without_name_gets_nothing():
    # The defaults block's prototypes have no Name: nothing to match on.
    out = _restore(strip_binary_sections(EXPORT))
    proto = out.index("        Begin TextBox\n            FontSize =11\n        End\n")
    assert proto > 0


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
