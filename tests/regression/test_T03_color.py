"""Test Color helpers — hex, scheme, tint/shade/alpha, WCAG contrast."""

from __future__ import annotations

import lxml.etree as etree

from kinea_pptx_plus.xml_primitives.color import Color


def _parse_golden(name: str) -> etree.Element:
    return etree.parse(f"tests/fixtures/xml_primitives/{name}").getroot()


def test_from_hex_basic():
    col = Color.from_hex("1F3864")
    assert col.rgb == "1F3864"
    assert col.scheme == ""


def test_from_hex_with_hash():
    col = Color.from_hex("#c0392b")
    assert col.rgb == "C0392B"


def test_from_hex_invalid_raises():
    import pytest
    with pytest.raises(ValueError, match="Invalid hex"):
        Color.from_hex("xyz")


def test_to_srgbClr_matches_golden():
    golden = _parse_golden("golden_srgbClr_basic.xml")
    col = Color.from_hex("1F3864")
    assert etree.tostring(col.to_srgbClr()) == etree.tostring(golden)


def test_to_srgbClr_with_tint():
    golden = _parse_golden("golden_srgbClr_tint.xml")
    col = Color.from_hex("1F3864").with_tint(0.6)
    assert etree.tostring(col.to_srgbClr()) == etree.tostring(golden)


def test_to_srgbClr_with_alpha():
    golden = _parse_golden("golden_srgbClr_alpha.xml")
    col = Color.from_hex("C0392B").with_alpha(50000)
    assert etree.tostring(col.to_srgbClr()) == etree.tostring(golden)


def test_to_schemeClr_matches_golden():
    golden = _parse_golden("golden_schemeClr.xml")
    col = Color.from_scheme("accent1")
    assert etree.tostring(col.to_element()) == etree.tostring(golden)


def test_from_scheme_roundtrip():
    col = Color.from_scheme("dk2")
    elem = col.to_element()
    parsed = Color.from_xml(elem)
    assert parsed.scheme == "dk2"


def test_from_srgbClr_roundtrip():
    col = Color.from_hex("AABBCC").with_tint(0.3).with_alpha(75000)
    elem = col.to_srgbClr()
    parsed = Color.from_xml(elem)
    assert parsed.rgb == "AABBCC"


def test_wcag_contrast_dark_bg():
    """Dark blue → white text."""
    assert Color.wcag_contrast("1F3864") == "FFFFFF"


def test_wcag_contrast_light_bg():
    """Light yellow → black text."""
    assert Color.wcag_contrast("F0E68C") == "000000"


def test_wcag_contrast_red():
    """Kinea red → white."""
    assert Color.wcag_contrast("C0392B") == "FFFFFF"


def test_color_cannot_set_both_rgb_and_scheme():
    import pytest
    with pytest.raises(ValueError, match="Cannot set both"):
        Color(rgb="1F3864", scheme="accent1")
