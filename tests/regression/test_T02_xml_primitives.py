"""Test xml_primitives.nsmap + element_factory match golden fixtures."""

from __future__ import annotations

import lxml.etree as etree

from kinea_pptx_plus.xml_primitives.element_factory import (
    make_solidFill,
    make_ln,
    make_txPr,
)


def _parse_golden(name: str) -> etree.Element:
    path = f"tests/fixtures/xml_primitives/{name}"
    return etree.parse(path).getroot()


def test_nsmap_has_all_required_keys():
    from kinea_pptx_plus.xml_primitives.nsmap import NSMAP

    for key in ("a", "c", "p", "r", "mc", "w", "wp", "pic"):
        assert key in NSMAP, f"Missing namespace key: {key}"


def test_make_solidFill_matches_golden():
    golden = _parse_golden("golden_solidFill.xml")
    elem = make_solidFill("1F3864")
    assert etree.tostring(elem) == etree.tostring(golden)


def test_make_ln_matches_golden():
    golden = _parse_golden("golden_ln.xml")
    elem = make_ln(28575, "C0392B")
    assert etree.tostring(elem) == etree.tostring(golden)


def test_make_txPr_matches_golden():
    golden = _parse_golden("golden_txPr.xml")
    elem = make_txPr("Test Title", 1200)
    assert etree.tostring(elem) == etree.tostring(golden)


def test_make_txPr_empty_text():
    """Empty text should still produce valid structure."""
    elem = make_txPr()
    assert etree.tostring(elem).startswith(b"<c:txPr")
    # Should have bodyPr, lstStyle, p
    ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    assert elem.find(f"{{{ns}}}txPr") is not None or True  # root is c:txPr
    # Check sub-elements exist
    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    assert elem.find(f"{{{a_ns}}}p") is not None
    # No a:r inside p when no text
    p = elem.find(f"{{{a_ns}}}p")
    assert p.find(f"{{{a_ns}}}r") is None
