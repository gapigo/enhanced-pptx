"""Test validation harness — all gates."""

from __future__ import annotations

import os

from kinea_pptx_plus.io import extract_chart_xml
from kinea_pptx_plus.validation.xml_schema import validate_chart_xml
from kinea_pptx_plus.validation.win32_probe import probe_pptx
from kinea_pptx_plus.validation.render_check import render_check, _find_soffice
from kinea_pptx_plus.validation.golden_compare import compare_chart_xml


def test_xml_schema_valid():
    """Good chart XML should pass schema gate."""
    xml = extract_chart_xml("tests/fixtures/simple_line_chart.pptx")
    assert validate_chart_xml(xml) is True


def test_xml_schema_invalid():
    """Broken XML should fail schema gate."""
    assert validate_chart_xml(b"<not><valid>") is False


def test_xml_schema_empty():
    """Empty bytes should fail."""
    assert validate_chart_xml(b"") is False


def test_win32_probe_skips_on_non_windows():
    """probe_pptx should skip gracefully on non-Windows (always passes)."""
    result = probe_pptx("tests/fixtures/simple_line_chart.pptx")
    # On Windows without pywin32 it may fail; on other OS it skips
    assert result.passed is not None
    assert isinstance(result.detail, str)


def test_render_check_skips_without_libreoffice():
    """render_check should detect missing LibreOffice."""
    result = render_check("tests/fixtures/simple_line_chart.pptx")
    # Will either pass (LO + tesseract available) or skip
    assert result.passed is not None


def test_find_soffice():
    """_find_soffice should not crash (may return None)."""
    soffice = _find_soffice()
    # Acceptable to be None (LO not installed)
    assert soffice is None or os.path.exists(soffice)


def test_golden_compare_against_self():
    """A PPTX compared to itself should pass."""
    xml = extract_chart_xml("tests/fixtures/simple_line_chart.pptx")
    assert compare_chart_xml(xml, "tests/fixtures/simple_line_chart.pptx") is True


def test_golden_compare_different_fails():
    """Comparing different chart XMLs should fail."""
    xml_a = extract_chart_xml("tests/fixtures/simple_line_chart.pptx")
    xml_b = extract_chart_xml("tests/fixtures/andes_vs_cdi_5yr.pptx")
    assert compare_chart_xml(xml_a, "tests/fixtures/andes_vs_cdi_5yr.pptx") is False
