"""Test line_with_end_label v0 — XML colour painting."""

from __future__ import annotations

import lxml.etree as etree

from kinea_pptx_plus.charts.line_with_end_label import paint_series_colors
from kinea_pptx_plus.io import extract_chart_xml


def test_paint_series_colors_simple_fixture():
    """Paint 2 series on the simple line chart."""
    xml = paint_series_colors(
        "tests/fixtures/simple_line_chart.pptx",
        ["1F3864", "C0392B"],
    )
    root = etree.fromstring(xml)
    c_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    colors_found = [
        clr.get("val")
        for spPr in root.iter(f"{{{c_ns}}}spPr")
        for ln in spPr.iter(f"{{{a_ns}}}ln")
        for sf in ln.iter(f"{{{a_ns}}}solidFill")
        for clr in sf.iter(f"{{{a_ns}}}srgbClr")
    ]
    assert len(colors_found) >= 2, f"Found {len(colors_found)} colours"
    assert "1F3864" in colors_found
    assert "C0392B" in colors_found


def test_paint_series_colors_5yr_fixture():
    """Paint 2 series on the 5-year Andes/CDI fixture."""
    xml = paint_series_colors(
        "tests/fixtures/andes_vs_cdi_5yr.pptx",
        ["1F3864", "C0392B"],
    )
    root = etree.fromstring(xml)
    c_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    colors_found = [
        clr.get("val")
        for spPr in root.iter(f"{{{c_ns}}}spPr")
        for ln in spPr.iter(f"{{{a_ns}}}ln")
        for sf in ln.iter(f"{{{a_ns}}}solidFill")
        for clr in sf.iter(f"{{{a_ns}}}srgbClr")
    ]
    assert len(colors_found) >= 2, f"Found {len(colors_found)} colours in 5yr fixture"


def test_output_preserves_chart_structure():
    """The chart should still contain lineChart after painting."""
    xml = paint_series_colors(
        "tests/fixtures/simple_line_chart.pptx",
        ["1F3864", "C0392B"],
    )
    assert b"lineChart" in xml


def test_invalid_color_count_raises():
    """Calling with fewer colours than series should error."""
    import pytest
    with pytest.raises(ValueError, match="Expected"):
        paint_series_colors(
            "tests/fixtures/simple_line_chart.pptx",
            ["1F3864"],  # only 1 colour for 2 series
        )


def test_can_save_output_pptx():
    """Saving to a temp path should produce a valid PPTX."""
    import tempfile
    import os
    from pptx import Presentation

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        out_path = tmp.name
    try:
        paint_series_colors(
            "tests/fixtures/simple_line_chart.pptx",
            ["1F3864", "C0392B"],
            output_path=out_path,
        )
        prs = Presentation(out_path)
        assert len(prs.slides) == 1
    finally:
        if os.path.exists(out_path):
            os.unlink(out_path)


def test_roundtrip_preserves_colors():
    """Save, reload, and verify colours persist."""
    import tempfile
    import os

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        out_path = tmp.name
    try:
        paint_series_colors(
            "tests/fixtures/simple_line_chart.pptx",
            ["1F3864", "C0392B"],
            output_path=out_path,
        )
        xml2 = extract_chart_xml(out_path)
        assert b"1F3864" in xml2
        assert b"C0392B" in xml2
    finally:
        if os.path.exists(out_path):
            os.unlink(out_path)
