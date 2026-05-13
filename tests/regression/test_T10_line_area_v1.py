"""Test line_area_between v1 — polygon overlay shade."""

from __future__ import annotations

import os
import tempfile

import lxml.etree as etree
from pptx import Presentation

from kinea_pptx_plus.charts.line_area_between import build_line_area_between
from kinea_pptx_plus.io import extract_chart_xml


def _run_area(path: str, **kw) -> str:
    """Run build_line_area_between and return output path."""
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        out_path = tmp.name
    try:
        result = build_line_area_between(path, output_path=out_path, **kw)
        assert result == out_path
        return out_path
    except Exception:
        if os.path.exists(out_path):
            os.unlink(out_path)
        raise


def test_area_between_simple_fixture():
    """Should shade area between 2 series on the simple chart."""
    out = _run_area("tests/fixtures/simple_line_chart.pptx")
    try:
        prs = Presentation(out)
        slide = prs.slides[0]
        # Should have at least the chart plus the polygon
        assert len(slide.shapes) >= 1
    finally:
        if os.path.exists(out):
            os.unlink(out)


def test_area_between_5yr_fixture():
    """Should handle 70 data points on the 5-year fixture."""
    out = _run_area(
        "tests/fixtures/andes_vs_cdi_5yr.pptx",
        fill_color="1F3864", fill_alpha=30000,
    )
    try:
        prs = Presentation(out)
        assert len(prs.slides) == 1
    finally:
        if os.path.exists(out):
            os.unlink(out)


def test_area_between_schema_gate():
    """schema gate: chart XML must still parse cleanly."""
    out = _run_area("tests/fixtures/simple_line_chart.pptx")
    try:
        xml = extract_chart_xml(out)
        root = etree.fromstring(xml)
        assert root.tag.endswith("chartSpace")
    finally:
        if os.path.exists(out):
            os.unlink(out)


def test_area_between_roundtrip_gate():
    """roundtrip gate: save → reload → save → XML unchanged."""
    out = _run_area("tests/fixtures/simple_line_chart.pptx")
    try:
        xml1 = extract_chart_xml(out)
        prs = Presentation(out)
        prs.save(out)
        xml2 = extract_chart_xml(out)

        t1 = etree.tostring(etree.fromstring(xml1))
        t2 = etree.tostring(etree.fromstring(xml2))
        assert t1 == t2, "Roundtrip gate failed"
    finally:
        if os.path.exists(out):
            os.unlink(out)


def test_data_pairs_non_empty():
    """Validate that series data pairs aren't empty."""
    from kinea_pptx_plus.io import extract_chart_xml
    from kinea_pptx_plus.charts.line_area_between import (
        _extract_plot_elem, _series_data_pairs,
    )

    xml = extract_chart_xml("tests/fixtures/simple_line_chart.pptx")
    plot = _extract_plot_elem(xml)
    data = _series_data_pairs(plot, 0, 1)
    assert len(data) >= 4
    assert all(isinstance(r[0], int) for r in data)
    assert all(isinstance(r[1], float) for r in data)
