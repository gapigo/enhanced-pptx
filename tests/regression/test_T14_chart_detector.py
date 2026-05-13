"""Test chart_detector — OCR heuristics."""

from __future__ import annotations

from kinea_pptx_plus.registry.chart_detector import (
    detect_chart_type,
    detect_from_ocr_text,
    CHART_TYPES,
)


def test_detect_line_with_end_label():
    """2 series, no shades, no bars → line_with_end_label."""
    result = detect_chart_type(num_series=2, num_annotations=1)
    assert result.chart_type == "line_with_end_label"
    assert result.confidence >= 0.5


def test_detect_line_area_between():
    """2 series, has_area_shade → line_area_between."""
    result = detect_chart_type(num_series=2, has_area_shade=True)
    assert result.chart_type == "line_area_between"
    assert result.confidence >= 0.7


def test_detect_line_stacked_combo():
    """has_bar_groups → line_stacked_combo."""
    result = detect_chart_type(num_series=3, has_bar_groups=True)
    assert result.chart_type == "line_stacked_combo"


def test_detect_from_ocr_keyword():
    """OCR with 'stacked' keyword."""
    result = detect_from_ocr_text("Stacked Bar Chart with Line Overlay")
    assert result.chart_type == "line_stacked_combo"


def test_detect_all_types_covered():
    """All CHART_TYPES should be detectable."""
    assert "line_with_end_label" in CHART_TYPES
    assert "line_area_between" in CHART_TYPES
    assert "line_stacked_combo" in CHART_TYPES


def test_unknown_when_no_heuristics():
    """No signals → unknown."""
    result = detect_chart_type(num_series=5)
    assert result.chart_type == "unknown"
