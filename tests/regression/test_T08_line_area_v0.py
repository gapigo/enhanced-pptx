"""Test line_area_between v0 — native OOXML analysis.

This module documents the limitations of native OOXML for
"area between two lines" and validates that the module exists.
"""

from __future__ import annotations

from kinea_pptx_plus.charts import line_area_between


def test_module_imports():
    """Module should be importable and have the required functions."""
    assert hasattr(line_area_between, "build_line_area_between")
    assert hasattr(line_area_between, "_series_data_pairs")


def test_series_data_pairs_reads_from_fixture():
    """Should read series data from the simple line chart fixture."""
    from kinea_pptx_plus.io import extract_chart_xml
    from kinea_pptx_plus.charts.line_area_between import _series_data_pairs, _extract_plot_elem

    xml = extract_chart_xml("tests/fixtures/simple_line_chart.pptx")
    plot = _extract_plot_elem(xml)

    data = _series_data_pairs(plot, 0, 1)
    assert len(data) >= 4  # Q1-Q4
    # Each row: (index, val_a, val_b)
    assert data[0][1] == 42.0  # Series A first value
    assert data[0][2] == 38.0  # Series B first value


def test_series_data_pairs_5yr():
    """Should read data from the 5-year fixture."""
    from kinea_pptx_plus.io import extract_chart_xml
    from kinea_pptx_plus.charts.line_area_between import _series_data_pairs, _extract_plot_elem

    xml = extract_chart_xml("tests/fixtures/andes_vs_cdi_5yr.pptx")
    plot = _extract_plot_elem(xml)

    data = _series_data_pairs(plot, 0, 1)
    assert len(data) >= 60  # ~70 monthly points
