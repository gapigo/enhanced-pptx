"""SQL → Chart mapper — executes SQL and dispatches to chart functions.

In synthetic mode (no database available), uses csv or in-memory data.

Example:
    >>> from kinea_pptx_plus.registry.sql_to_chart import chart_from_data
    >>> data = [("2024-Q1", 42.0, 38.0), ("2024-Q2", 55.0, 50.0)]
    >>> chart_from_data(data, chart_type="line_with_end_label",
    ...                 colors=["1F3864", "C0392B"],
    ...                 output_path="output.pptx")
    'output.pptx'
"""

from __future__ import annotations

import os
import tempfile
from typing import Any, Sequence

import lxml.etree as etree

from pptx import Presentation as PptxPresentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches


def chart_from_data(
    data: Sequence[tuple[str, float, ...]],
    chart_type: str = "line_with_end_label",
    series_names: list[str] | None = None,
    colors: list[str] | None = None,
    output_path: str | None = None,
    title: str = "Chart",
) -> str:
    """Create a chart from tabular data.

    Args:
        data:         List of ``(category, series0_value, series1_value, …)``.
        chart_type:   One of ``"line_with_end_label"``, ``"line_area_between"``,
                      ``"line_stacked_combo"``.
        series_names: Names for each data series.
        colors:       Hex colours for each series.
        output_path:  Output PPTX path (auto-temp if None).
        title:        Chart title.

    Returns:
        Path to the generated PPTX.
    """
    if not data:
        raise ValueError("No data provided")

    if output_path is None:
        fd, output_path = tempfile.mkstemp(suffix=".pptx")
        os.close(fd)

    num_series = len(data[0]) - 1
    series_names = series_names or [f"Series {i+1}" for i in range(num_series)]

    # Build chart data for python-pptx-ng
    chart_data = CategoryChartData()
    chart_data.categories = [row[0] for row in data]
    for i in range(num_series):
        chart_data.add_series(series_names[i], [row[i + 1] for row in data])

    # Create PPTX with chart
    prs = PptxPresentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Map chart type
    chart_type_map = {
        "line_with_end_label": XL_CHART_TYPE.LINE_MARKERS,
        "line_area_between": XL_CHART_TYPE.LINE_MARKERS,
        "line_stacked_combo": XL_CHART_TYPE.LINE_MARKERS,
    }
    xl_type = chart_type_map.get(chart_type, XL_CHART_TYPE.LINE_MARKERS)

    chart_frame = slide.shapes.add_chart(
        xl_type,
        Inches(0.8), Inches(0.5),
        Inches(11.5), Inches(6.5),
        chart_data,
    )

    chart = chart_frame.chart
    chart.has_legend = True
    chart.chart_title.has_text_frame = True
    chart.chart_title.text_frame.text = title

    prs.save(output_path)

    # Apply colours if specified
    if colors:
        from kinea_pptx_plus.charts.line_with_end_label import paint_series_colors
        paint_series_colors(output_path, colors, output_path=output_path)

    # Build overlays for line_with_end_label
    if chart_type == "line_with_end_label":
        from kinea_pptx_plus.charts.line_with_end_label import build_line_with_end_label
        build_line_with_end_label(
            output_path,
            colors=colors,
            output_path=output_path,
        )
    elif chart_type == "line_area_between":
        from kinea_pptx_plus.charts.line_area_between import build_line_area_between
        build_line_area_between(
            output_path,
            fill_color=(colors[0] if colors else "4472C4"),
            output_path=output_path,
        )

    return output_path


def csv_to_chart(
    csv_path: str,
    chart_type: str = "line_with_end_label",
    colors: list[str] | None = None,
    output_path: str | None = None,
) -> str:
    """Read a CSV and create a chart.

    First column: categories.  Remaining columns: data series.

    Args:
        csv_path:    Path to CSV file.
        chart_type:  Chart type string.
        colors:      Hex colours.
        output_path: Output PPTX path.

    Returns:
        Generated PPTX path.
    """
    import csv

    with open(csv_path, newline="") as f:
        reader = csv.reader(f)
        rows = list(reader)

    if not rows:
        raise ValueError("Empty CSV")

    # First row headers as series names
    series_names = rows[0][1:] if len(rows) > 0 else []
    data = [(r[0], *(float(v) for v in r[1:])) for r in rows[1:] if len(r) > 1]

    return chart_from_data(
        data,
        chart_type=chart_type,
        series_names=series_names,
        colors=colors,
        output_path=output_path,
    )
