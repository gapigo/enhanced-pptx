"""Line chart with end-of-series coloured labels — XML path.

v0 — paints each series with a specified colour via
``c:ser/c:spPr/a:ln/a:solidFill/a:srgbClr``.

Example:
    >>> from kinea_pptx_plus.charts.line_with_end_label import paint_series_colors
    >>> from kinea_pptx_plus.io import extract_chart_xml
    >>> paint_series_colors("tests/fixtures/simple_line_chart.pptx",
    ...                     ["1F3864", "C0392B"])
"""

from __future__ import annotations

from typing import Sequence

import lxml.etree as etree

from pptx import Presentation as PptxPresentation
from pptx.chart.chart import Chart
from pptx.shapes.graphfrm import GraphicFrame

from kinea_pptx_plus.xml_primitives.color import Color
from kinea_pptx_plus.xml_primitives.nsmap import A_NS, C_NS


def paint_series_colors(
    pptx_path: str,
    colors: Sequence[str],
    chart_index: int = 0,
    output_path: str | None = None,
) -> bytes:
    """Paint each series in the chart with the corresponding colour.

    Modifies ``c:ser/c:spPr/a:ln/a:solidFill/a:srgbClr`` for each series.
    If a series already has a ``c:spPr``, its line fill is updated;
    otherwise a new ``c:spPr`` structure is inserted.

    Args:
        pptx_path:  Path to the source PPTX.
        colors:     One hex colour string per series (``"RRGGBB"`` format).
        chart_index: Index of the chart to modify (default 0).
        output_path: If set, save the modified PPTX to this path.

    Returns:
        The modified ``chart1.xml`` as bytes (for in-memory inspection).

    Raises:
        IndexError: If ``chart_index`` is out of range.
        ValueError: If the number of colours doesn't match series count.
    """
    prs = PptxPresentation(pptx_path)

    charts: list[Chart] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if isinstance(shape, GraphicFrame) and shape.has_chart:
                charts.append(shape.chart)

    if chart_index >= len(charts):
        raise IndexError(
            f"chart_index={chart_index} out of range; "
            f"{len(charts)} chart(s) found"
        )

    chart: Chart = charts[chart_index]
    plot = chart.plots[0]
    series_count = len(plot.series)

    if len(colors) < series_count:
        raise ValueError(
            f"Expected {series_count} colours, got {len(colors)}"
        )

    # Access chart XML directly via lxml
    chart_space = chart._chartSpace  # type: ignore[attr-defined]
    nsmap_chart = {"c": C_NS, "a": A_NS}
    series_elems = chart_space.findall(
        f".//{{{C_NS}}}plotArea/{{{C_NS}}}lineChart/{{{C_NS}}}ser"
    )

    # Fallback: try to find any series in plotArea
    if not series_elems:
        series_elems = chart_space.findall(
            f".//{{{C_NS}}}ser"
        )

    for idx, ser in enumerate(series_elems):
        if idx >= len(colors):
            break
        color = colors[idx]
        _set_series_line_color(ser, color)

    if output_path is not None:
        prs.save(output_path)

    return etree.tostring(chart_space, xml_declaration=True, encoding="UTF-8")


def _set_series_line_color(ser: etree.Element, rgb: str) -> None:
    """Set the line colour for a single ``c:ser`` element.

    Either updates an existing ``c:spPr/a:ln/a:solidFill/a:srgbClr``
    or creates the full structure.
    """
    sp_pr = ser.find(f"{{{C_NS}}}spPr")

    if sp_pr is None:
        sp_pr = etree.SubElement(ser, f"{{{C_NS}}}spPr")

    a_ln = sp_pr.find(f"{{{A_NS}}}ln")

    if a_ln is None:
        a_ln = etree.SubElement(sp_pr, f"{{{A_NS}}}ln")

    solid = a_ln.find(f"{{{A_NS}}}solidFill")
    if solid is None:
        solid = etree.SubElement(a_ln, f"{{{A_NS}}}solidFill")

    clr = solid.find(f"{{{A_NS}}}srgbClr")
    if clr is None:
        clr = etree.SubElement(solid, f"{{{A_NS}}}srgbClr")

    clr.set("val", rgb.upper())
