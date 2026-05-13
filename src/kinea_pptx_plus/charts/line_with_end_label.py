"""Line chart with end-of-series coloured labels.

v0 — paints series colours via XML (``c:ser/c:spPr/a:ln/a:solidFill``).
v1 — adds coloured label-box overlays positioned at the last data point.

Example:
    >>> from kinea_pptx_plus.charts.line_with_end_label import (
    ...     paint_series_colors, build_line_with_end_label)
    >>> paint_series_colors("fixture.pptx", ["1F3864", "C0392B"])
    >>> build_line_with_end_label("fixture.pptx", output_path="out.pptx")
"""

from __future__ import annotations

import os
import tempfile
from typing import Sequence

import lxml.etree as etree

from pptx import Presentation as PptxPresentation
from pptx.chart.chart import Chart
from pptx.shapes.graphfrm import GraphicFrame

from kinea_pptx_plus.io import extract_chart_xml
from kinea_pptx_plus.overlays.label_box import add_label_box
from kinea_pptx_plus.overlays.positioner import Positioner
from kinea_pptx_plus.xml_primitives.color import Color
from kinea_pptx_plus.xml_primitives.nsmap import A_NS, C_NS

KINEA_COLORS = [
    "1F3864", "C0392B", "2E86AB", "A23B72", "F18F01", "3B8C5F",
]


# ═══════════════════════════════════════════════════════════════════
# v0 — XML colour path
# ═══════════════════════════════════════════════════════════════════

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

    chart_space = chart._chartSpace  # type: ignore[attr-defined]
    series_elems = chart_space.findall(
        f".//{{{C_NS}}}ser"
    )

    for idx, ser in enumerate(series_elems):
        if idx >= len(colors):
            break
        _set_series_line_color(ser, colors[idx])

    if output_path is not None:
        prs.save(output_path)

    return etree.tostring(chart_space, xml_declaration=True, encoding="UTF-8")


def _set_series_line_color(ser: etree.Element, rgb: str) -> None:
    """Set the line colour for a single ``c:ser`` element."""
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


# ═══════════════════════════════════════════════════════════════════
# v1 — overlay path (colours + label boxes)
# ═══════════════════════════════════════════════════════════════════

def build_line_with_end_label(
    pptx_path: str,
    colors: Sequence[str] | None = None,
    output_path: str | None = None,
    chart_index: int = 0,
    label_font_size_pt: float = 9.0,
    label_padding_emu: int = 80000,
    label_cx_emu: int = 1100000,
    label_cy_emu: int = 330000,
) -> str:
    """Create a line chart with end-of-series coloured label boxes.

    Combines ``paint_series_colors`` (XML path) with ``add_label_box``
    (overlay path): each series is coloured via XML and a label box
    is placed at the last data point of each series.

    Args:
        pptx_path:          Source PPTX file path.
        colors:             One hex colour per series.  Defaults to Kinea palette.
        output_path:        Output PPTX path.  If omitted, uses a temp file.
        chart_index:        Index of the chart to modify (default 0).
        label_font_size_pt: Font size for label boxes.
        label_padding_emu:  Extra gap between last data point and label.
        label_cx_emu:       Width of each label box.
        label_cy_emu:       Height of each label box.

    Returns:
        Path to the output PPTX.
    """
    if output_path is None:
        fd, output_path = tempfile.mkstemp(suffix=".pptx")
        os.close(fd)

    # Step 1: paint series colours via XML
    paint_series_colors(
        pptx_path,
        list(colors or KINEA_COLORS),
        chart_index=chart_index,
        output_path=output_path,
    )

    # Step 2: read chart XML for positioner
    chart_xml = extract_chart_xml(output_path, chart_index)
    pos = Positioner.from_chart_xml(chart_xml)

    # Step 3: open output PPTX and add label boxes
    prs = PptxPresentation(output_path)
    slide = prs.slides[0]

    actual_colors = list(colors or KINEA_COLORS)
    last_idx = pos.last_cat_index()

    for series_idx in range(len(pos.cat_labels)):
        act_color = (
            actual_colors[series_idx]
            if series_idx < len(actual_colors)
            else "1F3864"
        )

        # Get last Y value for this series
        last_y = pos.series_last_y(
            _extract_plot_elem(chart_xml), series_idx
        )

        # Convert last data point to EMU
        emu_x, emu_y = pos.data_to_emu(last_idx, last_y)
        emu_x += label_padding_emu
        emu_y -= int(label_cy_emu / 2)

        text = f"{last_y:.1f}%"

        add_label_box(
            slide,
            x_emu=emu_x,
            y_emu=emu_y,
            cx_emu=label_cx_emu,
            cy_emu=label_cy_emu,
            fill_hex=act_color,
            text=text,
            font_size_pt=label_font_size_pt,
        )

    prs.save(output_path)
    return output_path


def _extract_plot_elem(chart_xml: bytes) -> etree.Element | None:
    """Extract the ``c:plotArea`` element from chart XML."""
    root = etree.fromstring(chart_xml)
    plot = root.find(f"{{{C_NS}}}chart/{{{C_NS}}}plotArea")
    if plot is None:
        plot = root.find(f".//{{{C_NS}}}plotArea")
    return plot
