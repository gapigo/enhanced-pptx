"""Area-between chart — shaded region between two data series.

v0 — Native OOXML attempt (``c:areaChart`` with transparency).
v1 — Polygon overlay path (``a:custGeom`` freeform shapes).

## OOXML Native Analysis

OOXML has no single element to shade "area between two lines".
Approaches considered:

1. **Area chart + line chart in same plotArea** (combo):
   - Create an ``c:areaChart`` with series A, then overlay ``c:lineChart``
     with series A + B. The area fills below series A to the bottom.
   - **Limitation**: fills *below* series A, not *between* A and B.
   - **Verdict**: does not match the spec.

2. **Fill series with transparency stacked**:
   - Area chart with 2 series, the higher one with transparency.
   - **Limitation**: PowerPoint renders area fills as cumulative stacks
     by default, which distorts the visual.
   - **Verdict**: unusable for "delta" shading.

3. **Polygon overlay** (v1 strategy):
   - Draw ``a:custGeom`` freeform shapes with alpha fill covering only
     the space between the two curves.
   - **Verdict**: the correct approach. Implemented in v1.

Example:
    >>> from kinea_pptx_plus.charts.line_area_between import (
    ...     build_line_area_between)
    >>> build_line_area_between("fixture.pptx", output_path="out.pptx")
"""

from __future__ import annotations

import os
import tempfile
from typing import Sequence

import lxml.etree as etree

from pptx import Presentation as PptxPresentation

from kinea_pptx_plus.io import extract_chart_xml
from kinea_pptx_plus.overlays.positioner import Positioner
from kinea_pptx_plus.overlays.polygon_freeform import add_freeform_polygon
from kinea_pptx_plus.xml_primitives.color import Color
from kinea_pptx_plus.xml_primitives.nsmap import A_NS, C_NS


def build_line_area_between(
    pptx_path: str,
    series_a: int = 0,
    series_b: int = 1,
    fill_color: str = "4472C4",
    fill_alpha: int = 40000,
    output_path: str | None = None,
    chart_index: int = 0,
) -> str:
    """Shade the area between two line series using a polygon overlay.

    Args:
        pptx_path:  Source PPTX file path.
        series_a:   Index of the first (lower) series.
        series_b:   Index of the second (upper) series.
        fill_color: Hex colour for the shaded area.
        fill_alpha: Alpha in 1/1000% units (0-100000; default 40000 = 40%).
        output_path: Output PPTX path.  Auto-temp if omitted.
        chart_index: Index of the chart to modify (default 0).

    Returns:
        Path to the output PPTX.
    """
    if output_path is None:
        fd, output_path = tempfile.mkstemp(suffix=".pptx")
        os.close(fd)

    # Step 1: read chart XML for positioner and data
    chart_xml = extract_chart_xml(pptx_path, chart_index)
    pos = Positioner.from_chart_xml(chart_xml)

    # Step 2: read series data from chart XML
    plot = _extract_plot_elem(chart_xml)
    rows = _series_data_pairs(plot, series_a, series_b)

    if len(rows) < 2:
        raise ValueError(f"Need at least 2 data points, got {len(rows)}")

    # Step 3: convert data to EMU vertices
    vertices: list[tuple[int, int]] = []
    for cat_idx, val_a, val_b in rows:
        x_a, y_a = pos.data_to_emu(cat_idx, val_a, cat_index=cat_idx)
        x_b, y_b = pos.data_to_emu(cat_idx, val_b, cat_index=cat_idx)
        vertices.append((x_a, y_a))

    # Reverse direction for the upper series (bottom-right to top-right,
    # then back left along the top)
    for cat_idx, val_a, val_b in reversed(rows):
        x_b, y_b = pos.data_to_emu(cat_idx, val_b, cat_index=cat_idx)
        vertices.append((x_b, y_b))

    # Close the polygon
    vertices.append(vertices[0])

    # Step 4: open PPTX and add polygon overlay
    prs = PptxPresentation(pptx_path)
    slide = prs.slides[0]

    add_freeform_polygon(
        slide,
        vertices=vertices,
        fill_color=fill_color,
        fill_alpha=fill_alpha,
    )

    prs.save(output_path)
    return output_path


def _extract_plot_elem(chart_xml: bytes) -> etree.Element | None:
    root = etree.fromstring(chart_xml)
    plot = root.find(f"{{{C_NS}}}chart/{{{C_NS}}}plotArea")
    if plot is None:
        plot = root.find(f".//{{{C_NS}}}plotArea")
    return plot


def _series_data_pairs(
    plot: etree.Element | None,
    idx_a: int,
    idx_b: int,
) -> list[tuple[int, float, float]]:
    """Read paired data from two series in chart XML."""
    if plot is None:
        raise ValueError("No plotArea found")

    sers = plot.findall(f".//{{{C_NS}}}ser")
    if idx_a >= len(sers) or idx_b >= len(sers):
        raise ValueError(f"Series index out of range: {len(sers)} series")

    def _read_series(ser_idx: int) -> list[float]:
        ser = sers[ser_idx]
        vals: list[float] = []
        # Try numCache/pt/v first
        for pt in ser.findall(f"{{{C_NS}}}val/{{{C_NS}}}numRef/"
                              f"{{{C_NS}}}numCache/{{{C_NS}}}pt"):
            v = pt.find(f"{{{C_NS}}}v")
            if v is not None and v.text:
                try:
                    vals.append(float(v.text))
                except ValueError:
                    vals.append(0.0)
        if not vals:
            # Fallback: ptc elements
            for ptc in ser.findall(f"{{{C_NS}}}val/{{{C_NS}}}numCache/{{{C_NS}}}ptc"):
                v = ptc.find(f"{{{C_NS}}}v")
                if v is not None and v.text:
                    try:
                        vals.append(float(v.text))
                    except ValueError:
                        vals.append(0.0)
        return vals

    va = _read_series(idx_a)
    vb = _read_series(idx_b)

    n = min(len(va), len(vb))
    return [(i, va[i], vb[i]) for i in range(n)]
