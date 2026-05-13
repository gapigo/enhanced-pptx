"""EMU ↔ data-coordinate positioner for chart overlays.

Given a chart space, reads the plot-area bounding box in EMU and the
axis scaling, then provides ``data_to_emu(x, y)`` to convert a data
point into slide-level EMU coordinates suitable for positioning an
overlay shape.

Example:
    >>> with open("chart1.xml") as f:
    ...     chart_xml = f.read()
    >>> pos = Positioner.from_chart_xml(chart_xml, slide_w=9144000, slide_h=6858000)
    >>> emu_x, emu_y = pos.data_to_emu("2025-01-01", 100.0)
"""

from __future__ import annotations

import datetime
import math
import re
from typing import Any

import lxml.etree as etree

from kinea_pptx_plus.xml_primitives.nsmap import A_NS, C_NS

# Regex for EMU values in element text
_EMU_RE = re.compile(r"^-?\d+$")


class Positioner:
    """Converts data coordinates to EMU positions for overlay shapes.

    Usage:
        1. Parse chart XML.
        2. Call ``from_chart_xml`` or build manually.
        3. Call ``data_to_emu(x, y)`` for each data point.
    """

    def __init__(
        self,
        plot_x_emu: int,
        plot_y_emu: int,
        plot_cx_emu: int,
        plot_cy_emu: int,
        cat_labels: list[str],
        val_min: float,
        val_max: float,
        categories_are_dates: bool = False,
    ):
        self.plot_x = plot_x_emu
        self.plot_y = plot_y_emu
        self.plot_cx = plot_cx_emu
        self.plot_cy = plot_cy_emu
        self.cat_labels = cat_labels
        self.val_min = val_min
        self.val_max = val_max
        self.categories_are_dates = categories_are_dates
        self._num_cats = max(len(cat_labels), 1)

    # ---- factory ---------------------------------------------------------

    @classmethod
    def from_chart_xml(
        cls,
        chart_xml: str | bytes,
        slide_w: int = 9144000,
        slide_h: int = 6858000,
    ) -> Positioner:
        """Parse chart XML and build a Positioner.

        Args:
            chart_xml: The ``<c:chartSpace>`` XML as string or bytes.
            slide_w:   Slide width in EMU (default 10″).
            slide_h:   Slide height in EMU (default 7.5″).

        Returns:
            A ``Positioner`` ready for ``data_to_emu`` calls.
        """
        root = etree.fromstring(chart_xml) if isinstance(chart_xml, bytes) \
            else etree.fromstring(chart_xml.encode("utf-8"))

        # --- plot area bounding box ---
        plot = root.find(f"{{{C_NS}}}chart/{{{C_NS}}}plotArea")
        if plot is None:
            raise ValueError("No c:plotArea found in chart XML")

        layout = plot.find(f"{{{C_NS}}}layout/{{{C_NS}}}manualLayout")
        margin_w = int(slide_w * 0.10)  # fallback: 10% margins
        margin_h = int(slide_h * 0.12)

        if layout is not None:
            # Try to read manual layout dimensions
            x_emu = _layout_val(layout, "x") or margin_w
            y_emu = _layout_val(layout, "y") or margin_h
            cx_emu = _layout_val(layout, "cx") or (slide_w - 2 * margin_w)
            cy_emu = _layout_val(layout, "cy") or (slide_h - 2 * margin_h)
        else:
            x_emu = margin_w
            y_emu = margin_h
            cx_emu = slide_w - 2 * margin_w
            cy_emu = slide_h - 2 * margin_h

        # --- categories ---
        cat_labels: list[str] = []
        is_date = False
        cat_ax = plot.find(f"{{{C_NS}}}catAx")
        if cat_ax is not None:
            # Get categories from chart data rather than XML (easier)
            # from the catAx scaling info
            ax_scale = cat_ax.find(f"{{{C_NS}}}scaling")
            if ax_scale is not None:
                pass  # orientation, etc.

        # We'll read categories from the first series
        first_ser = plot.find(f".//{{{C_NS}}}ser[1]")
        if first_ser is not None:
            tx_elems = first_ser.findall(f"{{{C_NS}}}cat/{{{C_NS}}}strRef/"
                                         f"{{{C_NS}}}strCache/{{{C_NS}}}ptc")
            # also try numRef for dates
            num_pt = first_ser.findall(f"{{{C_NS}}}cat/{{{C_NS}}}numRef/"
                                       f"{{{C_NS}}}numCache/{{{C_NS}}}ptc")
            if num_pt:
                is_date = True
                cat_labels = [p.get("idx", "") or (p.text or "") for p in num_pt]
            else:
                cat_labels = [p.get("idx", "") or (p.text or "")
                              for p in tx_elems]
            # If no cache found, try the string reference
            if not cat_labels:
                ref_elems = first_ser.findall(f"{{{C_NS}}}cat/{{{C_NS}}}strRef/"
                                              f"{{{C_NS}}}strCache/{{{C_NS}}}pt/"
                                              f"{{{C_NS}}}v")
                cat_labels = [e.text or "" for e in ref_elems]
        if not cat_labels:
            cat_labels = ["Cat1", "Cat2", "Cat3"]  # fallback

        # --- value axis min/max ---
        val_min = 0.0
        val_max = 100.0
        val_ax = plot.find(f"{{{C_NS}}}valAx")
        if val_ax is not None:
            ax_scale = val_ax.find(f"{{{C_NS}}}scaling")
            if ax_scale is not None:
                min_e = ax_scale.find(f"{{{C_NS}}}min")
                max_e = ax_scale.find(f"{{{C_NS}}}max")
                if min_e is not None:
                    try:
                        val_min = float(min_e.text or 0)
                    except (ValueError, TypeError):
                        pass
                if max_e is not None:
                    try:
                        val_max = float(max_e.text or 100)
                    except (ValueError, TypeError):
                        pass
            else:
                # Guess from chart data
                vals = cls._read_val_data(plot)
                if vals:
                    val_min = min(vals)
                    val_max = max(vals)

        return cls(
            plot_x_emu=x_emu,
            plot_y_emu=y_emu,
            plot_cx_emu=cx_emu,
            plot_cy_emu=cy_emu,
            cat_labels=cat_labels,
            val_min=val_min,
            val_max=val_max,
            categories_are_dates=is_date,
        )

    @staticmethod
    def _read_val_data(plot: etree.Element) -> list[float]:
        """Read all value data points from a plot to guess axis range."""
        values: list[float] = []
        for ser in plot.findall(f".//{{{C_NS}}}ser"):
            for val_elem in ser.findall(f"{{{C_NS}}}val/{{{C_NS}}}numRef/"
                                        f"{{{C_NS}}}numCache/{{{C_NS}}}ptc"):
                try:
                    values.append(float(val_elem.text or 0))
                except (ValueError, TypeError):
                    pass
            for v_elem in ser.findall(f"{{{C_NS}}}val/{{{C_NS}}}numRef/"
                                      f"{{{C_NS}}}numCache/{{{C_NS}}}pt/"
                                      f"{{{C_NS}}}v"):
                try:
                    values.append(float(v_elem.text or 0))
                except (ValueError, TypeError):
                    pass
        return values

    # ---- conversion ------------------------------------------------------

    def data_to_emu(
        self,
        x_data: str | int | float,
        y_data: float,
        cat_index: int | None = None,
    ) -> tuple[int, int]:
        """Convert data coordinates to slide-level EMU.

        Args:
            x_data:   Category label (for string cats) or index.
            y_data:   Value axis data point.
            cat_index: If provided, use this index directly (faster).

        Returns:
            ``(emu_x, emu_y)`` relative to the **slide** (not plot area).
        """
        # X position: map category index to plot-area X, then add plot offset
        if cat_index is not None:
            idx = cat_index
        elif isinstance(x_data, (int, float)):
            idx = int(x_data)
        else:
            idx = self._cat_index(str(x_data))

        # Evenly space categories across plot area width
        # First category at left edge, last at right edge
        spacing = self.plot_cx / max(self._num_cats, 1)
        emu_x = self.plot_x + int(spacing * idx) + int(spacing / 2)

        # Y position: linear interpolation within value range
        val_range = self.val_max - self.val_min
        if val_range <= 0:
            val_range = 1.0
        normalized_y = (y_data - self.val_min) / val_range
        # In PowerPoint, Y=0 at plot bottom, Y=max at plot top
        emu_y = self.plot_y + self.plot_cy - int(normalized_y * self.plot_cy)

        return (emu_x, emu_y)

    def _cat_index(self, label: str) -> int:
        try:
            return self.cat_labels.index(label)
        except ValueError:
            return 0

    def last_cat_index(self) -> int:
        """Return the index of the last category (for end-of-series labels)."""
        return max(self._num_cats - 1, 0)

    def series_last_y(self, plot: etree.Element, series_idx: int = 0) -> float:
        """Read the last Y value of a series from chart XML."""
        values = self._read_val_data(plot)
        # Per series — we need to split by series
        sers = plot.findall(f".//{{{C_NS}}}ser")
        if series_idx < len(sers):
            ser = sers[series_idx]
            ser_vals: list[float] = []
            for v_elem in ser.findall(f"{{{C_NS}}}val/{{{C_NS}}}numRef/"
                                      f"{{{C_NS}}}numCache/{{{C_NS}}}pt/"
                                      f"{{{C_NS}}}v"):
                try:
                    ser_vals.append(float(v_elem.text or 0))
                except (ValueError, TypeError):
                    pass
            if ser_vals:
                return ser_vals[-1]
        return 0.0


def _layout_val(layout: etree.Element, name: str) -> int | None:
    """Helper: read a numeric value from manualLayout child element."""
    val_el = layout.find(f"{{{C_NS}}}{name}")
    if val_el is not None:
        try:
            # Values in manualLayout are fraction of parent (chart area) × 1000
            return int(float(val_el.text or 0))
        except (ValueError, TypeError):
            pass
    return None
