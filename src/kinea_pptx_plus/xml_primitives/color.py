"""Color helpers for OOXML chart colour manipulation.

Provides ``Color`` with factory methods for hex, scheme, tint/shade/alpha,
and conversion to ``a:srgbClr`` / ``a:schemeClr`` XML elements.

Example:
    >>> from lxml import etree
    >>> col = Color.from_hex("1F3864")
    >>> elem = col.to_srgbClr()
    >>> etree.tostring(elem)
    b'<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="1F3864"/>'
"""

from __future__ import annotations

import re
from typing import Literal

import lxml.etree as etree

from kinea_pptx_plus.xml_primitives.nsmap import A_NS

_HEX_RE = re.compile(r"^#?([0-9a-fA-F]{6})$")

SchemeColor = Literal[
    "bg1", "tx1", "bg2", "tx2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "dk1", "dk2", "lt1", "lt2",
    "hlink", "folHlink",
]


class Color:
    """An OOXML colour value — either an explicit RGB or a theme-scheme reference.

    Two mutually exclusive modes:
    * **Hex mode** — ``rgb`` is a 6-char uppercase hex string.
    * **Scheme mode** — ``scheme`` is a theme colour name.

    Both modes support optional tint / shade / alpha transforms.
    """

    def __init__(
        self,
        rgb: str = "",
        scheme: str = "",
        tint: float | None = None,
        shade: float | None = None,
        alpha: int | None = None,
    ):
        if rgb and scheme:
            raise ValueError("Cannot set both rgb and scheme on one Color")
        self.rgb = rgb
        self.scheme = scheme
        self.tint = tint
        self.shade = shade
        self.alpha = alpha

    # ---- factories -------------------------------------------------------

    @classmethod
    def from_hex(cls, hex_str: str) -> Color:
        """Parse a 6-digit hex colour string (with or without ``#``)."""
        match = _HEX_RE.match(hex_str.strip())
        if not match:
            raise ValueError(f"Invalid hex colour: {hex_str!r}")
        return cls(rgb=match.group(1).upper())

    @classmethod
    def from_scheme(cls, scheme: SchemeColor) -> Color:
        """Create a theme-scheme colour reference."""
        return cls(scheme=scheme)

    @classmethod
    def from_xml(cls, elem: etree.Element) -> Color:
        """Parse ``Color`` from an ``a:srgbClr`` or ``a:schemeClr`` element."""
        tag = elem.tag
        val = elem.get("val", "")
        if tag.endswith("srgbClr"):
            return cls.from_hex(val)
        if tag.endswith("schemeClr"):
            return cls(scheme=val)
        raise ValueError(f"Unknown colour element: {tag}")

    # ---- serialisation ---------------------------------------------------

    def _q(self, tag: str) -> str:
        return f"{{{A_NS}}}{tag}"

    def to_srgbClr(self) -> etree.Element:
        """Build ``<a:srgbClr val="RRGGBB"/>``."""
        elem = etree.Element(self._q("srgbClr"), nsmap={"a": A_NS})
        elem.set("val", self.rgb)
        self._apply_transforms(elem)
        return elem

    def to_schemeClr(self) -> etree.Element:
        """Build ``<a:schemeClr val="…"/>``."""
        elem = etree.Element(self._q("schemeClr"), nsmap={"a": A_NS})
        elem.set("val", self.scheme)
        self._apply_transforms(elem)
        return elem

    def to_element(self) -> etree.Element:
        """Return the appropriate element based on mode."""
        if self.scheme:
            return self.to_schemeClr()
        return self.to_srgbClr()

    def with_tint(self, fraction: float) -> Color:
        """Add a tint transform (negative → darken, positive → lighten)."""
        self.tint = fraction
        return self

    def with_shade(self, fraction: float) -> Color:
        """Add a shade transform."""
        self.shade = fraction
        return self

    def with_alpha(self, pct_100k: int) -> Color:
        """Set alpha in 1/1000% units (0–100000)."""
        self.alpha = pct_100k
        return self

    def _apply_transforms(self, elem: etree.Element) -> None:
        if self.tint is not None:
            t = etree.SubElement(elem, self._q("tint"))
            t.set("val", str(int(self.tint * 100_000)))
        if self.shade is not None:
            s = etree.SubElement(elem, self._q("shade"))
            s.set("val", str(int(self.shade * 100_000)))
        if self.alpha is not None:
            a = etree.SubElement(elem, self._q("alpha"))
            a.set("val", str(self.alpha))

    # ---- helpers ---------------------------------------------------------

    @staticmethod
    def wcag_contrast(bg_hex: str) -> str:
        """Return ``"000000"`` (black) or ``"FFFFFF"`` (white) for best contrast.

        Uses WCAG 2.1 relative-luminance formula.
        """
        r, g, b = (
            int(bg_hex[i : i + 2], 16) / 255.0
            for i in (0, 2, 4)
        )

        def linearize(c: float) -> float:
            return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4

        L = 0.2126 * linearize(r) + 0.7152 * linearize(g) + 0.0722 * linearize(b)
        return "000000" if L > 0.179 else "FFFFFF"
