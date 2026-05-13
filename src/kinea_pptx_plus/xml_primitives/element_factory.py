"""Functions to create common OOXML chart elements.

Each function returns a fresh ``lxml.etree.Element`` that can be
appended to an existing chart-space tree.

Example:
    >>> from lxml import etree
    >>> solid = make_solidFill("1F3864")
    >>> etree.tostring(solid, pretty_print=True)
    b'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\\n  <a:srgbClr val="1F3864"/>\\n</a:solidFill>\\n'
"""

from __future__ import annotations

import lxml.etree as etree

from kinea_pptx_plus.xml_primitives.nsmap import A_NS


def _Q(tag: str, ns: str) -> str:
    """Build a fully qualified Clark notation tag."""
    return f"{{{ns}}}{tag}"


def make_solidFill(rgb_hex: str) -> etree.Element:
    """Create ``<a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>``.

    Args:
        rgb_hex: Six-character hex colour string (e.g. ``"1F3864"``).
                 Leading ``#`` is *not* expected.

    Returns:
        An ``lxml.etree.Element`` rooted at ``a:solidFill``.
    """
    root = etree.Element(_Q("solidFill", A_NS), nsmap={"a": A_NS})
    clr = etree.SubElement(root, _Q("srgbClr", A_NS))
    clr.set("val", rgb_hex.upper())
    return root


def make_ln(w_emu: int, rgb_hex: str) -> etree.Element:
    """Create ``<a:ln w="…"><a:solidFill><a:srgbClr val="…"/></a:solidFill></a:ln>``.

    Args:
        w_emu: Line width in EMU (1 pt ≈ 12 700 EMU).
        rgb_hex: Six-character hex colour string.

    Returns:
        An ``lxml.etree.Element`` rooted at ``a:ln``.
    """
    root = etree.Element(_Q("ln", A_NS), nsmap={"a": A_NS})
    root.set("w", str(w_emu))
    # build <a:solidFill><a:srgbClr val="…"/>
    solid = etree.SubElement(root, _Q("solidFill", A_NS))
    clr = etree.SubElement(solid, _Q("srgbClr", A_NS))
    clr.set("val", rgb_hex.upper())
    return root


def make_txPr(text: str = "", font_sz: int = 1200) -> etree.Element:
    """Create ``<c:txPr>`` with optional paragraph body.

    Args:
        text: Run text. Empty string produces an empty paragraph.
        font_sz: Font size in hundredths of a point (default 1200 = 12 pt).

    Returns:
        An ``lxml.etree.Element`` rooted at ``c:txPr``.
    """
    C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    root = etree.Element(
        _Q("txPr", C_NS),
        nsmap={
            "c": C_NS,
            "a": A_NS,
        },
    )
    body = etree.SubElement(root, _Q("bodyPr", A_NS))
    lst = etree.SubElement(root, _Q("lstStyle", A_NS))
    para = etree.SubElement(root, _Q("p", A_NS))
    if text:
        r_elem = etree.SubElement(para, _Q("r", A_NS))
        rpr = etree.SubElement(r_elem, _Q("rPr", A_NS))
        rpr.set("sz", str(font_sz))
        t_elem = etree.SubElement(r_elem, _Q("t", A_NS))
        t_elem.text = text
    return root
