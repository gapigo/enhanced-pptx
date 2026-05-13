"""Label box overlay — coloured rectangle with auto-contrast text.

Places an ``<p:sp>`` shape on a slide containing a filled rectangle
and a text run.  Text colour (black or white) is chosen for WCAG
contrast against the fill colour.

Example:
    >>> from kinea_pptx_plus.overlays.label_box import add_label_box
    >>> from pptx import Presentation
    >>> prs = Presentation("fixture.pptx")
    >>> slide = prs.slides[0]
    >>> add_label_box(slide, 5000000, 3000000, 1200000, 400000,
    ...               fill_hex="1F3864", text="42.5%")
    >>> prs.save("out.pptx")
"""

from __future__ import annotations

from pptx import Presentation as PptxPresentation
from pptx.oxml.ns import qn
from pptx.util import Emu

from kinea_pptx_plus.xml_primitives.color import Color
from kinea_pptx_plus.xml_primitives.nsmap import A_NS, P_NS


def add_label_box(
    prs_or_slide: PptxPresentation | object,
    x_emu: int,
    y_emu: int,
    cx_emu: int,
    cy_emu: int,
    fill_hex: str,
    text: str,
    font_size_pt: float = 9.0,
) -> object:
    """Add a coloured label box to a slide.

    The shape is positioned at ``(x_emu, y_emu)`` with dimensions
    ``(cx_emu, cy_emu)``.  ``fill_hex`` is a 6-digit hex colour
    (no ``#``).

    Args:
        prs_or_slide: A |python-pptx ``Slide`` or ``Presentation``
                      (if Presentation, uses ``slides[0]``).
        x_emu:        Left edge in EMU.
        y_emu:        Top edge in EMU.
        cx_emu:       Width in EMU.
        cy_emu:       Height in EMU.
        fill_hex:     Background colour (e.g. ``"1F3864"``).
        text:         Label text (e.g. ``"42.5%"``).
        font_size_pt: Font size in points (default 9).

    Returns:
        The created ``p:sp`` XML element (for further manipulation).
    """
    # Resolve slide
    if hasattr(prs_or_slide, "part"):
        slide = prs_or_slide
    else:
        slide = prs_or_slide.slides[0]  # type: ignore[union-attr]

    # --- build the shape XML via python-pptx internals ---
    sp_xml = _build_label_sp(
        x_emu, y_emu, cx_emu, cy_emu, fill_hex, text, font_size_pt
    )

    # Add shape to slide
    sp_tree = slide.shapes._spTree  # type: ignore[attr-defined]
    sp_tree.append(sp_xml)

    return sp_xml


def _build_label_sp(
    x: int, y: int, cx: int, cy: int,
    fill_hex: str, text: str, font_size_pt: float,
) -> object:
    """Build the ``<p:sp>`` XML element for a label box."""
    import lxml.etree as etree

    text_color = Color.wcag_contrast(fill_hex)
    font_emu = int(font_size_pt * 100)  # hundredths of a point

    # Namespace map for serialization
    nsmap = {
        "p": P_NS,
        "a": A_NS,
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }

    sp = etree.Element(f"{{{P_NS}}}sp", nsmap=nsmap)

    # --- non-visual properties ---
    nvSpPr = etree.SubElement(sp, f"{{{P_NS}}}nvSpPr")
    cNvPr = etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvPr")
    cNvPr.set("id", "0")
    cNvPr.set("name", f"LabelBox_{x}_{y}")
    cNvSpPr = etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvSpPr")

    # --- shape properties (fill) ---
    spPr = etree.SubElement(sp, f"{{{P_NS}}}spPr")
    # Transform (position + size)
    xfrm = etree.SubElement(spPr, f"{{{A_NS}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{A_NS}}}off")
    off.set("x", str(x))
    off.set("y", str(y))
    ext = etree.SubElement(xfrm, f"{{{A_NS}}}ext")
    ext.set("cx", str(cx))
    ext.set("cy", str(cy))

    # Solid fill
    solidFill = etree.SubElement(spPr, f"{{{A_NS}}}solidFill")
    srgb = etree.SubElement(solidFill, f"{{{A_NS}}}srgbClr")
    srgb.set("val", fill_hex.upper())

    # Rectangle preset geometry
    prstGeom = etree.SubElement(spPr, f"{{{A_NS}}}prstGeom")
    prstGeom.set("prst", "rect")

    # --- text body ---
    txBody = etree.SubElement(sp, f"{{{P_NS}}}txBody")
    bodyPr = etree.SubElement(txBody, f"{{{A_NS}}}bodyPr")
    bodyPr.set("rIns", "0")
    bodyPr.set("lIns", "27432")  # ~ 0.03 inch
    bodyPr.set("tIns", "0")
    bodyPr.set("bIns", "0")
    bodyPr.set("anchor", "ctr")
    lstStyle = etree.SubElement(txBody, f"{{{A_NS}}}lstStyle")
    p = etree.SubElement(txBody, f"{{{A_NS}}}p")
    pPr = etree.SubElement(p, f"{{{A_NS}}}pPr")
    pPr.set("algn", "ctr")
    r = etree.SubElement(p, f"{{{A_NS}}}r")
    rPr = etree.SubElement(r, f"{{{A_NS}}}rPr")
    rPr.set("sz", str(font_emu))
    rPr.set("b", "1")  # bold
    rPr.set("spc", "0")
    # Set text colour
    rPr_solid = etree.SubElement(rPr, f"{{{A_NS}}}solidFill")
    rPr_clr = etree.SubElement(rPr_solid, f"{{{A_NS}}}srgbClr")
    rPr_clr.set("val", text_color)
    t = etree.SubElement(r, f"{{{A_NS}}}t")
    t.text = text
    # Consume carriage returns so XML stays clean
    if t.text:
        t.text = t.text.replace("\r", "")

    return sp
