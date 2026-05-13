"""Column annotation overlay — text box with leader line.

Places a ``<p:sp>`` (text box) with an optional
``<p:cxnSp>`` (connector / leader line) pointing from the
annotation to a specific chart column.

Example:
    >>> from kinea_pptx_plus.overlays.column_annotation import add_column_annotation
    >>> from pptx import Presentation
    >>> prs = Presentation("fixture.pptx")
    >>> slide = prs.slides[0]
    >>> add_column_annotation(slide, 5000000, 3000000, 2000000, 500000,
    ...                       text="Key insight", target_x=5000000, target_y=4000000)
    >>> prs.save("out.pptx")
"""

from __future__ import annotations

import lxml.etree as etree

from pptx import Presentation as PptxPresentation

from kinea_pptx_plus.xml_primitives.nsmap import A_NS, P_NS


def add_column_annotation(
    prs_or_slide: PptxPresentation | object,
    x_emu: int,
    y_emu: int,
    cx_emu: int,
    cy_emu: int,
    text: str = "",
    target_x: int | None = None,
    target_y: int | None = None,
    font_size_pt: float = 10.0,
    fill_hex: str = "F2F2F2",
) -> object:
    """Add an annotation text box (with optional leader line) to a slide.

    Args:
        prs_or_slide: Slide or Presentation.
        x_emu, y_emu, cx_emu, cy_emu: Position and size in EMU.
        text:         Annotation text.
        target_x, target_y: If set, draw a connector line from
                           annotation to this point.
        font_size_pt: Font size in points.
        fill_hex:     Background colour of text box.

    Returns:
        The ``p:sp`` element.
    """
    if hasattr(prs_or_slide, "part"):
        slide = prs_or_slide
    else:
        slide = prs_or_slide.slides[0]  # type: ignore[union-attr]

    sp = _build_annotation_sp(x_emu, y_emu, cx_emu, cy_emu,
                              text, font_size_pt, fill_hex)
    sp_tree = slide.shapes._spTree  # type: ignore[attr-defined]
    sp_tree.append(sp)

    # Add leader line if target specified
    if target_x is not None and target_y is not None:
        cxn = _build_leader_cxn_sp(
            x_emu + cx_emu // 2, y_emu + cy_emu // 2,
            target_x, target_y,
        )
        sp_tree.append(cxn)

    return sp


def _build_annotation_sp(
    x: int, y: int, cx: int, cy: int,
    text: str, font_size_pt: float, fill_hex: str,
) -> etree.Element:
    """Build a text box ``<p:sp>``."""
    nsmap = {"p": P_NS, "a": A_NS}
    sp = etree.Element(f"{{{P_NS}}}sp", nsmap=nsmap)

    # Non-visual
    nvSpPr = etree.SubElement(sp, f"{{{P_NS}}}nvSpPr")
    cNvPr = etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvPr")
    cNvPr.set("id", "0")
    cNvPr.set("name", f"Annotation_{int(x)}_{int(y)}")
    cNvSpPr = etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvSpPr")

    # Shape properties
    spPr = etree.SubElement(sp, f"{{{P_NS}}}spPr")
    xfrm = etree.SubElement(spPr, f"{{{A_NS}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{A_NS}}}off")
    off.set("x", str(x)); off.set("y", str(y))
    ext = etree.SubElement(xfrm, f"{{{A_NS}}}ext")
    ext.set("cx", str(cx)); ext.set("cy", str(cy))

    if fill_hex:
        sf = etree.SubElement(spPr, f"{{{A_NS}}}solidFill")
        clr = etree.SubElement(sf, f"{{{A_NS}}}srgbClr")
        clr.set("val", fill_hex.upper())

    prstGeom = etree.SubElement(spPr, f"{{{A_NS}}}prstGeom")
    prstGeom.set("prst", "rect")
    # Rounded corners
    etree.SubElement(prstGeom, f"{{{A_NS}}}avLst")
    ln = etree.SubElement(spPr, f"{{{A_NS}}}ln")
    ln_w = etree.SubElement(ln, f"{{{A_NS}}}solidFill")
    ln_w_clr = etree.SubElement(ln_w, f"{{{A_NS}}}srgbClr")
    ln_w_clr.set("val", "999999")

    # Text body
    txBody = etree.SubElement(sp, f"{{{P_NS}}}txBody")
    bodyPr = etree.SubElement(txBody, f"{{{A_NS}}}bodyPr")
    bodyPr.set("wrap", "square")
    bodyPr.set("rIns", "91440"); bodyPr.set("lIns", "91440")
    bodyPr.set("tIns", "45720"); bodyPr.set("bIns", "45720")
    lstStyle = etree.SubElement(txBody, f"{{{A_NS}}}lstStyle")
    p = etree.SubElement(txBody, f"{{{A_NS}}}p")
    r = etree.SubElement(p, f"{{{A_NS}}}r")
    rPr = etree.SubElement(r, f"{{{A_NS}}}rPr")
    rPr.set("sz", str(int(font_size_pt * 100)))
    rPr.set("spc", "0")
    t = etree.SubElement(r, f"{{{A_NS}}}t")
    t.text = text or ""

    return sp


def _build_leader_cxn_sp(
    start_x: int, start_y: int,
    end_x: int, end_y: int,
) -> etree.Element:
    """Build a connector shape ``<p:cxnSp>`` (leader line)."""
    nsmap = {"p": P_NS, "a": A_NS}
    cxn = etree.Element(f"{{{P_NS}}}cxnSp", nsmap=nsmap)

    nvCxnSpPr = etree.SubElement(cxn, f"{{{P_NS}}}nvCxnSpPr")
    cNvPr = etree.SubElement(nvCxnSpPr, f"{{{P_NS}}}cNvPr")
    cNvPr.set("id", "0")
    cNvPr.set("name", f"Leader_{start_x}-{end_x}")
    cNvCxnSpPr = etree.SubElement(nvCxnSpPr, f"{{{P_NS}}}cNvCxnSpPr")

    spPr = etree.SubElement(cxn, f"{{{P_NS}}}spPr")
    xfrm = etree.SubElement(spPr, f"{{{A_NS}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{A_NS}}}off")
    off.set("x", str(min(start_x, end_x)))
    off.set("y", str(min(start_y, end_y)))
    ext = etree.SubElement(xfrm, f"{{{A_NS}}}ext")
    ext.set("cx", str(abs(end_x - start_x)))
    ext.set("cy", str(abs(end_y - start_y)))

    prstGeom = etree.SubElement(spPr, f"{{{A_NS}}}prstGeom")
    prstGeom.set("prst", "line")
    etree.SubElement(prstGeom, f"{{{A_NS}}}avLst")

    ln = etree.SubElement(spPr, f"{{{A_NS}}}ln")
    ln.set("w", "12700")
    sf = etree.SubElement(ln, f"{{{A_NS}}}solidFill")
    clr = etree.SubElement(sf, f"{{{A_NS}}}srgbClr")
    clr.set("val", "999999")

    # Connection start/end
    stCxn = etree.SubElement(nvCxnSpPr, f"{{{P_NS}}}stCxn")
    stCxn.set("x", str(start_x))
    stCxn.set("y", str(start_y))
    endCxn = etree.SubElement(nvCxnSpPr, f"{{{P_NS}}}endCxn")
    endCxn.set("x", str(end_x))
    endCxn.set("y", str(end_y))

    return cxn
