"""Freeform polygon overlay for shaded areas between chart series.

Creates an ``<a:custGeom>`` shape that renders as a filled polygon
with optional transparency.

Example:
    >>> from kinea_pptx_plus.overlays.polygon_freeform import add_freeform_polygon
    >>> from pptx import Presentation
    >>> prs = Presentation()
    >>> slide = prs.slides.add_slide(prs.slide_layouts[6])
    >>> add_freeform_polygon(slide, vertices=[(0,0), (1000000,0), (500000,500000)],
    ...                      fill_color="4472C4", fill_alpha=40000)
    >>> prs.save("out.pptx")
"""

from __future__ import annotations

import lxml.etree as etree

from pptx import Presentation as PptxPresentation

from kinea_pptx_plus.xml_primitives.nsmap import A_NS, P_NS


def add_freeform_polygon(
    prs_or_slide: PptxPresentation | object,
    vertices: list[tuple[int, int]],
    fill_color: str = "4472C4",
    fill_alpha: int = 40000,
) -> object:
    """Add a filled freeform polygon shape to a slide.

    Args:
        prs_or_slide: Slide or Presentation (uses ``slides[0]``).
        vertices:     List of ``(x_emu, y_emu)`` points.
        fill_color:   Hex colour string (e.g. ``"4472C4"``).
        fill_alpha:   Alpha in 1/1000% units (0-100000).

    Returns:
        The ``p:sp`` XML element (for further manipulation).
    """
    if hasattr(prs_or_slide, "part"):
        slide = prs_or_slide
    else:
        slide = prs_or_slide.slides[0]  # type: ignore[union-attr]

    sp = _build_freeform_sp(vertices, fill_color, fill_alpha)
    sp_tree = slide.shapes._spTree  # type: ignore[attr-defined]
    sp_tree.append(sp)
    return sp


def _build_freeform_sp(
    vertices: list[tuple[int, int]],
    fill_color: str,
    fill_alpha: int,
) -> etree.Element:
    """Build the ``<p:sp>`` element with ``<a:custGeom>``."""
    nsmap = {
        "p": P_NS,
        "a": A_NS,
    }

    sp = etree.Element(f"{{{P_NS}}}sp", nsmap=nsmap)

    # Non-visual properties
    nvSpPr = etree.SubElement(sp, f"{{{P_NS}}}nvSpPr")
    cNvPr = etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvPr")
    cNvPr.set("id", "0")
    cNvPr.set("name", f"FreeformPoly_{id(vertices)}")
    cNvSpPr = etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvSpPr")

    # Shape properties
    spPr = etree.SubElement(sp, f"{{{P_NS}}}spPr")

    # Transform: compute bounding box from vertices
    xs = [v[0] for v in vertices]
    ys = [v[1] for v in vertices]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)

    xfrm = etree.SubElement(spPr, f"{{{A_NS}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{A_NS}}}off")
    off.set("x", str(min_x))
    off.set("y", str(min_y))
    ext = etree.SubElement(xfrm, f"{{{A_NS}}}ext")
    ext.set("cx", str(max_x - min_x))
    ext.set("cy", str(max_y - min_y))

    # Solid fill with optional alpha
    solidFill = etree.SubElement(spPr, f"{{{A_NS}}}solidFill")
    srgb = etree.SubElement(solidFill, f"{{{A_NS}}}srgbClr")
    srgb.set("val", fill_color.upper())
    if fill_alpha < 100000:
        alpha = etree.SubElement(srgb, f"{{{A_NS}}}alpha")
        alpha.set("val", str(fill_alpha))

    # Custom geometry (freeform polygon)
    custGeom = etree.SubElement(spPr, f"{{{A_NS}}}custGeom")
    avLst = etree.SubElement(custGeom, f"{{{A_NS}}}avLst")
    pathLst = etree.SubElement(custGeom, f"{{{A_NS}}}pathLst")
    path = etree.SubElement(pathLst, f"{{{A_NS}}}path")
    path.set("w", str(max_x - min_x))
    path.set("h", str(max_y - min_y))

    # First vertex: moveTo
    first = vertices[0]
    moveTo = etree.SubElement(path, f"{{{A_NS}}}moveTo")
    movePt = etree.SubElement(moveTo, f"{{{A_NS}}}pt")
    movePt.set("x", str(first[0] - min_x))
    movePt.set("y", str(first[1] - min_y))

    # Subsequent vertices: lnTo
    for v in vertices[1:]:
        lnTo = etree.SubElement(path, f"{{{A_NS}}}lnTo")
        pt = etree.SubElement(lnTo, f"{{{A_NS}}}pt")
        pt.set("x", str(v[0] - min_x))
        pt.set("y", str(v[1] - min_y))

    # Close path
    etree.SubElement(path, f"{{{A_NS}}}close")

    return sp
