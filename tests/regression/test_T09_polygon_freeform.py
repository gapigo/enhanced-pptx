"""Test overlays.polygon_freeform — freeform polygon shapes."""

from __future__ import annotations

import tempfile
import os

import lxml.etree as etree
from pptx import Presentation

from kinea_pptx_plus.overlays.polygon_freeform import (
    add_freeform_polygon,
    _build_freeform_sp,
)


def test_build_freeform_sp_creates_valid_xml():
    """The generated p:sp XML should parse cleanly."""
    sp = _build_freeform_sp(
        vertices=[(0, 0), (1000000, 0), (500000, 500000)],
        fill_color="4472C4", fill_alpha=40000,
    )
    xml = etree.tostring(sp)
    # Should have custGeom
    assert b"custGeom" in xml
    # Should have fill colour
    assert b"4472C4" in xml
    # Should have alpha
    assert b"40000" in xml


def test_triangle_polygon():
    """A triangle should have 3 moveTo + 3 lnTo + 1 close."""
    sp = _build_freeform_sp(
        vertices=[(0, 0), (1000000, 0), (500000, 500000), (0, 0)],
        fill_color="FF0000", fill_alpha=50000,
    )
    xml = etree.tostring(sp)
    # Should have moveTo, lnTo, close
    assert b"moveTo" in xml
    assert b"lnTo" in xml
    assert b"close" in xml


def test_50_vertex_polygon():
    """A 50-vertex polygon should render correctly."""
    vertices = [(i * 100000, (i % 5) * 100000) for i in range(50)]
    vertices.append(vertices[0])  # close
    sp = _build_freeform_sp(
        vertices=vertices,
        fill_color="1F3864", fill_alpha=30000,
    )
    xml = etree.tostring(sp)
    assert b"custGeom" in xml


def test_add_polygon_to_slide():
    """Add polygon to a real slide and verify PPTX saves."""
    prs = Presentation()
    prs.slide_width = 9144000
    prs.slide_height = 6858000
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_freeform_polygon(
        slide,
        vertices=[(1000000, 1000000), (2000000, 1000000),
                  (1500000, 2000000), (1000000, 1000000)],
        fill_color="4472C4", fill_alpha=40000,
    )

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        out_path = tmp.name
    try:
        prs.save(out_path)
        prs2 = Presentation(out_path)
        assert len(prs2.slides) == 1
    finally:
        if os.path.exists(out_path):
            os.unlink(out_path)
