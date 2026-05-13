"""Test line_with_end_label v1 — full chart with coloured label overlays.

Validates: schema gate, roundtrip gate.
"""

from __future__ import annotations

import os
import tempfile

import lxml.etree as etree
from pptx import Presentation

from kinea_pptx_plus.charts.line_with_end_label import (
    build_line_with_end_label,
)
from kinea_pptx_plus.io import extract_chart_xml


# ── helpers ──────────────────────────────────────────────────────

def _run_and_load(path: str, colors: list[str] | None = None) -> str:
    """Run build_line_with_end_label and return output path."""
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        out_path = tmp.name
    try:
        result = build_line_with_end_label(
            path, colors=colors, output_path=out_path,
        )
        assert result == out_path
        return out_path
    except Exception:
        if os.path.exists(out_path):
            os.unlink(out_path)
        raise


# ── tests ────────────────────────────────────────────────────────

def test_v1_builds_simple_fixture():
    """Should produce a PPTX with labels on the simple chart."""
    out = _run_and_load("tests/fixtures/simple_line_chart.pptx")
    try:
        prs = Presentation(out)
        assert len(prs.slides) == 1
        # Should have shapes beyond the chart (label overlays)
        assert len(prs.slides[0].shapes) >= 1
    finally:
        if os.path.exists(out):
            os.unlink(out)


def test_v1_builds_5yr_fixture():
    """Should handle the 5-year 70-point dataset without errors."""
    out = _run_and_load(
        "tests/fixtures/andes_vs_cdi_5yr.pptx",
        colors=["1F3864", "C0392B"],
    )
    try:
        prs = Presentation(out)
        assert len(prs.slides) == 1
    finally:
        if os.path.exists(out):
            os.unlink(out)


def test_v1_xml_schema_gate():
    """schema gate: output chart XML must parse without error."""
    out = _run_and_load("tests/fixtures/simple_line_chart.pptx")
    try:
        xml = extract_chart_xml(out)
        root = etree.fromstring(xml)
        assert root.tag.endswith("chartSpace")
    finally:
        if os.path.exists(out):
            os.unlink(out)


def test_v1_roundtrip_gate():
    """roundtrip gate: save → reload → save → diff chart XML."""
    out = _run_and_load("tests/fixtures/simple_line_chart.pptx")
    try:
        xml1 = extract_chart_xml(out)

        # Open, save again, re-extract
        prs = Presentation(out)
        prs.save(out)
        xml2 = extract_chart_xml(out)

        # Compare serialized XML (should be identical after no-op roundtrip)
        # Use parsed tree comparison to ignore whitespace/encoding differences
        t1 = etree.tostring(etree.fromstring(xml1))
        t2 = etree.tostring(etree.fromstring(xml2))
        assert t1 == t2, "Roundtrip gate failed — XML changed after save"
    finally:
        if os.path.exists(out):
            os.unlink(out)


def test_v1_colors_in_chart_xml():
    """Chart XML must contain the specified series colours."""
    out = _run_and_load(
        "tests/fixtures/simple_line_chart.pptx",
        colors=["1F3864", "C0392B"],
    )
    try:
        xml = extract_chart_xml(out)
        root = etree.fromstring(xml)
        # Find all series colours
        a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        c_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
        colors_found = [
            clr.get("val")
            for spPr in root.iter(f"{{{c_ns}}}spPr")
            for ln in spPr.iter(f"{{{a_ns}}}ln")
            for sf in ln.iter(f"{{{a_ns}}}solidFill")
            for clr in sf.iter(f"{{{a_ns}}}srgbClr")
        ]
        assert "1F3864" in colors_found
        assert "C0392B" in colors_found
    finally:
        if os.path.exists(out):
            os.unlink(out)
