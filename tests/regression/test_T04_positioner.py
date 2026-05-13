"""Test Positioner (EMU ↔ data coord) and tree_ops."""

from __future__ import annotations

import lxml.etree as etree

from kinea_pptx_plus.io import extract_chart_xml
from kinea_pptx_plus.overlays.positioner import Positioner
from kinea_pptx_plus.xml_primitives.tree_ops import (
    insert_sorted,
    replace_child,
    deep_clone,
)


def test_positioner_created_from_fixture():
    """Read the fixture chart XML and build a Positioner."""
    xml = extract_chart_xml("tests/fixtures/simple_line_chart.pptx")
    pos = Positioner.from_chart_xml(xml)
    assert pos.plot_cx > 0
    assert pos.plot_cy > 0
    assert len(pos.cat_labels) >= 4
    assert pos.val_max > pos.val_min


def test_positioner_returns_reasonable_emu():
    """data_to_emu should return coordinates inside slide bounds."""
    xml = extract_chart_xml("tests/fixtures/simple_line_chart.pptx")
    pos = Positioner.from_chart_xml(xml)
    emu_x, emu_y = pos.data_to_emu(0, 42.0)
    # EMU for a 10" x 7.5" slide
    assert 0 < emu_x < 9144000, f"EMU X {emu_x} out of bounds"
    assert 0 < emu_y < 6858000, f"EMU Y {emu_y} out of bounds"


def test_positioner_last_cat_has_larger_x():
    """Last category should have larger X than first."""
    xml = extract_chart_xml("tests/fixtures/simple_line_chart.pptx")
    pos = Positioner.from_chart_xml(xml)
    x0, _ = pos.data_to_emu(0, 42.0)
    x_last, _ = pos.data_to_emu(pos.last_cat_index(), 42.0)
    assert x_last > x0


def test_positioner_higher_y_is_smaller_emu():
    """Higher Y value should produce smaller EMU Y (since Y grows downward in PPTX)."""
    xml = extract_chart_xml("tests/fixtures/simple_line_chart.pptx")
    pos = Positioner.from_chart_xml(xml)
    _, y_low = pos.data_to_emu(0, pos.val_min)
    _, y_high = pos.data_to_emu(0, pos.val_max)
    assert y_high < y_low, "Higher data value should map to lower EMU Y"


def test_insert_sorted_appends_when_no_match():
    """With no key, insert_sorted appends."""
    parent = etree.Element("{c}chartSpace")
    child = etree.Element("{c}chart")
    insert_sorted(child, parent, "")
    assert len(parent) == 1
    assert parent[0].tag == "{c}chart"


def test_insert_sorted_after_match():
    parent = etree.Element("{c}chartSpace")
    a = etree.SubElement(parent, "{c}chart")
    b = etree.Element("{c}plotArea")
    insert_sorted(b, parent, "{c}chart")
    assert len(parent) == 2
    assert parent[0].tag == "{c}chart"
    assert parent[1].tag == "{c}plotArea"


def test_replace_child():
    parent = etree.Element("{c}chartSpace")
    old = etree.SubElement(parent, "{c}chart")
    new = etree.Element("{c}plotArea")
    replace_child(parent, old, new)
    assert len(parent) == 1
    assert parent[0].tag == "{c}plotArea"


def test_deep_clone_independent():
    orig = etree.Element("{c}chartSpace")
    etree.SubElement(orig, "{c}chart")
    cloned = deep_clone(orig)
    assert len(cloned) == 1
    # Modify original — clone should be unaffected
    orig.append(etree.Element("{c}plotArea"))
    assert len(cloned) == 1
    assert len(orig) == 2
