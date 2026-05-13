"""Test sql_to_chart — end-to-end chart generation from data."""

from __future__ import annotations

import os
import tempfile

from kinea_pptx_plus.registry.sql_to_chart import chart_from_data, csv_to_chart


def test_chart_from_data_simple():
    """Basic data produces a valid PPTX."""
    data = [("Q1", 42.0, 38.0), ("Q2", 55.0, 50.0)]
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        out = tmp.name
    try:
        result = chart_from_data(data, output_path=out)
        assert os.path.exists(result)
        assert result == out
    finally:
        if os.path.exists(out):
            os.unlink(out)


def test_chart_from_data_5yr():
    """Generate from the 5-year dataset."""
    import sys, importlib
    spec = importlib.util.spec_from_file_location(
        "synth_andes_vs_cdi",
        "scripts/synth_andes_vs_cdi.py"
    )
    synth = importlib.util.module_from_spec(spec)
    sys.modules["synth_andes_vs_cdi"] = synth
    spec.loader.exec_module(synth)
    from datetime import date

    points = synth.generate_series(start=date(2020, 1, 1), end=date(2025, 11, 1), seed=42)
    monthly = synth.to_monthly_endpoints(points)

    data = [(d.isoformat()[:7], f, c) for d, f, c in monthly]
    assert len(data) >= 60

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        out = tmp.name
    try:
        result = chart_from_data(
            data,
            chart_type="line_with_end_label",
            series_names=["Andes", "CDI"],
            colors=["1F3864", "C0392B"],
            output_path=out,
            title="Andes vs CDI — Retorno Acumulado",
        )
        assert os.path.exists(result)
    finally:
        if os.path.exists(out):
            os.unlink(out)


def test_chart_from_data_area_between():
    """Generate area-between chart from data."""
    data = [("Jan", 10.0, 5.0), ("Feb", 20.0, 12.0), ("Mar", 15.0, 8.0)]
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        out = tmp.name
    try:
        result = chart_from_data(
            data,
            chart_type="line_area_between",
            colors=["4472C4", "C0392B"],
            output_path=out,
        )
        assert os.path.exists(result)
    finally:
        if os.path.exists(out):
            os.unlink(out)
