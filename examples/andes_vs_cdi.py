"""End-to-end example: Andes vs CDI - Retorno Acumulado.

Gera 5+ anos de dados sinteticos, cria um grafico line_with_end_label
com labels coloridas no final de cada serie, e salva como PPTX.

Uso:
    python examples/andes_vs_cdi.py
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scripts.synth_andes_vs_cdi import generate_series, to_monthly_endpoints
from datetime import date
from kinea_pptx_plus.registry.sql_to_chart import chart_from_data


def main():
    # 1. Generate 5+ years of synthetic data
    print("Generating synthetic Andes vs CDI data...")
    points = generate_series(start=date(2020, 1, 1), end=date(2025, 11, 1), seed=42)
    monthly = to_monthly_endpoints(points)
    print(f"  {len(monthly)} monthly points ({monthly[0][0]} to {monthly[-1][0]})")

    # 2. Format for chart_from_data
    data = [(d.isoformat()[:7], f, c) for d, f, c in monthly]

    # 3. Generate the chart
    output_path = "andes_vs_cdi_output.pptx"
    print(f"Generating chart -> {output_path}")

    result = chart_from_data(
        data,
        chart_type="line_with_end_label",
        series_names=["Andes", "CDI"],
        colors=["1F3864", "C0392B"],
        output_path=output_path,
        title="Andes vs CDI - Retorno Acumulado",
    )

    print(f"Done! Output: {result}")
    print(f"  Size: {os.path.getsize(result)} bytes")
    print(f"  Open in PowerPoint to verify.")


if __name__ == "__main__":
    main()
