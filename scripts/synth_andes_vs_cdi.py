"""Gera dataset sintético no estilo do gráfico Andes vs CDI.

Roda standalone para popular tests/fixtures/ com CSV pronto.
Em produção, o sql_to_chart vai conectar no fin-data-lab e substituir isso.

Usage:
    python scripts/synth_andes_vs_cdi.py > tests/fixtures/andes_vs_cdi.csv
"""

from __future__ import annotations
import csv
import sys
from datetime import date
from random import Random


def generate_series(
    start: date = date(2021, 9, 1),
    end: date = date(2026, 4, 1),
    cdi_annual: float = 0.135,
    fund_alpha_annual: float = 0.02,
    fund_vol_daily: float = 0.0035,
    seed: int = 42,
) -> list[tuple[date, float, float]]:
    """Gera retornos acumulados diários de um fundo vs CDI.
    
    Retorna lista de (date, fund_cumulative_pct, cdi_cumulative_pct).
    Valores partem de 0 e crescem como o gráfico exemplo.
    """
    rng = Random(seed)
    cdi_daily = (1 + cdi_annual) ** (1 / 252) - 1
    fund_drift = cdi_daily + (fund_alpha_annual / 252)
    
    out = []
    cur = start
    fund_acc = 1.0
    cdi_acc = 1.0
    while cur <= end:
        if cur.weekday() < 5:
            shock = rng.gauss(0, fund_vol_daily)
            fund_acc *= 1 + fund_drift + shock
            cdi_acc *= 1 + cdi_daily
            out.append((cur, (fund_acc - 1) * 100, (cdi_acc - 1) * 100))
        cur = date.fromordinal(cur.toordinal() + 1)
    return out


def to_monthly_endpoints(daily):
    """Reduz pra um ponto por mês (último dia útil)."""
    seen = set()
    monthly = []
    for d, f, c in daily:
        key = (d.year, d.month)
        if key not in seen:
            seen.add(key)
            monthly.append((d, f, c))
    # último dia útil de cada mês: pega o último de cada bucket
    last = {}
    for d, f, c in daily:
        last[(d.year, d.month)] = (d, f, c)
    return [last[k] for k in sorted(last)]


if __name__ == "__main__":
    daily = generate_series()
    monthly = to_monthly_endpoints(daily)
    writer = csv.writer(sys.stdout)
    writer.writerow(["date", "andes_pct", "cdi_pct"])
    for d, f, c in monthly:
        writer.writerow([d.isoformat(), f"{f:.2f}", f"{c:.2f}"])
