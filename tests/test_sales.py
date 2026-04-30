"""Test del modulo analytics.sales con fixture sintetiche."""

from __future__ import annotations

import pandas as pd

from src.analytics.sales import (
    build_all,
    dormant_clients,
    purchase_frequency,
    returns_rate,
    revenue_by_year,
    top_clients,
)


def test_revenue_by_year_nets_returns(synthetic_sales: pd.DataFrame) -> None:
    g = revenue_by_year(synthetic_sales)
    # 2025: 1000 (Alpha) - 200 (reso Alpha) + 300 (Zeta) = 1100
    val_2025 = float(g.loc[g["ANNO"] == 2025, "FATTURATO_NETTO"].iloc[0])
    assert val_2025 == 1100.0


def test_top_clients_orders_by_revenue(synthetic_sales: pd.DataFrame) -> None:
    top = top_clients(synthetic_sales, top_n=5)
    # Cliente 1 dovrebbe avere fatturato 1000 - 200 = 800; cliente 7 ha 900.
    first = top.iloc[0]["CLIENTE"]
    assert first == 7  # Eta SRL: 900


def test_returns_rate_year_2025(synthetic_sales: pd.DataFrame) -> None:
    rates = returns_rate(synthetic_sales)
    by_year = rates["by_year"]
    row = by_year[by_year["ANNO"] == 2025].iloc[0]
    assert row["RESI"] == -200.0
    assert row["VENDITE"] == 1300.0  # 1000 + 300


def test_dormant_clients_excludes_recent(synthetic_sales: pd.DataFrame) -> None:
    dorm12 = dormant_clients(synthetic_sales, months=12, reference_date="2026-04-30")
    codes = set(dorm12["CLIENTE"].tolist())
    # 5 (2020) e 7 (giu 2024) sono dormienti a 12m.
    assert 5 in codes
    assert 7 in codes
    # 1, 3, 6 hanno acquistato negli ultimi 12 mesi.
    assert 3 not in codes
    assert 6 not in codes


def test_purchase_frequency(synthetic_sales: pd.DataFrame) -> None:
    pf = purchase_frequency(synthetic_sales)
    # Tutti devono avere almeno 1 transazione.
    assert (pf["NUMERO_TRANSAZIONI"] >= 1).all()


def test_build_all_writes_excel(synthetic_sales: pd.DataFrame, tmp_output_dir) -> None:
    out = tmp_output_dir / "vendite.xlsx"
    res = build_all(synthetic_sales, output_path=out, reference_date="2026-04-30")
    assert out.exists()
    assert "fatturato_per_anno" in res
