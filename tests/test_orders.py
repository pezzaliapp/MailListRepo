"""Test del modulo analytics.orders con fixture sintetiche."""

from __future__ import annotations

import pandas as pd

from src.analytics.orders import (
    backlog_by_client,
    backlog_summary,
    build_all,
    delivery_buckets,
    open_orders,
    order_age,
)


def test_open_orders_filter(synthetic_orders: pd.DataFrame) -> None:
    op = open_orders(synthetic_orders)
    # Tutte le righe sintetiche hanno QTA INEVASA > 0.
    assert len(op) == len(synthetic_orders)


def test_backlog_summary(synthetic_orders: pd.DataFrame) -> None:
    s = backlog_summary(open_orders(synthetic_orders))
    assert int(s.iloc[0]["BACKLOG_TOTALE"]) == 4300  # 2000 + 1500 + 800
    assert int(s.iloc[0]["ORDINI_APERTI_DISTINTI"]) == 3


def test_backlog_by_client(synthetic_orders: pd.DataFrame) -> None:
    g = backlog_by_client(open_orders(synthetic_orders))
    # Cliente 3 ha 1500 + 800 = 2300, cliente 2 ha 2000 → top è 3.
    assert g.iloc[0]["CLIENTE"] == 3


def test_delivery_buckets(synthetic_orders: pd.DataFrame) -> None:
    b = delivery_buckets(open_orders(synthetic_orders), reference_date="2026-04-30")
    in_ritardo = b[b["FASCIA_CONSEGNA"] == "In ritardo"]
    assert not in_ritardo.empty
    # 1500 (consegna 2026-04-20) deve essere in ritardo
    assert float(in_ritardo.iloc[0]["BACKLOG"]) == 1500.0


def test_order_age(synthetic_orders: pd.DataFrame) -> None:
    a = order_age(open_orders(synthetic_orders), reference_date="2026-04-30")
    assert a["BACKLOG"].sum() == 4300.0


def test_build_all_writes_excel(synthetic_orders: pd.DataFrame, tmp_output_dir) -> None:
    out = tmp_output_dir / "ordini.xlsx"
    res = build_all(synthetic_orders, output_path=out, reference_date="2026-04-30")
    assert out.exists()
    assert "riepilogo_backlog" in res
