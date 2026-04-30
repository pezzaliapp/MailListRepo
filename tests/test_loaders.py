"""Test del modulo loaders contro i file reali in data/."""

from __future__ import annotations

import pandas as pd
import pytest

from src import CLIENTS_FILE, ORDERS_FILE, SALES_FILE
from src.loaders import load_clients, load_orders, load_sales

real_data_present = (
    CLIENTS_FILE.exists() and ORDERS_FILE.exists() and SALES_FILE.exists()
)


@pytest.mark.skipif(not real_data_present, reason="File reali non presenti")
def test_load_clients_shape_and_keys() -> None:
    df = load_clients()
    assert {"CONTATTO CLIENTI", "RAGIONE SOCIALE 1", "EMAIL"}.issubset(df.columns)
    assert df["CONTATTO CLIENTI"].dtype == "int64"
    assert df["CONTATTO CLIENTI"].nunique() > 0


@pytest.mark.skipif(not real_data_present, reason="File reali non presenti")
def test_load_orders_basics() -> None:
    df = load_orders()
    assert df["CLIENTE"].dtype == "int64"
    assert pd.api.types.is_datetime64_any_dtype(df["DATA CREAZIONE"])
    assert df["IMPORTO INEVASO"].dtype.kind == "f"


@pytest.mark.skipif(not real_data_present, reason="File reali non presenti")
def test_load_sales_returns_flag_and_net_revenue() -> None:
    df = load_sales()
    assert "IS_RESO" in df.columns
    assert "IMPORTO_NETTO" in df.columns
    # I resi devono avere importo netto <= 0.
    resi = df[df["IS_RESO"]]
    assert (resi["IMPORTO_NETTO"] <= 0).all()
