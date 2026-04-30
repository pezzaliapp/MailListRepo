"""Test del modulo mailing.builder con fixture sintetiche."""

from __future__ import annotations

import pandas as pd

from src.mailing.builder import (
    _is_valid_email,
    _split_emails,
    aggregate_clients,
    build_all,
    build_mailing_list,
)


def test_split_emails_handles_separators() -> None:
    assert _split_emails("a@x.com - b@x.com") == ["a@x.com", "b@x.com"]
    assert _split_emails("a@x.com;b@x.com") == ["a@x.com", "b@x.com"]
    assert _split_emails("a@x.com, b@x.com") == ["a@x.com", "b@x.com"]
    assert _split_emails(None) == []
    assert _split_emails("") == []
    assert _split_emails("nan") == []


def test_split_emails_dedups_case_insensitive() -> None:
    out = _split_emails("Foo@Bar.com; foo@bar.com")
    assert out == ["Foo@Bar.com"]


def test_is_valid_email() -> None:
    assert _is_valid_email("valid@example.com")
    assert not _is_valid_email("non-una-email")
    assert not _is_valid_email("@example.com")


def test_aggregate_clients_dedups_emails(synthetic_clients: pd.DataFrame) -> None:
    agg = aggregate_clients(synthetic_clients)
    # Cliente 7 ha eta@example.com e ETA@example.com → una sola dopo dedup.
    eta = agg[agg["codice_cliente"] == 7].iloc[0]
    assert len(eta["emails"]) == 1
    # Cliente 2 ha 2 email nella stessa cella.
    beta = agg[agg["codice_cliente"] == 2].iloc[0]
    assert len(beta["emails"]) == 2


def test_build_mailing_list_segments_and_flags(
    synthetic_clients: pd.DataFrame,
    synthetic_sales: pd.DataFrame,
    synthetic_orders: pd.DataFrame,
) -> None:
    m = build_mailing_list(
        synthetic_clients, synthetic_sales, synthetic_orders, reference_date="2026-04-30"
    )

    # Cliente 1: solo storico.
    r1 = m[m["codice_cliente"] == 1].iloc[0]
    assert r1["ha_acquistato_storico"] and not r1["ha_ordine_aperto"]
    # Cliente 2: solo ordine aperto.
    r2 = m[m["codice_cliente"] == 2].iloc[0]
    assert r2["ha_ordine_aperto"] and not r2["ha_acquistato_storico"]
    # Cliente 3: entrambi.
    r3 = m[m["codice_cliente"] == 3].iloc[0]
    assert r3["ha_acquistato_storico"] and r3["ha_ordine_aperto"]
    # Cliente 4: nessuno.
    r4 = m[m["codice_cliente"] == 4].iloc[0]
    assert not r4["ha_acquistato_storico"] and not r4["ha_ordine_aperto"]
    # Cliente 5: storico ma email vuota → email_valida = False, email = "".
    r5 = m[m["codice_cliente"] == 5].iloc[0]
    assert r5["email"] == "" and not r5["email_valida"]
    # Cliente 6: storico ma email non valida → flag False ma email mantenuta.
    r6 = m[m["codice_cliente"] == 6].iloc[0]
    assert r6["email"] == "non-una-email" and not r6["email_valida"]


def test_build_all_creates_files(
    synthetic_clients: pd.DataFrame,
    synthetic_sales: pd.DataFrame,
    synthetic_orders: pd.DataFrame,
    tmp_output_dir,
) -> None:
    res = build_all(
        synthetic_clients, synthetic_sales, synthetic_orders,
        output_dir=tmp_output_dir, reference_date="2026-04-30",
    )
    assert (tmp_output_dir / "mailing_list_acquirenti.csv").exists()
    assert (tmp_output_dir / "mailing_list_ordini_aperti.csv").exists()
    assert (tmp_output_dir / "mailing_list_completa.xlsx").exists()
    seg = res["segments"]
    # Cliente 5 (storico ma senza email) deve apparire in "senza_email".
    assert 5 in seg["senza_email"]["codice_cliente"].tolist()
