"""Analisi degli ordini aperti (con merce ancora da evadere)."""

from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

from .. import OUTPUT_DIR, REFERENCE_DATE

logger = logging.getLogger(__name__)


def open_orders(orders: pd.DataFrame) -> pd.DataFrame:
    """Filtra le sole righe con QTA INEVASA > 0 o IMPORTO INEVASO > 0."""
    mask = (orders["QTA INEVASA"] > 0) | (orders["IMPORTO INEVASO"] > 0)
    return orders[mask].copy()


def backlog_summary(orders_open: pd.DataFrame) -> pd.DataFrame:
    """Backlog totale e numero ordini distinti aperti."""
    n_orders = orders_open[["ANNO", "TIPO", "NUM."]].drop_duplicates().shape[0]
    return pd.DataFrame(
        [{
            "BACKLOG_TOTALE": round(float(orders_open["IMPORTO INEVASO"].sum()), 2),
            "QTA_INEVASA_TOTALE": int(orders_open["QTA INEVASA"].sum()),
            "ORDINI_APERTI_DISTINTI": int(n_orders),
            "RIGHE_APERTE": int(len(orders_open)),
        }]
    )


def backlog_by_client(orders_open: pd.DataFrame) -> pd.DataFrame:
    """Backlog per cliente, ordinato decrescente."""
    g = (
        orders_open.groupby(["CLIENTE", "CLIENTE.1"], dropna=False)
        .agg(
            BACKLOG=("IMPORTO INEVASO", "sum"),
            QTA_INEVASA=("QTA INEVASA", "sum"),
            RIGHE=("IMPORTO INEVASO", "size"),
        )
        .reset_index()
        .rename(columns={"CLIENTE.1": "RAGIONE_SOCIALE"})
        .sort_values("BACKLOG", ascending=False)
    )
    g["BACKLOG"] = g["BACKLOG"].round(2)
    return g.reset_index(drop=True)


def backlog_by_macro(orders_open: pd.DataFrame) -> pd.DataFrame:
    """Backlog per macro-famiglia articolo (DESCRIZIONE ELEMENTO.5 negli ordini)."""
    macro_col = "DESCRIZIONE ELEMENTO.5"
    g = (
        orders_open.groupby([macro_col], dropna=False)["IMPORTO INEVASO"]
        .sum()
        .reset_index()
        .rename(columns={macro_col: "MACRO_FAMIGLIA", "IMPORTO INEVASO": "BACKLOG"})
        .sort_values("BACKLOG", ascending=False)
    )
    g["BACKLOG"] = g["BACKLOG"].round(2)
    return g.reset_index(drop=True)


def backlog_by_geography(orders_open: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Backlog per provincia e per regione (SOTTOAREA / DESCRIZIONE ELEMENTO.1)."""
    by_province = (
        orders_open.groupby(["PROVINCIA"], dropna=False)["IMPORTO INEVASO"]
        .sum()
        .reset_index()
        .rename(columns={"IMPORTO INEVASO": "BACKLOG"})
        .sort_values("BACKLOG", ascending=False)
        .reset_index(drop=True)
    )
    by_province["BACKLOG"] = by_province["BACKLOG"].round(2)

    by_region = (
        orders_open.groupby(["SOTTOAREA", "DESCRIZIONE ELEMENTO.1"], dropna=False)["IMPORTO INEVASO"]
        .sum()
        .reset_index()
        .rename(columns={"DESCRIZIONE ELEMENTO.1": "REGIONE", "IMPORTO INEVASO": "BACKLOG"})
        .sort_values("BACKLOG", ascending=False)
        .reset_index(drop=True)
    )
    by_region["BACKLOG"] = by_region["BACKLOG"].round(2)

    return {"provincia": by_province, "regione": by_region}


def delivery_buckets(
    orders_open: pd.DataFrame, reference_date: str | None = None
) -> pd.DataFrame:
    """Distribuzione per data di consegna (in ritardo / settimana / 30gg / oltre)."""
    ref = pd.Timestamp(reference_date or REFERENCE_DATE)
    df = orders_open.copy()

    def bucket(d: pd.Timestamp) -> str:
        if pd.isna(d):
            return "Senza data"
        if d < ref:
            return "In ritardo"
        delta = (d - ref).days
        if delta <= 7:
            return "Questa settimana"
        if delta <= 30:
            return "Prossimi 30 gg"
        return "Oltre 30 gg"

    df["FASCIA_CONSEGNA"] = df["DATA CONSEGNA"].apply(bucket)
    g = (
        df.groupby("FASCIA_CONSEGNA")
        .agg(BACKLOG=("IMPORTO INEVASO", "sum"), RIGHE=("IMPORTO INEVASO", "size"))
        .reset_index()
    )
    order = ["In ritardo", "Questa settimana", "Prossimi 30 gg", "Oltre 30 gg", "Senza data"]
    g["FASCIA_CONSEGNA"] = pd.Categorical(g["FASCIA_CONSEGNA"], categories=order, ordered=True)
    g = g.sort_values("FASCIA_CONSEGNA").reset_index(drop=True)
    g["BACKLOG"] = g["BACKLOG"].round(2)
    g["FASCIA_CONSEGNA"] = g["FASCIA_CONSEGNA"].astype(str)
    return g


def order_age(orders_open: pd.DataFrame, reference_date: str | None = None) -> pd.DataFrame:
    """Età ordini in giorni dalla creazione, con classi di anzianità."""
    ref = pd.Timestamp(reference_date or REFERENCE_DATE)
    df = orders_open.copy()
    df["GIORNI_DA_CREAZIONE"] = (ref - df["DATA CREAZIONE"]).dt.days

    def age_class(d: float) -> str:
        if pd.isna(d):
            return "Senza data"
        if d < 30:
            return "<30"
        if d < 60:
            return "30-60"
        if d < 90:
            return "60-90"
        return ">90"

    df["FASCIA_ETA"] = df["GIORNI_DA_CREAZIONE"].apply(age_class)
    g = (
        df.groupby("FASCIA_ETA")
        .agg(BACKLOG=("IMPORTO INEVASO", "sum"), RIGHE=("IMPORTO INEVASO", "size"))
        .reset_index()
    )
    order = ["<30", "30-60", "60-90", ">90", "Senza data"]
    g["FASCIA_ETA"] = pd.Categorical(g["FASCIA_ETA"], categories=order, ordered=True)
    g = g.sort_values("FASCIA_ETA").reset_index(drop=True)
    g["BACKLOG"] = g["BACKLOG"].round(2)
    g["FASCIA_ETA"] = g["FASCIA_ETA"].astype(str)
    return g


def top_articles_open(orders_open: pd.DataFrame, top_n: int = 20) -> pd.DataFrame:
    """Top articoli più ordinati e ancora da evadere."""
    g = (
        orders_open.groupby(["ARTICOLO", "DESCRIZIONE"], dropna=False)
        .agg(
            QTA_INEVASA=("QTA INEVASA", "sum"),
            BACKLOG=("IMPORTO INEVASO", "sum"),
            RIGHE=("IMPORTO INEVASO", "size"),
        )
        .reset_index()
        .sort_values("QTA_INEVASA", ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )
    g["BACKLOG"] = g["BACKLOG"].round(2)
    return g


def build_all(
    orders: pd.DataFrame,
    output_path: Path | None = None,
    reference_date: str | None = None,
) -> dict[str, pd.DataFrame]:
    """Calcola tutte le analisi ordini aperti e le scrive in un Excel multi-foglio."""
    output_path = output_path or (OUTPUT_DIR / "analisi_ordini.xlsx")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    open_df = open_orders(orders)
    geo = backlog_by_geography(open_df)

    results: dict[str, pd.DataFrame] = {
        "riepilogo_backlog": backlog_summary(open_df),
        "backlog_per_cliente": backlog_by_client(open_df),
        "backlog_per_macro": backlog_by_macro(open_df),
        "backlog_per_provincia": geo["provincia"],
        "backlog_per_regione": geo["regione"],
        "scadenze_consegna": delivery_buckets(open_df, reference_date),
        "eta_ordini": order_age(open_df, reference_date),
        "top20_articoli_aperti": top_articles_open(open_df, 20),
    }

    logger.info("Scrittura Excel analisi ordini -> %s", output_path)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in results.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    csv_dir = output_path.parent / "ordini_csv"
    csv_dir.mkdir(parents=True, exist_ok=True)
    for name, df in results.items():
        df.to_csv(csv_dir / f"{name}.csv", index=False)

    return results
