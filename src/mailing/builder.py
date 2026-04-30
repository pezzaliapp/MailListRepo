"""Costruzione mailing list incrociando anagrafica, vendite e ordini."""

from __future__ import annotations

import logging
import re
from pathlib import Path

import pandas as pd
from dateutil.relativedelta import relativedelta

from .. import OUTPUT_DIR, REFERENCE_DATE

logger = logging.getLogger(__name__)

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
EMAIL_SPLIT_RE = re.compile(r"[;,/\s]+|\s-\s")


def _split_emails(raw: str | float | None) -> list[str]:
    """Divide una cella che può contenere più email separate da `;`, `,`, spazi o ` - `.

    Args:
        raw: contenuto grezzo della cella EMAIL.

    Returns:
        Lista di email pulite e deduplicate, o lista vuota se nulla.
    """
    if raw is None:
        return []
    try:
        if pd.isna(raw):
            return []
    except (TypeError, ValueError):
        pass
    s = str(raw).strip()
    if not s or s.lower() in {"nan", "<na>", "none"}:
        return []
    parts = [p.strip(" ;,-") for p in EMAIL_SPLIT_RE.split(s)]
    out: list[str] = []
    seen: set[str] = set()
    for p in parts:
        if not p:
            continue
        low = p.lower()
        if low in seen:
            continue
        seen.add(low)
        out.append(p)
    return out


def _is_valid_email(email: str) -> bool:
    return bool(EMAIL_RE.match(email))


def aggregate_clients(clients: pd.DataFrame) -> pd.DataFrame:
    """Aggrega anagrafica clienti per `CONTATTO CLIENTI`.

    Returns:
        DataFrame con una riga per cliente: ragione sociale, lista email,
        flag email_valida, telefoni concatenati.
    """
    rows = []
    for code, group in clients.groupby("CONTATTO CLIENTI"):
        emails: list[str] = []
        for raw in group["EMAIL"].tolist():
            emails.extend(_split_emails(raw))
        # Dedup case-insensitive preservando primo casing.
        seen = set()
        emails_unique: list[str] = []
        for e in emails:
            low = e.lower()
            if low not in seen:
                seen.add(low)
                emails_unique.append(e)

        ragione = group["RAGIONE SOCIALE 1"].dropna().iloc[0] if not group["RAGIONE SOCIALE 1"].dropna().empty else ""
        telefoni = group[["NR.TELEFONICO", "NR.CELLULARE"]].fillna("").agg(" / ".join, axis=1)
        telefoni_str = "; ".join(t for t in telefoni.unique() if t.strip(" /"))

        rows.append({
            "codice_cliente": int(code),
            "ragione_sociale": ragione,
            "emails": emails_unique,
            "telefoni": telefoni_str,
        })
    return pd.DataFrame(rows)


def compute_client_metrics(
    sales: pd.DataFrame,
    orders: pd.DataFrame,
    reference_date: str | None = None,
) -> pd.DataFrame:
    """Calcola metriche per cliente: fatturato, ultimo acquisto, backlog, ordini aperti."""
    ref = pd.Timestamp(reference_date or REFERENCE_DATE)
    cutoff_12m = ref - relativedelta(months=12)

    sales_total = (
        sales.groupby("CLIENTE")["IMPORTO_NETTO"].sum().reset_index()
        .rename(columns={"CLIENTE": "codice_cliente", "IMPORTO_NETTO": "fatturato_totale"})
    )
    sales_12m = (
        sales[sales["DATA SPEDIZIONE"] >= cutoff_12m]
        .groupby("CLIENTE")["IMPORTO_NETTO"].sum().reset_index()
        .rename(columns={"CLIENTE": "codice_cliente", "IMPORTO_NETTO": "fatturato_12m"})
    )
    last_purchase = (
        sales.groupby("CLIENTE")["DATA SPEDIZIONE"].max().reset_index()
        .rename(columns={"CLIENTE": "codice_cliente", "DATA SPEDIZIONE": "data_ultimo_acquisto"})
    )

    open_mask = (orders["QTA INEVASA"] > 0) | (orders["IMPORTO INEVASO"] > 0)
    open_o = orders[open_mask]
    backlog = (
        open_o.groupby("CLIENTE")
        .agg(
            backlog_attuale=("IMPORTO INEVASO", "sum"),
            n_ordini_aperti=("NUM.", lambda s: s.nunique()),
        )
        .reset_index()
        .rename(columns={"CLIENTE": "codice_cliente"})
    )

    # Provincia: la prendiamo dagli ordini se disponibile, altrimenti dalle vendite (DESCRIZIONE ELEMENTO.1).
    prov_orders = (
        orders.dropna(subset=["PROVINCIA"]).groupby("CLIENTE")["PROVINCIA"]
        .first().reset_index()
        .rename(columns={"CLIENTE": "codice_cliente", "PROVINCIA": "provincia"})
    )
    prov_sales = (
        sales.dropna(subset=["DESCRIZIONE ELEMENTO.1"])
        .groupby("CLIENTE")["DESCRIZIONE ELEMENTO.1"]
        .first().reset_index()
        .rename(columns={"CLIENTE": "codice_cliente", "DESCRIZIONE ELEMENTO.1": "provincia"})
    )
    provinces = pd.concat([prov_orders, prov_sales]).drop_duplicates("codice_cliente")

    metrics = (
        sales_total.merge(sales_12m, on="codice_cliente", how="outer")
        .merge(last_purchase, on="codice_cliente", how="outer")
        .merge(backlog, on="codice_cliente", how="outer")
        .merge(provinces, on="codice_cliente", how="left")
    )

    metrics["fatturato_totale"] = metrics["fatturato_totale"].fillna(0.0).round(2)
    metrics["fatturato_12m"] = metrics["fatturato_12m"].fillna(0.0).round(2)
    metrics["backlog_attuale"] = metrics["backlog_attuale"].fillna(0.0).round(2)
    metrics["n_ordini_aperti"] = metrics["n_ordini_aperti"].fillna(0).astype(int)

    metrics["ha_acquistato_storico"] = metrics["fatturato_totale"] != 0.0
    metrics["ha_acquistato_12m"] = metrics["fatturato_12m"] != 0.0
    metrics["ha_ordine_aperto"] = metrics["n_ordini_aperti"] > 0

    return metrics


def build_mailing_list(
    clients: pd.DataFrame,
    sales: pd.DataFrame,
    orders: pd.DataFrame,
    reference_date: str | None = None,
) -> pd.DataFrame:
    """Costruisce la mailing list una-riga-per-email arricchita con metriche.

    Args:
        clients: anagrafica caricata da loaders.load_clients.
        sales: vendite caricate da loaders.load_sales.
        orders: ordini caricati da loaders.load_orders.

    Returns:
        DataFrame con una riga per ogni email (e una riga per cliente senza
        email, segnalata da email vuota e flag email_valida=False).
    """
    aggregated = aggregate_clients(clients)
    metrics = compute_client_metrics(sales, orders, reference_date)

    # Fallback ragione sociale dalle vendite per clienti non presenti in anagrafica.
    rag_sales = (
        sales.dropna(subset=["RAGIONE SOCIALE 1"])
        .groupby("CLIENTE")["RAGIONE SOCIALE 1"]
        .first()
        .reset_index()
        .rename(columns={"CLIENTE": "codice_cliente", "RAGIONE SOCIALE 1": "ragione_sociale_sales"})
    )

    merged = aggregated.merge(metrics, on="codice_cliente", how="outer")
    merged = merged.merge(rag_sales, on="codice_cliente", how="left")
    merged["ragione_sociale"] = (
        merged["ragione_sociale"].fillna("").astype("string")
        .where(lambda s: s.str.len() > 0, merged["ragione_sociale_sales"].fillna(""))
    )
    merged["ragione_sociale"] = merged["ragione_sociale"].fillna("")
    merged.drop(columns=["ragione_sociale_sales"], inplace=True)
    merged["telefoni"] = merged["telefoni"].fillna("")
    merged["fatturato_totale"] = merged["fatturato_totale"].fillna(0.0)
    merged["fatturato_12m"] = merged["fatturato_12m"].fillna(0.0)
    merged["backlog_attuale"] = merged["backlog_attuale"].fillna(0.0)
    merged["n_ordini_aperti"] = merged["n_ordini_aperti"].fillna(0).astype(int)
    for flag in ("ha_acquistato_storico", "ha_acquistato_12m", "ha_ordine_aperto"):
        # Sostituiamo NaN con False senza l'auto-downcasting deprecato di pandas.
        merged[flag] = merged[flag].where(merged[flag].notna(), False).astype(bool)
    if "provincia" in merged.columns:
        merged["provincia"] = merged["provincia"].astype("string").fillna("")
    else:
        merged["provincia"] = ""

    rows = []
    for _, r in merged.iterrows():
        emails = r["emails"] if isinstance(r["emails"], list) else []
        prov = r.get("provincia", "")
        if prov is None or (isinstance(prov, float) and pd.isna(prov)) or pd.isna(prov):
            prov = ""
        base = {
            "codice_cliente": int(r["codice_cliente"]),
            "ragione_sociale": r["ragione_sociale"],
            "provincia": str(prov),
            "telefoni": r["telefoni"],
            "fatturato_totale": float(r["fatturato_totale"]),
            "fatturato_12m": float(r["fatturato_12m"]),
            "data_ultimo_acquisto": r["data_ultimo_acquisto"],
            "backlog_attuale": float(r["backlog_attuale"]),
            "n_ordini_aperti": int(r["n_ordini_aperti"]),
            "ha_acquistato_storico": bool(r["ha_acquistato_storico"]),
            "ha_acquistato_12m": bool(r["ha_acquistato_12m"]),
            "ha_ordine_aperto": bool(r["ha_ordine_aperto"]),
        }
        if not emails:
            rows.append({**base, "email": "", "email_valida": False})
        else:
            for em in emails:
                em_clean = "" if em is None else str(em)
                rows.append({**base, "email": em_clean, "email_valida": _is_valid_email(em_clean)})

    df = pd.DataFrame(rows)
    df = df.sort_values(["fatturato_totale", "ragione_sociale"], ascending=[False, True]).reset_index(drop=True)
    return df


def split_into_segments(mailing: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Suddivide la mailing list nei segmenti richiesti."""
    has_email = mailing["email"].astype(str).str.len() > 0
    return {
        "tutti": mailing,
        "solo_acquirenti": mailing[mailing["ha_acquistato_storico"] & ~mailing["ha_ordine_aperto"] & has_email],
        "solo_ordini_aperti": mailing[~mailing["ha_acquistato_storico"] & mailing["ha_ordine_aperto"] & has_email],
        "acquirenti_e_ordini": mailing[mailing["ha_acquistato_storico"] & mailing["ha_ordine_aperto"] & has_email],
        "senza_email": mailing[~has_email & (mailing["ha_acquistato_storico"] | mailing["ha_ordine_aperto"])],
    }


def build_all(
    clients: pd.DataFrame,
    sales: pd.DataFrame,
    orders: pd.DataFrame,
    output_dir: Path | None = None,
    reference_date: str | None = None,
) -> dict[str, pd.DataFrame]:
    """Costruisce la mailing list e scrive CSV ed Excel di output.

    Args:
        clients, sales, orders: dataset normalizzati.
        output_dir: directory di output (default `output/mailing/`).
        reference_date: data di riferimento.

    Returns:
        Dict con `mailing` (df completo) e `segments` (dict per export).
    """
    output_dir = output_dir or (OUTPUT_DIR / "mailing")
    output_dir.mkdir(parents=True, exist_ok=True)

    mailing = build_mailing_list(clients, sales, orders, reference_date)

    has_email = mailing["email"].astype(str).str.len() > 0

    acquirenti = mailing[mailing["ha_acquistato_storico"] & has_email]
    ordini_aperti = mailing[mailing["ha_ordine_aperto"] & has_email]

    acquirenti.to_csv(output_dir / "mailing_list_acquirenti.csv", index=False)
    ordini_aperti.to_csv(output_dir / "mailing_list_ordini_aperti.csv", index=False)

    segments = split_into_segments(mailing)
    excel_path = output_dir / "mailing_list_completa.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for sheet, df in {
            "Tutti": segments["tutti"],
            "Solo acquirenti": segments["solo_acquirenti"],
            "Solo ordini aperti": segments["solo_ordini_aperti"],
            "Acquirenti E ordini": segments["acquirenti_e_ordini"],
            "Senza email": segments["senza_email"],
        }.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)

    logger.info(
        "Mailing list: %d righe totali, %d acquirenti, %d ordini aperti, %d senza email",
        len(mailing), len(acquirenti), len(ordini_aperti), len(segments["senza_email"]),
    )
    return {"mailing": mailing, "segments": segments}
