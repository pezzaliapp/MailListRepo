# MailListRepo

Applicazione Python per **analizzare lo storico vendite, gli ordini aperti e
generare la mailing list dei clienti** a partire dagli export del gestionale
(file Excel in italiano).

I dati di input contengono PII e **non sono versionati**: la pipeline va
eseguita in locale con i tre file `.xlsx` posizionati in `data/`.

## Cosa fa

1. **Analisi vendite** (`src/analytics/sales.py`): fatturato per anno,
   anno × macro-famiglia, top clienti (storico e ultimi 12 mesi), per
   geografia, per agente, top articoli per fatturato e quantità, trend
   mensile 24 mesi con confronto YoY, tasso resi, clienti dormienti,
   frequenza acquisti e ticket medio.
2. **Analisi ordini aperti** (`src/analytics/orders.py`): backlog totale,
   per cliente, per macro-famiglia, per geografia; distribuzione per fascia
   di consegna ed età ordine; top articoli ancora da evadere.
3. **Mailing list** (`src/mailing/builder.py`): incrocio anagrafica + vendite
   + ordini, con segmentazione (solo acquirenti / solo ordini aperti /
   entrambi / senza email) e validazione dell'indirizzo email.
4. **Dashboard interattiva** (`app.py`): Streamlit + Plotly, con filtri per
   anno / agente / regione / macro-famiglia, ricerca cliente e download
   CSV/Excel della mailing list.

## Prerequisiti

- Python 3.11+
- I tre file Excel posizionati in `data/`:
  - `data/elenco_clienti_al_30_aprile_2026.xlsx`
  - `data/ordini_al_30_aprile_2026.xlsx`
  - `data/vendite_al_30_aprile_2026.xlsx`

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Esecuzione

### Generazione di tutti gli output

```bash
python run_pipeline.py
```

Produce in `output/`:

- `analisi_vendite.xlsx` (16 fogli)
- `analisi_ordini.xlsx` (8 fogli)
- `mailing/mailing_list_acquirenti.csv`
- `mailing/mailing_list_ordini_aperti.csv`
- `mailing/mailing_list_completa.xlsx` (5 fogli: Tutti, Solo acquirenti,
  Solo ordini aperti, Acquirenti E ordini, Senza email)
- `output/vendite_csv/*.csv` e `output/ordini_csv/*.csv` per import esterni.

A fine esecuzione lo script stampa:

- numero clienti totali in anagrafica
- clienti con almeno un acquisto storico
- clienti con ordini aperti

### Dashboard

```bash
streamlit run app.py
```

Si apre su `http://localhost:8501`. Le tab disponibili:

| Tab            | Cosa offre                                                                 |
| -------------- | -------------------------------------------------------------------------- |
| Vendite        | Tutti i grafici/tabelle del modulo analisi vendite, con filtri laterali.   |
| Ordini aperti  | KPI di backlog, scadenze, età, top articoli aperti.                        |
| Mailing list   | Tabella filtrabile per segmento + ricerca + bottoni di export CSV / Excel. |
| Anagrafica     | Ricerca cliente per ragione sociale o codice e drill-down su acquisti/ordini. |

### Test

```bash
pytest
```

I test usano fixture sintetiche (email finte tipo `cliente@example.com`) e
non richiedono i file reali, ad eccezione di `tests/test_loaders.py` che
viene saltato automaticamente se i file non sono presenti.

## Struttura del progetto

```
MailListRepo/
├── data/                              # input gitignored (PII)
├── output/                            # output gitignored
├── src/
│   ├── __init__.py                    # path/costanti centralizzati
│   ├── loaders.py                     # caricamento e pulizia dei 3 xlsx
│   ├── analytics/
│   │   ├── sales.py                   # 10+ analisi vendite
│   │   └── orders.py                  # analisi ordini aperti
│   └── mailing/
│       └── builder.py                 # mailing list e segmenti
├── tests/                             # pytest, fixture sintetiche
├── app.py                             # dashboard Streamlit
├── run_pipeline.py                    # CLI: rigenera tutti gli output
├── requirements.txt
├── .gitignore
├── LICENSE
└── README.md
```

## Convenzioni tecniche

- Tutti i path partono da `BASE_DIR` definito in `src/__init__.py`.
- Logging via modulo `logging`, livello pilotabile da env var `LOG_LEVEL`.
- `CLIENTE` e `CONTATTO CLIENTI` sono normalizzati a `int64` su tutti i
  dataset prima del join.
- I resi vengono identificati combinando `TIPO SPEDIZIONE` (prefisso `R`)
  e segno di `IMPORTO CONSEGNATO`; il fatturato netto è in `IMPORTO_NETTO`.
- Email validate con regex semplice; le email non valide vengono
  segnalate (flag `email_valida=False`) ma non rimosse.

## Licenza

MIT — vedi [LICENSE](LICENSE).
