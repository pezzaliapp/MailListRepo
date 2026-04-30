// MailListRepo — elaborazione client-side dei 3 file Excel.
// Tutto in browser, nessun upload remoto.

const SHEETS = {
  anagrafica: "RUBRICA CONTATTI",
  ordini: "ORDINI CLIENTI DETTAGLIO",
  vendite: "SPEDIZIONI E RESI CLIENTI DETTA",
};

const COLUMNS = {
  anagrafica: ["CLIENTE", "RAGIONE SOCIALE 1", "EMAIL"],
  ordini: [
    "CLIENTE", "CLIENTE.1", "NUM.", "DATA CREAZIONE",
    "ARTICOLO", "DESCRIZIONE", "QTA INEVASA", "IMPORTO INEVASO",
  ],
  vendite: [
    "CLIENTE", "RAGIONE SOCIALE 1", "DATA SPEDIZIONE",
    "ARTICOLO", "DESCRIZIONE", "QTA CONSEGNATA", "IMPORTO CONSEGNATO",
    "TIPO SPEDIZIONE",
  ],
};

const state = {
  anagrafica: null,
  ordini: null,
  vendite: null,
  result: null,
};

// ---------- Utilità ----------

const fmtEuro = (n) =>
  (Number(n) || 0).toLocaleString("it-IT", {
    style: "currency",
    currency: "EUR",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

const fmtInt = (n) => (Number(n) || 0).toLocaleString("it-IT");

const fmtDate = (d) => {
  if (!d) return "";
  const date = d instanceof Date ? d : new Date(d);
  if (isNaN(date.getTime())) return "";
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
};

const toNumber = (v) => {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const s = String(v).replace(/\./g, "").replace(",", ".").replace(/[^\d\-.]/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
};

const toDate = (v) => {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === "number") {
    // Excel serial date fallback
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(epoch.getTime() + v * 86400000);
    return isNaN(d.getTime()) ? null : d;
  }
  const s = String(v).trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) {
    let [, d, mo, y] = m;
    if (y.length === 2) y = "20" + y;
    const dt = new Date(Number(y), Number(mo) - 1, Number(d));
    return isNaN(dt.getTime()) ? null : dt;
  }
  const dt = new Date(s);
  return isNaN(dt.getTime()) ? null : dt;
};

const isValidEmail = (s) => {
  if (!s) return false;
  const t = String(s).trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(t);
};

const findSheet = (workbook, wanted) => {
  if (workbook.SheetNames.includes(wanted)) return workbook.Sheets[wanted];
  const norm = (s) => s.trim().toUpperCase();
  const target = norm(wanted);
  for (const name of workbook.SheetNames) {
    if (norm(name) === target) return workbook.Sheets[name];
  }
  return null;
};

const checkColumns = (rows, required, label) => {
  if (!rows.length) throw new Error(`${label}: nessuna riga trovata.`);
  const have = new Set(Object.keys(rows[0]));
  const missing = required.filter((c) => !have.has(c));
  if (missing.length) {
    throw new Error(`${label}: colonne mancanti → ${missing.join(", ")}`);
  }
};

const readWorkbook = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), {
          type: "array",
          cellDates: true,
        });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Errore lettura file " + file.name));
    reader.readAsArrayBuffer(file);
  });

const sheetRows = (sheet) =>
  XLSX.utils.sheet_to_json(sheet, { defval: null, raw: false });

// raw:false converts numbers/dates already; we'll re-parse with our helpers anyway.
// But dates from cellDates:true come as ISO strings if raw:false — switch to raw:true for dates.
// Safer: read with two passes — use raw:true for numeric/date parsing.
const sheetRowsRaw = (sheet) =>
  XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });

// ---------- Parsing per file ----------

function parseAnagrafica(workbook) {
  const sheet = findSheet(workbook, SHEETS.anagrafica);
  if (!sheet) throw new Error(`Anagrafica: foglio "${SHEETS.anagrafica}" non trovato.`);
  const rows = sheetRowsRaw(sheet);
  console.log("Anagrafica:", rows.length, "righe, prima riga:", rows[0]);
  checkColumns(rows, COLUMNS.anagrafica, "Anagrafica");

  // Una riga = un cliente. Email singola (può essere vuota).
  const map = new Map(); // code -> { ragione, email }
  for (const r of rows) {
    const codeRaw = r["CLIENTE"];
    if (codeRaw === null || codeRaw === undefined || codeRaw === "") continue;
    const code = String(codeRaw).trim();
    if (!code) continue;
    const ragione = r["RAGIONE SOCIALE 1"] ? String(r["RAGIONE SOCIALE 1"]).trim() : "";
    const emailRaw = r["EMAIL"] ? String(r["EMAIL"]).trim().toLowerCase() : "";
    const email = isValidEmail(emailRaw) ? emailRaw : "";
    map.set(code, { ragione, email });
  }
  return map;
}

function parseVendite(workbook) {
  const sheet = findSheet(workbook, SHEETS.vendite);
  if (!sheet) throw new Error(`Vendite: foglio "${SHEETS.vendite}" non trovato.`);
  const rows = sheetRowsRaw(sheet);
  checkColumns(rows, COLUMNS.vendite, "Vendite");

  const agg = new Map(); // code -> { ragione, count, revenue, lastDate }
  for (const r of rows) {
    const codeRaw = r["CLIENTE"];
    if (codeRaw === null || codeRaw === undefined || codeRaw === "") continue;
    const code = String(codeRaw).trim();
    if (!code) continue;

    const tipo = r["TIPO SPEDIZIONE"];
    const isReso = tipo && String(tipo).trim().toUpperCase().startsWith("R");
    const importo = toNumber(r["IMPORTO CONSEGNATO"]);
    const signed = isReso ? -importo : importo;
    const date = toDate(r["DATA SPEDIZIONE"]);

    if (!agg.has(code)) {
      agg.set(code, {
        ragione: r["RAGIONE SOCIALE 1"] ? String(r["RAGIONE SOCIALE 1"]).trim() : "",
        count: 0,
        revenue: 0,
        lastDate: null,
      });
    }
    const e = agg.get(code);
    if (!e.ragione && r["RAGIONE SOCIALE 1"]) e.ragione = String(r["RAGIONE SOCIALE 1"]).trim();
    e.count += 1;
    e.revenue += signed;
    if (date && (!e.lastDate || date > e.lastDate)) e.lastDate = date;
  }
  return agg;
}

function parseOrdini(workbook) {
  const sheet = findSheet(workbook, SHEETS.ordini);
  if (!sheet) throw new Error(`Ordini: foglio "${SHEETS.ordini}" non trovato.`);
  const rows = sheetRowsRaw(sheet);
  checkColumns(rows, COLUMNS.ordini, "Ordini");

  const agg = new Map(); // code -> { ragione, orders: Set, backlog, articoli: [] }
  for (const r of rows) {
    const qta = toNumber(r["QTA INEVASA"]);
    const imp = toNumber(r["IMPORTO INEVASO"]);
    if (qta <= 0 && imp <= 0) continue;

    const codeRaw = r["CLIENTE"];
    if (codeRaw === null || codeRaw === undefined || codeRaw === "") continue;
    const code = String(codeRaw).trim();
    if (!code) continue;

    if (!agg.has(code)) {
      agg.set(code, {
        ragione: r["CLIENTE.1"] ? String(r["CLIENTE.1"]).trim() : "",
        orders: new Set(),
        backlog: 0,
        articoli: [],
      });
    }
    const e = agg.get(code);
    if (!e.ragione && r["CLIENTE.1"]) e.ragione = String(r["CLIENTE.1"]).trim();
    if (r["NUM."] !== null && r["NUM."] !== undefined && r["NUM."] !== "") {
      e.orders.add(String(r["NUM."]).trim());
    }
    e.backlog += imp;
    const art = r["ARTICOLO"] ? String(r["ARTICOLO"]).trim() : "";
    const desc = r["DESCRIZIONE"] ? String(r["DESCRIZIONE"]).trim() : "";
    const label = [art, desc].filter(Boolean).join(" — ");
    if (label && !e.articoli.includes(label)) e.articoli.push(label);
  }
  return agg;
}

// ---------- Composizione mail list ----------

function buildMailList(anagrafica, vendite, ordini) {
  const out = [];
  for (const [code, info] of anagrafica) {
    const hasSales = vendite.has(code);
    const hasOpenOrders = ordini.has(code);
    if (!hasSales && !hasOpenOrders) continue;
    if (!info.email) continue;
    let ragione = info.ragione;
    if (!ragione && hasSales) ragione = vendite.get(code).ragione;
    if (!ragione && hasOpenOrders) ragione = ordini.get(code).ragione;
    out.push({
      email: info.email,
      code,
      ragione: ragione || "",
      hasSales,
      hasOpenOrders,
    });
  }
  out.sort((a, b) => a.email.localeCompare(b.email));
  return out;
}

// ---------- Render ----------

const $ = (id) => document.getElementById(id);

function renderTable(target, headers, rows) {
  const t = $(target);
  if (!rows.length) {
    t.innerHTML = `<tr><td class="empty">Nessun dato.</td></tr>`;
    return;
  }
  const thead =
    "<thead><tr>" +
    headers.map((h) => `<th class="${h.num ? "num" : ""}">${escapeHtml(h.label)}</th>`).join("") +
    "</tr></thead>";
  const tbody =
    "<tbody>" +
    rows
      .map(
        (r) =>
          "<tr>" +
          headers
            .map((h) => `<td class="${h.num ? "num" : ""}">${escapeHtml(r[h.key] ?? "")}</td>`)
            .join("") +
          "</tr>"
      )
      .join("") +
    "</tbody>";
  t.innerHTML = thead + tbody;
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderVendite(vendite) {
  const arr = [];
  let totaleFatturato = 0;
  for (const [code, e] of vendite) {
    arr.push({
      code,
      ragione: e.ragione,
      count: e.count,
      revenue: e.revenue,
      lastDate: e.lastDate,
    });
    totaleFatturato += e.revenue;
  }
  arr.sort((a, b) => b.revenue - a.revenue);

  $("summaryVendite").innerHTML =
    `<span><strong>${fmtInt(arr.length)}</strong> clienti con acquisti</span>` +
    `<span>fatturato totale storico: <strong>${fmtEuro(totaleFatturato)}</strong></span>`;

  renderTable(
    "tableVendite",
    [
      { key: "code", label: "Codice", num: false },
      { key: "ragione", label: "Ragione sociale", num: false },
      { key: "countF", label: "Spedizioni", num: true },
      { key: "revenueF", label: "Fatturato", num: true },
      { key: "lastDateF", label: "Ultimo acquisto", num: false },
    ],
    arr.map((r) => ({
      ...r,
      countF: fmtInt(r.count),
      revenueF: fmtEuro(r.revenue),
      lastDateF: fmtDate(r.lastDate),
    }))
  );
  $("cardVendite").style.display = "";
}

function renderOrdini(ordini) {
  const arr = [];
  let backlogTot = 0;
  for (const [code, e] of ordini) {
    arr.push({
      code,
      ragione: e.ragione,
      orders: e.orders.size,
      backlog: e.backlog,
      articoli: e.articoli.slice(0, 3).join(" · ") + (e.articoli.length > 3 ? " …" : ""),
    });
    backlogTot += e.backlog;
  }
  arr.sort((a, b) => b.backlog - a.backlog);

  $("summaryOrdini").innerHTML =
    `<span><strong>${fmtInt(arr.length)}</strong> clienti con ordini aperti</span>` +
    `<span>backlog totale: <strong>${fmtEuro(backlogTot)}</strong></span>`;

  renderTable(
    "tableOrdini",
    [
      { key: "code", label: "Codice", num: false },
      { key: "ragione", label: "Ragione sociale", num: false },
      { key: "ordersF", label: "Ordini aperti", num: true },
      { key: "backlogF", label: "Backlog", num: true },
      { key: "articoli", label: "Articoli (max 3)", num: false },
    ],
    arr.map((r) => ({
      ...r,
      ordersF: fmtInt(r.orders),
      backlogF: fmtEuro(r.backlog),
    }))
  );
  $("cardOrdini").style.display = "";
}

function renderMaillist(list) {
  $("summaryMaillist").innerHTML =
    `<span><strong>${fmtInt(list.length)}</strong> indirizzi email</span>`;
  renderTable(
    "tableMaillist",
    [
      { key: "email", label: "Email", num: false },
      { key: "code", label: "Codice", num: false },
      { key: "ragione", label: "Ragione sociale", num: false },
      { key: "salesF", label: "Ha acquistato?", num: false },
      { key: "ordersF", label: "Ordini aperti?", num: false },
    ],
    list.map((r) => ({
      ...r,
      salesF: r.hasSales ? "Sì" : "No",
      ordersF: r.hasOpenOrders ? "Sì" : "No",
    }))
  );
  $("cardMaillist").style.display = "";
}

// ---------- CSV ----------

function toCsv(list) {
  const headers = ["Email", "Codice cliente", "Ragione sociale", "Ha acquistato?", "Ha ordini aperti?"];
  const escape = (v) => {
    const s = v === null || v === undefined ? "" : String(v);
    if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  };
  const lines = [headers.map(escape).join(",")];
  for (const r of list) {
    lines.push(
      [
        r.email,
        r.code,
        r.ragione,
        r.hasSales ? "Sì" : "No",
        r.hasOpenOrders ? "Sì" : "No",
      ].map(escape).join(",")
    );
  }
  return "﻿" + lines.join("\r\n");
}

function downloadCsv(list) {
  const blob = new Blob([toCsv(list)], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "maillist.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ---------- UI wiring ----------

function setStatus(id, msg, ok) {
  const el = $(id);
  el.textContent = msg;
  el.className = "status " + (msg ? (ok ? "ok" : "err") : "");
}

function showError(msg) {
  const box = $("errorBox");
  if (!msg) {
    box.style.display = "none";
    box.textContent = "";
    return;
  }
  // Must be "block" (not ""): the CSS rule sets display:none, so clearing
  // the inline style would let the element fall back to hidden.
  box.style.display = "block";
  box.textContent = msg;
}

function bindFile(inputId, statusId, key, label) {
  $(inputId).addEventListener("change", async (ev) => {
    const f = ev.target.files[0];
    if (!f) {
      state[key] = null;
      setStatus(statusId, "", true);
      return;
    }
    setStatus(statusId, "Lettura in corso…", true);
    try {
      const wb = await readWorkbook(f);
      state[key] = wb;
      setStatus(statusId, `Caricato: ${f.name}`, true);
    } catch (err) {
      state[key] = null;
      setStatus(statusId, `Errore: ${err.message || err}`, false);
    }
  });
}

function reset() {
  state.anagrafica = null;
  state.ordini = null;
  state.vendite = null;
  state.result = null;
  ["fileAnagrafica", "fileOrdini", "fileVendite"].forEach((id) => ($(id).value = ""));
  ["statusAnagrafica", "statusOrdini", "statusVendite"].forEach((id) => setStatus(id, "", true));
  ["cardVendite", "cardOrdini", "cardMaillist"].forEach((id) => ($(id).style.display = "none"));
  showError("");
}

function process() {
  console.log("Elabora cliccato");
  try {
    console.log("File caricati:", {
      hasAnagrafica: !!state.anagrafica,
      hasOrdini: !!state.ordini,
      hasVendite: !!state.vendite,
    });

    const missing = [];
    if (!state.anagrafica) missing.push("Anagrafica");
    if (!state.ordini) missing.push("Ordini");
    if (!state.vendite) missing.push("Vendite");
    if (missing.length) {
      showError("Carica tutti e 3 i file prima di elaborare. Mancano: " + missing.join(", "));
      return;
    }

    showError("");

    console.log("Sheets nei workbook:", {
      anagrafica: state.anagrafica.SheetNames,
      ordini: state.ordini.SheetNames,
      vendite: state.vendite.SheetNames,
    });

    const anagrafica = parseAnagrafica(state.anagrafica);
    console.log("Anagrafica parsata. Clienti unici:", anagrafica.size);

    const vendite = parseVendite(state.vendite);
    console.log("Vendite parsate. Clienti con vendite:", vendite.size);

    const ordini = parseOrdini(state.ordini);
    console.log("Ordini parsati. Clienti con ordini aperti:", ordini.size);

    const maillist = buildMailList(anagrafica, vendite, ordini);
    console.log("Dopo aggregazione:", {
      clientiAnagrafica: anagrafica.size,
      clientiVendite: vendite.size,
      clientiOrdini: ordini.size,
      mailListRighe: maillist.length,
    });

    state.result = { anagrafica, vendite, ordini, maillist };

    renderVendite(vendite);
    renderOrdini(ordini);
    renderMaillist(maillist);

    $("cardVendite").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (err) {
    console.error("Errore in process:", err);
    showError(err && err.message ? err.message : String(err));
    ["cardVendite", "cardOrdini", "cardMaillist"].forEach((id) => ($(id).style.display = "none"));
  }
}

document.addEventListener("DOMContentLoaded", () => {
  bindFile("fileAnagrafica", "statusAnagrafica", "anagrafica", "Anagrafica");
  bindFile("fileOrdini", "statusOrdini", "ordini", "Ordini");
  bindFile("fileVendite", "statusVendite", "vendite", "Vendite");
  $("btnProcess").addEventListener("click", process);
  $("btnReset").addEventListener("click", reset);
  $("btnDownloadCsv").addEventListener("click", () => {
    if (state.result && state.result.maillist) downloadCsv(state.result.maillist);
  });
});
