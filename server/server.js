import "dotenv/config";
import express from "express";
import cors from "cors";
import nodemailer from "nodemailer";
import XLSX from "xlsx-js-style";

const PORT = Number(process.env.PORT || 5170);
const SEND_DELAY_MS = Number(process.env.SEND_DELAY_MS || 600);
const BRAND_LABEL = "Acea";

/* Fallback: le colonne vere arrivano dal pannello con le righe
   (vedi colonneDa()), così client e server non divergono mai. */
const DESIRED_COLUMNS = [
  "DEALER","AM","STATO","Runner","Nome opportunità","Stato PDC","Causale annullamento",
  "Data creazione","Data firma","Tipo cliente","Partita IVA","Codice Fiscale",
  "Ragione Sociale","Nome Cliente","Cognome Cliente","Billing Profile: Modalità di Pagamento",
  "Categoria d'uso","Service Point Code","Prestazione","Indirizzo","Nome Prodotto",
  "Tipo Integration Case","Stato e Causale Annullamento","Causale Ebdm","Causale Annullamento","Compenso"
];

/* ---------- SMTP transporter ---------- */
function buildTransporter(){
  const host = process.env.SMTP_HOST;
  const port = Number(process.env.SMTP_PORT || 465);
  const secure = String(process.env.SMTP_SECURE ?? "true") === "true";
  const user = process.env.SMTP_USER;
  const pass = process.env.SMTP_PASS;
  if (!host || !user || !pass){
    throw new Error("SMTP non configurato: controlla il file .env");
  }
  return nodemailer.createTransport({
    host, port, secure,
    auth: { user, pass },
    tls: { rejectUnauthorized: false }
  });
}

function fromHeader(){
  const name = process.env.SMTP_FROM_NAME || "Rendicontazione";
  const email = process.env.SMTP_FROM_EMAIL || process.env.SMTP_USER;
  return `"${name}" <${email}>`;
}

/* ---------- helpers ---------- */
function fmtEUR(n){
  const v = Number.isFinite(Number(n)) ? Number(n) : 0;
  return v.toLocaleString("it-IT", { style:"currency", currency:"EUR" });
}

function safeName(s){
  return String(s||"")
    .replace(/[^\w\-àèìòùÀÈÌÒÙ ]+/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function buildDealerFilename(dealer, monthLabel){
  const dealerSafe = safeName(dealer).slice(0, 60);
  return `rendicontazione ${BRAND_LABEL} ${monthLabel} (${dealerSafe}).xlsx`;
}

/* ---------- XLSX generation ---------- */
/* Le colonne (e il loro ordine) sono quelle che il pannello ha
   messo nelle righe: se il pannello ne aggiunge una, l'allegato la
   prende senza toccare il server. */
function colonneDa(rows){
  const k = rows && rows[0] ? Object.keys(rows[0]) : [];
  return k.length ? k : DESIRED_COLUMNS.slice();
}

function buildXlsxBuffer(rows, opzRows){
  const COLS = colonneDa(rows);
  const aoa = [COLS.slice()];

  rows.forEach(r => {
    const line = COLS.map(col => {
      const v = r?.[col];
      if (col === "Compenso") return Number.isFinite(Number(v)) ? Number(v) : 0;
      return (v === undefined || v === null) ? "" : v;
    });
    aoa.push(line);
  });

  const total = rows.reduce((acc, r) => acc + (Number.isFinite(Number(r.Compenso)) ? Number(r.Compenso) : 0), 0);
  const totalRow = Array(COLS.length).fill("");
  totalRow[0] = "TOTALE COMPENSI";
  totalRow[COLS.indexOf("Compenso")] = total;
  aoa.push(totalRow);

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  const compIdx = COLS.indexOf("Compenso");
  const lastRow = aoa.length;

  ws["!cols"] = COLS.map((_,i)=> ({ wch: (i === compIdx ? 21 : 13) }));

  const headerGrey = {
    fill: { patternType:"solid", fgColor:{ rgb:"BFBFBF" } },
    font: { bold:true, color:{ rgb:"000000" } },
    alignment: { horizontal:"center", vertical:"center", wrapText:true },
    border: {
      top:{style:"thin", color:{rgb:"A0A0A0"}},
      bottom:{style:"thin", color:{rgb:"A0A0A0"}},
      left:{style:"thin", color:{rgb:"A0A0A0"}},
      right:{style:"thin", color:{rgb:"A0A0A0"}}
    }
  };
  const compYellow = {
    fill: { patternType:"solid", fgColor:{ rgb:"FFFF00" } },
    font: { color:{ rgb:"000000" } },
    alignment: { horizontal:"right", vertical:"center" },
    border: {
      top:{style:"thin", color:{rgb:"C0C0C0"}},
      bottom:{style:"thin", color:{rgb:"C0C0C0"}},
      left:{style:"thin", color:{rgb:"C0C0C0"}},
      right:{style:"thin", color:{rgb:"C0C0C0"}}
    }
  };
  const textCell = {
    alignment: { vertical:"center", wrapText:false },
    border: {
      top:{style:"thin", color:{rgb:"E0E0E0"}},
      bottom:{style:"thin", color:{rgb:"E0E0E0"}},
      left:{style:"thin", color:{rgb:"E0E0E0"}},
      right:{style:"thin", color:{rgb:"E0E0E0"}}
    }
  };

  for (let c=0; c<COLS.length; c++){
    const addr = XLSX.utils.encode_cell({ r:0, c });
    if (!ws[addr]) continue;
    ws[addr].s = (c === compIdx) ? compYellow : headerGrey;
  }
  for (let r=2; r<=lastRow; r++){
    for (let c=0; c<COLS.length; c++){
      const addr = XLSX.utils.encode_cell({ r:r-1, c });
      const cell = ws[addr];
      if (!cell) continue;
      if (c === compIdx){
        cell.t = (typeof cell.v === "number") ? "n" : cell.t;
        cell.z = '#,##0.00';
        cell.s = compYellow;
      } else {
        cell.s = textCell;
      }
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Rendicontazione");

  /* foglio con il dettaglio delle opzioni del negozio, se presenti */
  if (Array.isArray(opzRows) && opzRows.length){
    const head = ["NUMERO CASO","CLIENTE","OGGETTO","DATA","SOTTO AGENZIA","COMPENSO"];
    const oaoa = [head];
    opzRows.forEach(o => oaoa.push([
      o?.caso ?? "", o?.cliente ?? "", o?.oggetto ?? "", o?.data ?? "", o?.sottoAgenzia ?? "",
      Number(o?.compenso) || 0
    ]));
    const tot = opzRows.reduce((a,o) => a + (Number(o?.compenso) || 0), 0);
    const totRow = Array(head.length).fill("");
    totRow[0] = "TOTALE OPZIONI";
    totRow[head.length-1] = tot;
    oaoa.push(totRow);

    const ws2 = XLSX.utils.aoa_to_sheet(oaoa);
    ws2["!cols"] = [{wch:14},{wch:28},{wch:34},{wch:12},{wch:30},{wch:14}];
    for (let c=0; c<head.length; c++){
      const a = XLSX.utils.encode_cell({ r:0, c });
      if (ws2[a]) ws2[a].s = headerGrey;
    }
    for (let r=1; r<oaoa.length; r++){
      const a = XLSX.utils.encode_cell({ r, c:head.length-1 });
      if (ws2[a]){ ws2[a].z = '#,##0.00'; ws2[a].s = compYellow; }
    }
    XLSX.utils.book_append_sheet(wb, ws2, "Opzioni");
  }

  return XLSX.write(wb, { type:"buffer", bookType:"xlsx" });
}

/* ---------- Mail body ---------- */
function buildMailBody({ dealer, okCount, okTotal, opzCount, opzTotal }){
  let body =
`Ciao,

in allegato la rendicontazione contratti del negozio:
- DEALER: ${dealer}
- Totale pratiche OK: ${okCount}
- Totale compensi: ${fmtEUR(okTotal)}`;

  if (opzCount > 0){
    body += `
- Totale opzioni: ${opzCount}
- Importo totale opzioni: ${fmtEUR(opzTotal)}`;
  }

  body += `

In allegato il dettaglio delle pratiche che ti chiedo di verificare.
La fattura verrà emessa in questi giorni tenendo conto di questa produzione. Dovesse mancare qualcosa ti chiedo di segnalarmelo oggi altrimenti verrà inserito nella prossima rendicontazione.

Grazie
`;
  return body;
}

/* ---------- App ---------- */
const app = express();
app.use(cors());
app.use(express.json({ limit: "100mb" }));

/* Health */
app.get("/api/health", (_req, res) => {
  const ok = !!(process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS);
  res.json({
    ok,
    smtp: ok ? {
      host: process.env.SMTP_HOST,
      port: Number(process.env.SMTP_PORT || 465),
      user: process.env.SMTP_USER,
      from: process.env.SMTP_FROM_EMAIL || process.env.SMTP_USER
    } : null
  });
});

/* SMTP verify */
app.get("/api/verify-smtp", async (_req, res) => {
  try{
    const t = buildTransporter();
    await t.verify();
    res.json({ ok:true });
  } catch(e){
    res.status(500).json({ ok:false, error: String(e?.message || e) });
  }
});

/* Bulk send
   Body atteso:
   {
     monthLabel: "marzo",
     dryRun: false,
     items: [
       {
         dealer: "...",
         email: "negozio@x.it",
         okCount, okTotal, opzCount, opzTotal,
         rows: [ { DEALER, AM, STATO, ... , Compenso }, ... ]
       },
       ...
     ]
   }
*/
app.post("/api/send-bulk", async (req, res) => {
  const { monthLabel = "", items = [], dryRun = false } = req.body || {};
  if (!Array.isArray(items) || !items.length){
    return res.status(400).json({ ok:false, error: "Nessun item da inviare." });
  }

  let transporter = null;
  if (!dryRun){
    try{ transporter = buildTransporter(); }
    catch(e){ return res.status(500).json({ ok:false, error: String(e?.message || e) }); }
  }

  const results = [];
  let sent = 0, skipped = 0, failed = 0;

  for (const it of items){
    const dealer = String(it?.dealer || "").trim();
    const email = String(it?.email || "").trim();
    const rows = Array.isArray(it?.rows) ? it.rows : [];

    if (!dealer || !email || !rows.length){
      skipped++;
      results.push({ dealer, email, status: "skipped", reason: "dati mancanti" });
      continue;
    }

    const subject = `Rendicontazione ${BRAND_LABEL} ${monthLabel || ""} – ${dealer}`.trim();
    const text = buildMailBody({
      dealer,
      okCount: Number(it.okCount || 0),
      okTotal: Number(it.okTotal || 0),
      opzCount: Number(it.opzCount || 0),
      opzTotal: Number(it.opzTotal || 0)
    });

    if (dryRun){
      results.push({
        dealer, email, status: "preview",
        subject,
        attachmentName: buildDealerFilename(dealer, monthLabel)
      });
      continue;
    }

    try{
      const buf = buildXlsxBuffer(rows, Array.isArray(it?.opzRows) ? it.opzRows : []);
      const filename = buildDealerFilename(dealer, monthLabel);
      await transporter.sendMail({
        from: fromHeader(),
        to: email,
        subject,
        text,
        attachments: [{ filename, content: buf }]
      });
      sent++;
      results.push({ dealer, email, status: "sent" });
    } catch(e){
      failed++;
      results.push({ dealer, email, status: "failed", reason: String(e?.message || e) });
    }

    if (SEND_DELAY_MS > 0) await new Promise(r => setTimeout(r, SEND_DELAY_MS));
  }

  res.json({ ok:true, summary: { total: items.length, sent, failed, skipped }, results });
});

/* ---------- Start ---------- */
app.listen(PORT, () => {
  console.log(`\n  Rendicontazione mailer in ascolto su http://localhost:${PORT}`);
  console.log(`  Endpoint:`);
  console.log(`    GET  /api/health`);
  console.log(`    GET  /api/verify-smtp`);
  console.log(`    POST /api/send-bulk\n`);
});
