# Rendicontazione Contratti Acea

Pannello HTML + server locale per gestire la rendicontazione mensile dei contratti Acea e inviare in massa le mail ai negozi con allegato Excel personalizzato.

## Struttura

```
RENDICONTAZIONE/
├── rendinew.html              ← pannello principale (importa Excel, filtra, dashboard, export, invio)
├── PROCEDURA_SERVER.png       ← infografica con la procedura di avvio del server
├── PROCEDURA_SERVER.svg       ← versione vettoriale dell'infografica
├── backup maggio26.html       ← snapshot precedente del pannello
├── OLD/                       ← versioni vecchie
└── server/                    ← server Node.js per invio mail massivo
    ├── server.js
    ├── package.json
    ├── .env.example           ← template configurazione SMTP (il .env reale NON è in repo)
    └── README.md              ← istruzioni dettagliate server
```

## Funzionalità principali (pannello HTML)

- Import file Excel mensile (legge automaticamente il foglio `DB Acea` in qualsiasi variante).
- Lettura compenso da `COMPENSO CORRETTO` (fallback a `COMPENSO`). Valore `-20` → trattato come `0`.
- Filtri: AM, negozi singoli (chip), stati (OK, OPZIONI, KO, NON FIRMATO).
- Dashboard per negozio con conteggio e fatturato OK + numero e importo opzioni.
- Export Excel filtrato o per singolo negozio con stile.
- Rubrica email negozi (salvata in `localStorage` del browser).
- Indicatore stato server locale (online/offline) in alto a destra.
- Modale di invio massivo con:
  - anteprima per dealer (email/OK/opzioni)
  - selezione manuale via checkbox per ogni negozio
  - log live, barra di progresso, interrompi a metà
- Testo email automatico che include numero e importo opzioni (se presenti).

## Server di invio (cartella `server/`)

Server Node.js + Express + Nodemailer che riceve dal pannello l'elenco invii, genera per ogni negozio l'xlsx personalizzato e lo invia via SMTP.

Vedi [`server/README.md`](server/README.md) per setup e avvio.

## Avvio rapido

1. Doppio click su `AVVIA SERVER.command` (sul Desktop) → il server si avvia.
2. Apri `rendinew.html` nel browser.
3. Importa il file Excel del mese, filtra, dashboard.
4. Per l'invio massivo: clicca "📧 Invia mail a tutti i negozi", controlla l'anteprima, conferma.

## Note

- L'invio massivo richiede il server locale in esecuzione.
- Il file `server/.env` (credenziali SMTP) **non è versionato**. Copia `server/.env.example` come `server/.env` e configura.
- Il pannello funziona anche offline (solo file HTML), ma in tal caso non si può fare invio massivo.
