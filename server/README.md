# Mailer Rendicontazione – istruzioni

Server locale per inviare in massa le rendicontazioni ai negozi con allegato Excel personalizzato, usando il tuo SMTP Register.it.

## 1. Setup iniziale (una sola volta)

Apri il Terminale e vai nella cartella del server:

```bash
cd "/Users/marioisernia/Desktop/HTML MARIO/RENDICONTAZIONE/server"
```

Installa le dipendenze:

```bash
npm install
```

Copia il file `.env.example` come `.env`:

```bash
cp .env.example .env
```

Apri `.env` e compila i campi:

```
SMTP_HOST=authsmtp.register.it
SMTP_PORT=465
SMTP_SECURE=true
SMTP_USER=tuamail@tuodominio.it
SMTP_PASS=la_tua_password_casella
SMTP_FROM_NAME=Mario Isernia
SMTP_FROM_EMAIL=tuamail@tuodominio.it
```

**Nota Register.it:** se `authsmtp.register.it:465` non funziona, prova `smtps.register.it:465` o il server SMTP del tuo dominio (lo trovi nel pannello Register).

## 2. Avvio (ogni volta che vuoi inviare)

```bash
cd "/Users/marioisernia/Desktop/HTML MARIO/RENDICONTAZIONE/server"
npm start
```

Vedrai:

```
Rendicontazione mailer in ascolto su http://localhost:5170
```

Lascia il Terminale aperto.

## 3. Uso

1. Apri `index.html` (o https://rendicontazione.vercel.app).
2. Importa il file Excel del mese.
3. Seleziona AM (o tutti).
4. Clicca **"Invia mail a tutti i negozi"**.
5. Si apre un riepilogo: vedi negozi, email, n. pratiche, totali.
6. Conferma → il server genera gli xlsx e li manda uno alla volta.
7. Vedi il log in tempo reale: ✅ inviati, ❌ errori, ⚠️ saltati.

## 4. Fermare il server

Nel Terminale premi `Ctrl+C`.

## 5. Sicurezza

- Il file `.env` contiene la tua password mail. **Non condividerlo**.
- Il server gira solo in locale (`localhost`), nessuno da fuori può accedere.
- Nessun dato viene salvato: ogni invio è in memoria e poi scartato.
