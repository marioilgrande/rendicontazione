#!/bin/bash
# Avvia il server di invio mail rendicontazione.
# Doppio click su questo file dal Desktop.

cd "/Users/marioisernia/Desktop/HTML MARIO/RENDICONTAZIONE/server" || {
  echo ""
  echo "❌ Cartella server non trovata."
  echo "   Atteso: /Users/marioisernia/Desktop/HTML MARIO/RENDICONTAZIONE/server"
  echo ""
  echo "Premi Invio per chiudere…"
  read
  exit 1
}

clear
echo "======================================================"
echo "  RENDICONTAZIONE · Server invio mail"
echo "======================================================"
echo ""
echo "  Cartella: $(pwd)"
echo "  Avvio in corso…"
echo ""
echo "  Quando vedi 'in ascolto su http://localhost:5170',"
echo "  apri index.html (o https://rendicontazione.vercel.app) e procedi."
echo ""
echo "  Per FERMARE il server: premi Ctrl+C oppure chiudi"
echo "  questa finestra."
echo "------------------------------------------------------"
echo ""

npm start

echo ""
echo "Server fermato. Premi Invio per chiudere…"
read
