#!/bin/bash
# Télécharge l'export CSV Euronext (2025) des 5 actions dans donnees/euronext_2025/ (jeu de test de la macro).
set -eu
B="https://live.euronext.com/en/ajax/AwlHistoricalPrice/getFullDownloadAjax"
mkdir -p donnees/euronext_2025
for paire in MC:FR0000121014 TTE:FR0000120271 SAN:FR0000120578 BNP:FR0000131104 AI:FR0000120073; do
  code=${paire%%:*}; isin=${paire##*:}
  curl -sS -m 60 -o "donnees/euronext_2025/$code.csv" "$B/$isin-XPAR?format=csv&decimal_separator=.&date_form=d/m/Y&adjusted=N&startdate=2025-01-01&enddate=2025-12-31"
  echo "$code : $(wc -l < donnees/euronext_2025/$code.csv) lignes"
done
