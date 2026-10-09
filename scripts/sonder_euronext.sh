#!/bin/bash
# Sonde l'export CSV d'Euronext (diagnostic) : ajustement des dividendes, requête minimale.
set -u
B="https://live.euronext.com/en/ajax/AwlHistoricalPrice/getFullDownloadAjax"
P="format=csv&decimal_separator=.&date_form=d/m/Y&op=&base100=&startdate=2025-01-01&enddate=2025-12-31"
mkdir -p sonde
for adj in Y N; do
  curl -sS -m 60 -o sonde/MC_$adj.csv -w "MC adjusted=$adj : HTTP %{http_code} %{size_download} octets\n" "$B/FR0000121014-XPAR?$P&adjusted=$adj"
  echo "lignes: $(wc -l < sonde/MC_$adj.csv)"; grep -E "^2[2-8]/04/2025" sonde/MC_$adj.csv; tail -2 sonde/MC_$adj.csv
done
echo "### Requete minimale (format + dates)"
curl -sS -m 60 -o sonde/min.csv -w "HTTP %{http_code} %{size_download}\n" "$B/FR0000121014-XPAR?format=csv&startdate=2025-01-01&enddate=2025-12-31"
head -6 sonde/min.csv; wc -l < sonde/min.csv
echo "### Les 5 titres (adjusted=N) : nb de lignes, premiere et derniere date"
for isin in FR0000121014 FR0000120271 FR0000120578 FR0000131104 FR0000120073; do
  curl -sS -m 60 -o sonde/$isin.csv "$B/$isin-XPAR?$P&adjusted=N"
  echo "$isin $(wc -l < sonde/$isin.csv) $(sed -n 5p sonde/$isin.csv | cut -d';' -f1,6) ... $(tail -1 sonde/$isin.csv | cut -d';' -f1,6)"
done
echo "### TTE autour du 26/03/2025 (dividende 0.79) Y vs N"
curl -sS -m 60 "$B/FR0000120271-XPAR?$P&adjusted=Y" | grep -E "^2[4-8]/03/2025"
echo ---; grep -E "^2[4-8]/03/2025" sonde/FR0000120271.csv
echo "### Periode de plus de 2 ans (2023)"
curl -sS -m 60 "$B/FR0000121014-XPAR?format=csv&decimal_separator=.&date_form=d/m/Y&startdate=2023-01-01&enddate=2023-01-31&adjusted=N" | head -6
exit 0
