#!/bin/bash
# Sonde les points d'accès Euronext pour l'historique des cours (diagnostic).
set -u
UA="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
ISIN="FR0000121014-XPAR"   # LVMH
mkdir -p sonde
echo "### 1. getHistoricalPricePopup adjusted=Y"
curl -sS -m 60 -A "$UA" -H "X-Requested-With: XMLHttpRequest" -o sonde/popup_Y.html -w "HTTP %{http_code} %{size_download} octets\n" \
  --data "adjusted=Y&startdate=2025-01-01&enddate=2025-12-31&nbSession=" \
  "https://live.euronext.com/en/ajax/getHistoricalPricePopup/$ISIN"
head -c 1500 sonde/popup_Y.html; echo; grep -o "<tr" sonde/popup_Y.html | wc -l
echo "### 2. getHistoricalPricePopup adjusted=N"
curl -sS -m 60 -A "$UA" -H "X-Requested-With: XMLHttpRequest" -o sonde/popup_N.html -w "HTTP %{http_code} %{size_download} octets\n" \
  --data "adjusted=N&startdate=2025-01-01&enddate=2025-12-31&nbSession=" \
  "https://live.euronext.com/en/ajax/getHistoricalPricePopup/$ISIN"
grep -o "<tr" sonde/popup_N.html | wc -l
echo "### 3. getFullDownloadAjax csv"
curl -sS -m 60 -A "$UA" -H "X-Requested-With: XMLHttpRequest" -o sonde/full.csv -w "HTTP %{http_code} %{size_download} octets\n" \
  --data "format=csv&decimal_separator=.&date_form=Y-m-d&op=&adjusted=Y&base100=&startdate=2025-01-01&enddate=2025-12-31" \
  "https://live.euronext.com/en/ajax/AwlHistoricalPrice/getFullDownloadAjax/$ISIN"
head -c 1500 sonde/full.csv; echo; wc -l sonde/full.csv
echo "### 4. getFullDownloadAjax csv adjusted=N"
curl -sS -m 60 -A "$UA" -H "X-Requested-With: XMLHttpRequest" -o sonde/full_N.csv -w "HTTP %{http_code} %{size_download} octets\n" \
  --data "format=csv&decimal_separator=.&date_form=Y-m-d&op=&adjusted=N&base100=&startdate=2025-01-01&enddate=2025-12-31" \
  "https://live.euronext.com/en/ajax/AwlHistoricalPrice/getFullDownloadAjax/$ISIN"
head -c 600 sonde/full_N.csv; echo
echo "### 5. Même requête CSV en GET"
curl -sS -m 60 -A "$UA" -o sonde/full_get.csv -w "HTTP %{http_code} %{size_download} octets\n" \
  "https://live.euronext.com/en/ajax/AwlHistoricalPrice/getFullDownloadAjax/$ISIN?format=csv&decimal_separator=.&date_form=Y-m-d&op=&adjusted=Y&base100=&startdate=2025-01-01&enddate=2025-12-31"
head -c 600 sonde/full_get.csv; echo
echo "### 6. Sans en-tete X-Requested-With ni user-agent navigateur (comme curl par defaut)"
curl -sS -m 60 -o sonde/full_plain.csv -w "HTTP %{http_code} %{size_download} octets\n" \
  --data "format=csv&decimal_separator=.&date_form=Y-m-d&op=&adjusted=Y&base100=&startdate=2025-01-01&enddate=2025-12-31" \
  "https://live.euronext.com/en/ajax/AwlHistoricalPrice/getFullDownloadAjax/$ISIN"
head -c 300 sonde/full_plain.csv; echo
echo "### 7. Lignes autour du dividende du 24/04/2025 (adjusted Y puis N)"
grep -E "2025-04-2[2-8]" sonde/full.csv; echo ---; grep -E "2025-04-2[2-8]" sonde/full_N.csv
