"""Télécharge les cours journaliers de 5 actions du CAC 40 depuis Yahoo Finance.

Produit, pour chaque titre et chaque période, un CSV : Date,Close,AdjClose,Dividend
(séparateur virgule, point décimal, dates ISO AAAA-MM-JJ).
"""
import csv
import os
import time

import yfinance as yf

TICKERS = ["MC.PA", "TTE.PA", "SAN.PA", "BNP.PA", "AI.PA"]
PERIODS = {
    "2025": ("2025-01-01", "2026-01-01"),            # année civile 2025
    "12_derniers_mois": ("2025-10-01", "2026-10-08"),  # fenêtre glissante (secours)
}

for folder, (start, end) in PERIODS.items():
    out_dir = os.path.join("donnees", folder)
    os.makedirs(out_dir, exist_ok=True)
    for ticker in TICKERS:
        for attempt in range(5):
            df = yf.Ticker(ticker).history(start=start, end=end, interval="1d",
                                           auto_adjust=False, actions=True)
            if not df.empty:
                break
            time.sleep(5 * (attempt + 1))
        if df.empty:
            raise SystemExit(f"Aucune donnée pour {ticker} ({folder})")
        path = os.path.join(out_dir, f"{ticker.replace('.PA', '')}.csv")
        with open(path, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["Date", "Close", "AdjClose", "Dividend"])
            for idx, row in df.iterrows():
                w.writerow([idx.strftime("%Y-%m-%d"), f"{row['Close']:.4f}",
                            f"{row['Adj Close']:.4f}", f"{row.get('Dividends', 0.0):.4f}"])
        print(f"=== {path} ({len(df)} lignes)")
        with open(path) as f:
            print(f.read())
