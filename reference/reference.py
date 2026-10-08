"""Calculs de référence en Python pour vérifier les résultats de la macro VBA.

Conventions du cours (document LEARN) :
  - rendement journalier R_t = (D_t + P_t - P_{t-1}) / P_{t-1}
  - moyenne arithmétique, variance / écart-type / covariance divisés par n
  - poids entre 0 et 100 %, somme = 100 % (pas de vente à découvert)
"""
import csv
import sys

import numpy as np
from scipy.optimize import minimize

DOSSIER = sys.argv[1] if len(sys.argv) > 1 else "donnees/2025"
TITRES = ["MC", "TTE", "SAN", "BNP", "AI"]


def lire(titre):
    rows = list(csv.DictReader(open(f"{DOSSIER}/{titre}.csv")))
    return ([r["Date"] for r in rows], np.array([float(r["Close"]) for r in rows]),
            np.array([float(r.get("Dividend") or 0) for r in rows]))


def rendements(close, div):
    return (div[1:] + close[1:] - close[:-1]) / close[:-1]


# --- même algorithme que le VBA : gradient projeté + recherche de lambda ---
def projeter_simplexe(v):
    u = np.sort(v)[::-1]
    css = np.cumsum(u)
    rho = max(j for j in range(len(u)) if u[j] - (css[j] - 1) / (j + 1) > 0)
    theta = (css[rho] - 1) / (rho + 1)
    return np.maximum(v - theta, 0)


def minimiser(cov_n, mu_n, lam, w):
    pas = 1 / (2 * np.abs(cov_n).sum(axis=1).max())
    for _ in range(100000):
        w_new = projeter_simplexe(w - pas * (2 * cov_n @ w - lam * mu_n))
        if np.abs(w_new - w).max() < 1e-13:
            return w_new
        w = w_new
    return w


def portefeuille_cible(cov, mu, cible, w0):
    s = np.abs(cov).sum(axis=1).max(); m = np.abs(mu).max()
    cov_n, mu_n = cov / s, mu / m
    w_mvp = minimiser(cov_n, mu_n, 0.0, w0)
    r_mvp = mu @ w_mvp
    sens = 1.0 if cible >= r_mvp else -1.0
    lo, hi = 0.0, sens
    w = w_mvp
    for _ in range(80):
        w = minimiser(cov_n, mu_n, hi, w)
        if sens * (mu @ w - cible) >= 0:
            break
        lo, hi = hi, hi * 2
    for _ in range(100):
        mid = (lo + hi) / 2
        w = minimiser(cov_n, mu_n, mid, w)
        if abs(mu @ w - cible) < 1e-9 * (mu.max() - mu.min()):
            break
        if sens * (mu @ w - cible) < 0:
            lo = mid
        else:
            hi = mid
    return w


def qp_scipy(cov, mu, cible):
    n = len(mu)
    cons = [{"type": "eq", "fun": lambda w: w.sum() - 1}]
    if cible is not None:
        cons.append({"type": "eq", "fun": lambda w: (mu @ w - cible) * 1e3})
    r = minimize(lambda w: w @ cov @ w * 1e4, np.ones(n) / n, bounds=[(0, 1)] * n,
                 constraints=cons, method="SLSQP", options={"ftol": 1e-15, "maxiter": 1000})
    return r.x


if __name__ == "__main__":
    data = {t: lire(t) for t in TITRES}
    R = np.column_stack([rendements(data[t][1], data[t][2]) for t in TITRES])
    mu = R.mean(axis=0)
    cov = np.cov(R, rowvar=False, bias=True)  # division par n
    np.set_printoptions(linewidth=150)
    print("Nombre de rendements :", len(R), "| plus forte variation :",
          {t: f"{np.abs(R[:, i]).max():.2%}" for i, t in enumerate(TITRES)})
    print(f"{'Titre':6}{'E(R)':>10}{'Var(R)':>12}{'sigma(R)':>10}")
    for i, t in enumerate(TITRES):
        print(f"{t:6}{mu[i]:>10.4%}{cov[i, i]:>12.6f}{np.sqrt(cov[i, i]):>10.4%}")
    print("Matrice de covariance :\n", cov)
    print("Corrélations :\n", np.corrcoef(R, rowvar=False).round(3))

    w0 = np.ones(len(mu)) / len(mu)
    s = np.abs(cov).sum(axis=1).max()
    w_mvp = minimiser(cov / s, mu / np.abs(mu).max(), 0.0, w0)
    print("\nPortefeuille de variance minimale :", dict(zip(TITRES, w_mvp.round(4))),
          f"R={mu @ w_mvp:.4%} sigma={np.sqrt(w_mvp @ cov @ w_mvp):.4%}")
    print("Contrôle SLSQP :", qp_scipy(cov, mu, None).round(4))

    print("\nFrontière (rendement cible -> sigma algo / sigma scipy, écart max des poids)")
    ecart_max = 0
    for cible in np.linspace(mu @ w_mvp, mu.max(), 21)[:-1]:
        w = portefeuille_cible(cov, mu, cible, w_mvp)
        ws = qp_scipy(cov, mu, cible)
        ecart_max = max(ecart_max, np.abs(w - ws).max())
        print(f"  {cible:.4%}  {np.sqrt(w @ cov @ w):.5%}  {np.sqrt(ws @ cov @ ws):.5%}  {w.round(3)}")
    print("Écart maximal des poids algo vs scipy :", f"{ecart_max:.2e}")
