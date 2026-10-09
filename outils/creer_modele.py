"""Génère le classeur modèle Frontiere_Efficiente.xlsx (feuilles Mode d'emploi et Parametres)."""
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

SORTIE = sys.argv[1] if len(sys.argv) > 1 else "Frontiere_Efficiente.xlsx"
BLEU = "002060"
JAUNE = "FFE699"
entete = dict(font=Font(bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor=BLEU))

wb = Workbook()

# ---------------------------------------------------------------- Mode d'emploi
me = wb.active
me.title = "Mode d'emploi"
me.column_dimensions["A"].width = 4
me.column_dimensions["B"].width = 110
lignes = [
    ("t", "Frontière efficiente de Markowitz — mode d'emploi"),
    ("", ""),
    ("h", "Objectif"),
    ("", "À partir des cours journaliers de plusieurs actions, la macro calcule le rendement moyen, la variance et l'écart-type de chaque action,"),
    ("", "la matrice de covariance, le rendement et le risque d'un portefeuille selon les poids, puis trace la frontière efficiente."),
    ("", ""),
    ("h", "1. Choisir les actions et la période (feuille Parametres)"),
    ("", "• Source « Euronext » : la macro télécharge elle-même les cours sur live.euronext.com (connexion Internet nécessaire)."),
    ("", "• Pour chaque action : « Oui » pour l'inclure, son code, son nom et son code ISIN (ex. FR0000121014 pour LVMH), marché XPAR (Paris)."),
    ("", "• Période : dates de début et de fin. Euronext ne fournit que les deux dernières années."),
    ("", "• Dividendes : Euronext ne les fournit pas ; ils sont lus dans la feuille Dividendes (déjà remplie pour 2025)."),
    ("", "• Source « Fichiers CSV » (secours, ou période plus ancienne) : un fichier CODE.csv par action dans le dossier du classeur"),
    ("", "   (export Euronext du bouton « Télécharger » de la page de l'action, export Yahoo Finance ou Google Sheets)."),
    ("", ""),
    ("h", "2. Réglages de l'affichage"),
    ("", "• Nombre de points : portefeuilles calculés sur la partie efficiente de la frontière (20 par défaut)."),
    ("", "• Portefeuilles aléatoires : nuage de portefeuilles tirés au hasard, affiché derrière la frontière (0 pour aucun)."),
    ("", "• Pas de la grille et titres A/B : réglages du tableau à deux titres (xA de 0 à 100 %)."),
    ("", ""),
    ("h", "3. Lancer"),
    ("", "• Cliquez sur « Lancer l'analyse ». En cas de problème (connexion, ISIN, paramètre…), un message explique quoi corriger."),
    ("", ""),
    ("h", "4. Lire les résultats"),
    ("", "• Cours : cours de clôture alignés sur les dates communes à toutes les actions."),
    ("", "• Rendements : R(t) = (D(t) + P(t) − P(t−1)) / P(t−1), avec D(t) le dividende détaché ce jour-là."),
    ("", "• Statistiques : E(R), Var(R), σ(R) (divisés par n, comme dans le cours), covariances, corrélations, et contrôle par les fonctions Excel."),
    ("", "• Deux titres : tableau xA / xB → R(pf), var(pf), σ(pf) et graphique, pour les titres A et B choisis."),
    ("", "• Frontiere : pour chaque niveau de rendement, le portefeuille de risque minimal (poids entre 0 et 100 %) et le graphique complet."),
    ("", "• Calculateur : saisissez vos propres poids, le rendement et le risque se calculent en direct."),
    ("", ""),
    ("h", "Hypothèses"),
    ("", "• Données journalières ; les valeurs annualisées (× 252 et × √252) sont données à titre indicatif."),
    ("", "• Pas de vente à découvert : chaque poids est compris entre 0 et 100 % et leur somme vaut 100 %."),
    ("", "• Variance, écart-type et covariance divisés par n (formules du cours « Utiliser la théorie du portefeuille »)."),
    ("", "• Cours réellement cotés (non ajustés) et dividendes réellement versés."),
    ("", ""),
    ("h", "Installation (une seule fois)"),
    ("", "1. Enregistrez ce classeur au format « Classeur Excel (prenant en charge les macros) » (.xlsm)."),
    ("", "2. Outils > Macro > Éditeur Visual Basic, puis Insertion > Module, et collez le contenu de CODE_COMPLET_a_copier.txt."),
    ("", "3. Revenez dans Excel : Outils > Macro > Macros… > CreerBoutons > Exécuter. Les boutons apparaissent sur Parametres."),
]
for i, (style, texte) in enumerate(lignes, start=1):
    c = me.cell(row=i, column=2, value=texte)
    if style == "t":
        c.font = Font(bold=True, size=16, color=BLEU)
    elif style == "h":
        c.font = Font(bold=True, size=12, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BLEU)

# ------------------------------------------------------------------- Parametres
pa = wb.create_sheet("Parametres")
for col, largeur in zip("ABCDEFG", (3, 52, 24, 22, 18, 12, 4)):
    pa.column_dimensions[col].width = largeur
pa["B1"] = "Paramètres de l'analyse"
pa["B1"].font = Font(bold=True, size=16, color=BLEU)
pa["B3"] = "Réglage"
pa["C3"] = "Valeur"
for c in ("B3", "C3"):
    pa[c].font, pa[c].fill = entete["font"], entete["fill"]
reglages = [
    ("Source des cours (Euronext ou Fichiers CSV)", "Euronext", None),
    ("Date de début de la période", date(2025, 1, 1), "dd/mm/yyyy"),
    ("Date de fin de la période", date(2025, 12, 31), "dd/mm/yyyy"),
    ("Dossier des fichiers CSV (source CSV ; vide = dossier du classeur)", None, None),
    ("Nombre de points de la frontière efficiente (3 à 200)", 20, "0"),
    ("Nombre de portefeuilles aléatoires (0 = aucun)", 2000, "0"),
    ("Pas de la grille du tableau à deux titres", 0.1, "0%"),
    ("Titre A du tableau à deux titres (code)", "BNP", None),
    ("Titre B du tableau à deux titres (code)", "AI", None),
]
for i, (libelle, valeur, fmt) in enumerate(reglages, start=4):
    pa.cell(row=i, column=2, value=libelle)
    c = pa.cell(row=i, column=3, value=valeur)
    c.fill = PatternFill("solid", fgColor=JAUNE)
    c.alignment = Alignment(horizontal="center")
    if fmt:
        c.number_format = fmt
dv_source = DataValidation(type="list", formula1='"Euronext,Fichiers CSV"', allow_blank=False)
pa.add_data_validation(dv_source)
dv_source.add("C4")

pa["B14"] = "Actions à analyser"
pa["B14"].font = Font(bold=True, size=12, color=BLEU)
titres = ("Utiliser ? (Oui / Non)", "Code (= nom du fichier CSV)", "Nom de l'entreprise", "Code ISIN (Euronext)", "Marché")
for col, texte in zip("BCDEF", titres):
    c = pa[f"{col}15"]
    c.value = texte
    c.font, c.fill = entete["font"], entete["fill"]
    c.alignment = Alignment(horizontal="center", wrap_text=True)
pa.row_dimensions[15].height = 30
actions = [("MC", "LVMH", "FR0000121014"), ("TTE", "TotalEnergies", "FR0000120271"),
           ("SAN", "Sanofi", "FR0000120578"), ("BNP", "BNP Paribas", "FR0000131104"),
           ("AI", "Air Liquide", "FR0000120073")]
for i in range(30):
    for col in "BCDEF":
        pa[f"{col}{16 + i}"].fill = PatternFill("solid", fgColor=JAUNE)
    for col in "BEF":
        pa[f"{col}{16 + i}"].alignment = Alignment(horizontal="center")
for i, (code, nom, isin) in enumerate(actions):
    pa[f"B{16 + i}"] = "Oui"
    pa[f"C{16 + i}"] = code
    pa[f"D{16 + i}"] = nom
    pa[f"E{16 + i}"] = isin
    pa[f"F{16 + i}"] = "XPAR"
dv = DataValidation(type="list", formula1='"Oui,Non"', allow_blank=True)
pa.add_data_validation(dv)
dv.add("B16:B45")
pa["B47"] = ("ISIN et marché : sur live.euronext.com, page de l'action (ex. FR0000121014-XPAR pour LVMH ; "
             "XPAR = Euronext Paris, XAMS = Amsterdam, XBRU = Bruxelles).")
pa["B47"].font = Font(italic=True, color="808080")
pa["H3"] = "Boutons (à créer une fois : macro CreerBoutons)"
pa["H3"].font = Font(italic=True, color="808080")

# ------------------------------------------------------------------- Dividendes
di = wb.create_sheet("Dividendes")
for col, largeur in zip("ABCD", (12, 22, 22, 70)):
    di.column_dimensions[col].width = largeur
di["A1"] = "Dividendes détachés (Euronext ne les fournit pas)"
di["A1"].font = Font(bold=True, size=14, color=BLEU)
di["A2"] = ("Utilisés dans R(t) = (D(t) + P(t) − P(t−1)) / P(t−1) quand la source des cours ne contient pas de dividendes. "
            "Ajoutez une ligne par dividende ; laissez la feuille vide pour ignorer les dividendes.")
di["A2"].font = Font(italic=True, color="808080")
for col, texte in zip("ABCD", ("Code", "Date de détachement", "Montant par action (€)", "Commentaire")):
    c = di[f"{col}3"]
    c.value = texte
    c.font, c.fill = entete["font"], entete["fill"]
dividendes = [
    ("TTE", date(2025, 1, 2), 0.79, "Acompte (avant le 1er rendement : sans effet)"),
    ("TTE", date(2025, 3, 26), 0.79, "Acompte"),
    ("MC", date(2025, 4, 24), 7.50, "Solde exercice 2024"),
    ("SAN", date(2025, 5, 12), 3.92, "Exercice 2024"),
    ("BNP", date(2025, 5, 19), 4.79, "Solde exercice 2024"),
    ("AI", date(2025, 5, 19), 3.30, "Exercice 2024"),
    ("TTE", date(2025, 6, 19), 0.85, "Acompte"),
    ("BNP", date(2025, 9, 26), 2.59, "Acompte exercice 2025"),
    ("TTE", date(2025, 10, 1), 0.85, "Acompte"),
    ("MC", date(2025, 12, 2), 5.50, "Acompte exercice 2025"),
    ("TTE", date(2025, 12, 31), 0.85, "Acompte"),
]
for i, (code, jour, montant, commentaire) in enumerate(dividendes, start=4):
    di.cell(i, 1, code).alignment = Alignment(horizontal="center")
    di.cell(i, 2, jour).number_format = "dd/mm/yyyy"
    di.cell(i, 3, montant).number_format = "0.00"
    di.cell(i, 4, commentaire)
    for j in (1, 2, 3):
        di.cell(i, j).fill = PatternFill("solid", fgColor=JAUNE)
di.cell(4 + len(dividendes) + 1, 1, "Source : historique des dividendes Yahoo Finance (montants réellement versés).").font = \
    Font(italic=True, color="808080")
wb.active = 1
wb.save(SORTIE)
print("Classeur créé :", SORTIE)
