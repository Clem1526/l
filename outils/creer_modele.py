"""Génère le classeur modèle Frontiere_Efficiente.xlsx (feuilles Mode d'emploi et Parametres)."""
import sys

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
    ("h", "1. Préparer les fichiers de cours"),
    ("", "• Un fichier CSV par action, nommé exactement comme le code saisi dans Parametres (ex. MC.csv pour le code MC)."),
    ("", "• Placez les fichiers dans le même dossier que ce classeur (ou indiquez leur dossier en cellule C4 de la feuille Parametres)."),
    ("", "• Formats acceptés : Date,Close,AdjClose,Dividend (fichiers fournis) ; export Yahoo Finance ; export Google Sheets (Date;Close)."),
    ("", ""),
    ("h", "2. Choisir les actions et les réglages (feuille Parametres)"),
    ("", "• Colonne « Utiliser ? » : Oui pour inclure l'action, Non pour l'exclure. Au moins 2 actions, jusqu'à 30."),
    ("", "• Nombre de points : nombre de portefeuilles calculés sur la partie efficiente de la frontière (20 par défaut)."),
    ("", "• Portefeuilles aléatoires : nuage de portefeuilles tirés au hasard, affiché derrière la frontière (0 pour aucun)."),
    ("", "• Pas de la grille et titres A/B : réglages du tableau à deux titres (xA de 0 à 100 %)."),
    ("", ""),
    ("h", "3. Lancer"),
    ("", "• Cliquez sur « Lancer l'analyse ». Sur Mac, autorisez Excel à lire les fichiers quand il le demande."),
    ("", "• En cas de problème (fichier introuvable, paramètre invalide…), un message explique quoi corriger."),
    ("", ""),
    ("h", "4. Lire les résultats"),
    ("", "• Cours : cours de clôture alignés sur les dates communes à toutes les actions."),
    ("", "• Rendements : R(t) = (D(t) + P(t) − P(t−1)) / P(t−1), avec D(t) le dividende versé ce jour-là."),
    ("", "• Statistiques : E(R), Var(R), σ(R) (divisés par n, comme dans le cours), covariances, corrélations, et contrôle par les fonctions Excel."),
    ("", "• Deux titres : tableau xA / xB → R(pf), var(pf), σ(pf) et graphique, pour les titres A et B choisis."),
    ("", "• Frontiere : pour chaque niveau de rendement, le portefeuille de risque minimal (poids entre 0 et 100 %) et le graphique complet."),
    ("", "• Calculateur : saisissez vos propres poids, le rendement et le risque se calculent en direct."),
    ("", ""),
    ("h", "Hypothèses"),
    ("", "• Données journalières ; les valeurs annualisées (× 252 et × √252) sont données à titre indicatif."),
    ("", "• Pas de vente à découvert : chaque poids est compris entre 0 et 100 % et leur somme vaut 100 %."),
    ("", "• Variance, écart-type et covariance divisés par n (formules du cours « Utiliser la théorie du portefeuille »)."),
    ("", ""),
    ("h", "Installation (une seule fois)"),
    ("", "1. Enregistrez ce classeur au format « Classeur Excel (prenant en charge les macros) » (.xlsm)."),
    ("", "2. Outils > Macro > Visual Basic Editor (ou Alt+F11 / Option+F11), puis Fichier > Importer : importez les 8 fichiers .bas."),
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
pa.column_dimensions["A"].width = 3
pa.column_dimensions["B"].width = 52
pa.column_dimensions["C"].width = 24
pa.column_dimensions["D"].width = 26
pa.column_dimensions["E"].width = 4
pa["B1"] = "Paramètres de l'analyse"
pa["B1"].font = Font(bold=True, size=16, color=BLEU)
pa["B3"] = "Réglage"
pa["C3"] = "Valeur"
for c in ("B3", "C3"):
    pa[c].font, pa[c].fill = entete["font"], entete["fill"]
reglages = [
    ("Dossier des fichiers CSV (vide = dossier du classeur)", None, None),
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

pa["B11"] = "Actions à analyser (un fichier CODE.csv par action)"
pa["B11"].font = Font(bold=True, size=12, color=BLEU)
for col, texte in zip("BCD", ("Utiliser ? (Oui / Non)", "Code (= nom du fichier)", "Nom de l'entreprise")):
    c = pa[f"{col}12"]
    c.value = texte
    c.font, c.fill = entete["font"], entete["fill"]
actions = [("MC", "LVMH"), ("TTE", "TotalEnergies"), ("SAN", "Sanofi"),
           ("BNP", "BNP Paribas"), ("AI", "Air Liquide")]
for i in range(30):
    for col in "BCD":
        pa[f"{col}{13 + i}"].fill = PatternFill("solid", fgColor=JAUNE)
    pa[f"B{13 + i}"].alignment = Alignment(horizontal="center")
for i, (code, nom) in enumerate(actions):
    pa[f"B{13 + i}"] = "Oui"
    pa[f"C{13 + i}"] = code
    pa[f"D{13 + i}"] = nom
dv = DataValidation(type="list", formula1='"Oui,Non"', allow_blank=True)
pa.add_data_validation(dv)
dv.add("B13:B42")

pa["F3"] = "Boutons (à créer une fois : macro CreerBoutons)"
pa["F3"].font = Font(italic=True, color="808080")
wb.active = 1
wb.save(SORTIE)
print("Classeur créé :", SORTIE)
