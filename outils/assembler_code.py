"""Assemble les 8 modules .bas en un seul texte à copier-coller dans un module VBA.

Les déclarations (Const, Type, variables publiques) doivent précéder les procédures
dans un module VBA : on les regroupe donc en tête, module par module.
"""
import glob
import os
import re
import sys

dossier, sortie = sys.argv[1], sys.argv[2]
ordre = ["modConfig", "modMain", "modEuronext", "modImport", "modStats", "modPortefeuille",
         "modOptimisation", "modSorties", "modGraphiques"]
declarations, procedures = [], []
for nom in ordre:
    lignes = open(os.path.join(dossier, nom + ".bas"), encoding="ascii").read().replace("\r", "").split("\n")
    lignes = [l for l in lignes if not l.startswith("Attribute VB_Name") and l.strip() != "Option Explicit"]
    # début de la première procédure (en remontant les commentaires qui la précèdent)
    debut = next(i for i, l in enumerate(lignes) if re.match(r"^(Public |Private )?(Sub|Function) ", l))
    while debut > 0 and lignes[debut - 1].startswith("'"):
        debut -= 1
    entete = "\n".join(lignes[:debut]).strip("\n")
    declarations.append(f"'{'=' * 78}\n' Module d'origine : {nom}\n{entete}\n")
    procedures.append(f"'{'#' * 78}\n' {nom}\n'{'#' * 78}\n" + "\n".join(lignes[debut:]).strip("\n") + "\n")
texte = ("Option Explicit\n' Frontiere efficiente - code complet (les 8 modules reunis en un seul)\n\n"
         + "\n".join(declarations) + "\n" + "\n".join(procedures))
open(sortie, "w", encoding="ascii", newline="\r\n").write(texte)
print(sortie, len(texte.splitlines()), "lignes")
