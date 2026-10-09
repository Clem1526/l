# Version simplifiée et pédagogique

C'est la même analyse que la version complète. Le code est écrit pour être lu par un étudiant de niveau intermédiaire en VBA :
- **4 modules**, un par grande étape de la consigne ;
- des noms de variables en français ;
- des commentaires qui relient chaque calcul à la formule du cours ;
- des résultats en couleurs.

## Les fichiers

| Fichier | Rôle |
|---|---|
| `Frontiere_Efficiente.xlsx` | classeur modèle : *Mode d'emploi*, *Parametres*, *Dividendes* |
| `CODE_A_COPIER.txt` | **les 4 modules réunis, à coller dans un module VBA** |
| `vba/Module1_Lancement.bas` … `Module4_Affichage.bas` | les 4 modules séparés (pour lecture) |
| `MC.csv` … `AI.csv` | exports Euronext 2025, en secours pour la source *Fichiers CSV* |

## Installation (5 minutes)

1. Ouvrez `Frontiere_Efficiente.xlsx` et enregistrez-le en **.xlsm** dans ce dossier.
2. Ouvrez **Outils > Macro > Éditeur Visual Basic**, puis **Insertion > Module**. Collez tout le contenu de `CODE_A_COPIER.txt` et gardez un seul `Option Explicit`.
3. Faites **Débogage > Compiler VBAProject**. Si aucun message n'apparaît, tout est bon.
4. De retour dans Excel, lancez **Outils > Macro > Macros… > CreerBoutons**, puis cliquez sur **Lancer l'analyse**.

## Comment le code est organisé

```
Bouton « Lancer l'analyse »
   └─ Module1_Lancement.LancerAnalyse
        ├─ LireParametres                  (feuille Parametres)
        ├─ Module2_Donnees
        │    ├─ RecupererCours             étape 1 : Euronext (ou fichiers CSV)
        │    └─ AjouterDividendes          feuille Dividendes
        ├─ Module3_Calculs
        │    ├─ CalculerRendements         étape 2 : R(t) = (D + P(t) − P(t−1)) / P(t−1)
        │    ├─ CalculerMoyennes           étape 2 : E(R)
        │    ├─ CalculerCovariances        étapes 2-3 : Var(R) sur la diagonale, Cov(Ri,Rj)
        │    └─ TesterTousLesPortefeuilles étapes 4-5 : R(pf) et σ(pf) de chaque répartition
        └─ Module4_Affichage
             ├─ AfficherCours / AfficherStatistiques
             ├─ AfficherDeuxTitres         le tableau de la vidéo + graphique
             └─ AfficherFrontiere          tableau de la frontière + graphique
```

Les données partagées sont des variables `Public`, déclarées en haut du Module 1 : `Prix`, `Rendements`, `Moyennes`, `Covariances`, etc. Chaque module les lit et les remplit, ce qui évite de passer de longues listes de paramètres.

## La méthode pour trouver la frontière

1. **Tous les portefeuilles possibles sont testés.** On répartit 100 % entre les actions par pas de 5 %. Avec 5 actions, cela donne **10 626 portefeuilles**. Une procédure *récursive*, `ChoisirPoids`, joue le rôle de 5 boucles `For` imbriquées.
2. Pour chacun, on calcule **R(pf) = Σ xᵢ·E(Rᵢ)** et **σ(pf) = √(Σᵢ Σⱼ xᵢ·xⱼ·Cov(i, j))**.
3. On découpe l'intervalle des rendements en 20 tranches. Dans chaque tranche, on garde le portefeuille **le moins risqué**.
4. Les portefeuilles au-dessus du portefeuille de variance minimale forment la **frontière efficiente**, en vert. Ceux en dessous sont **dominés**, en gris.

Le pas de 5 % donne une précision suffisante. Le portefeuille de variance minimale a un σ de 0,919 %, contre 0,918 % avec une optimisation exacte (version complète).

## Les couleurs des résultats

- **Onglets** : Cours en bleu, Statistiques en vert, Deux titres en orange, Frontiere en doré, Portefeuilles en gris.
- **Rendements moyens** : en vert s'ils sont positifs, en rouge s'ils sont négatifs.
- **Matrice des corrélations** : plus une case est orange, plus les deux actions sont corrélées.
- **Tableaux** : en doré, le portefeuille le moins risqué ; en vert, les portefeuilles efficients ; en gris, les dominés.

## Testé

Le code a été exécuté dans LibreOffice, qui sait lancer du VBA, avec les vrais fichiers Euronext 2025, et comparé à un calcul indépendant en Python. Les statistiques et la frontière sont exactes.

Ce qui n'a **pas** pu être testé hors d'Excel pour Mac :
- les **graphiques** ;
- les **boutons** ;
- le **téléchargement par curl**.

En cas de souci de téléchargement, passez la source en *Fichiers CSV*.
