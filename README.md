# Projet VBA : frontière efficiente (5 actions du CAC 40, année 2025)

> **Nouveau : une version simplifiée et pédagogique** (4 modules commentés, résultats en couleurs) se trouve dans [`version_simple/`](version_simple/). C'est la version recommandée pour le rendu. Le dossier `livrable/` contient la version complète, avec une optimisation exacte.

Classeur Excel + macro VBA qui :
1. **télécharge lui-même les cours journaliers de 2025 sur Euronext** (live.euronext.com), ou lit des fichiers CSV en secours ;
2. calcule pour chaque action le rendement journalier moyen, la variance et l'écart-type ;
3. calcule la matrice de covariance (et les corrélations) ;
4. donne le rendement et l'écart-type d'un portefeuille selon les poids ;
5. construit le tableau et le graphique de la frontière efficiente (2 titres comme dans la vidéo, puis 5 titres).

Les formules suivent le cours *Utiliser la théorie du portefeuille* :

| Indicateur | Formule |
|---|---|
| Rendement journalier | R(t) = (D(t) + P(t) − P(t−1)) / P(t−1), avec les cours réellement cotés sur Euronext et D(t) le dividende détaché ce jour-là (feuille *Dividendes*) |
| Moyenne | moyenne arithmétique |
| Variance, écart-type, covariance | **divisés par n** (fonctions `VAR.P.N`, `ECARTYPE.PEARSON` et `COVARIANCE.PEARSON` d'Excel) |
| Portefeuille | R(pf) = Σ xᵢ·E(Rᵢ) et var(pf) = Σᵢ Σⱼ xᵢ·xⱼ·Cov(i, j) |
| Poids | entre 0 et 100 %, somme = 100 % (pas de vente à découvert) |

---

## 1. Contenu du dossier `livrable/`

| Fichier | Rôle |
|---|---|
| `Frontiere_Efficiente.xlsx` | Classeur modèle : feuilles *Mode d'emploi*, *Parametres* (actions, codes ISIN, période, source) et *Dividendes* (dividendes de 2025, déjà remplis) |
| `vba/*.bas` | Les 9 modules VBA (pour lecture ; sur Mac l'import des .bas ne marche pas, utilisez le fichier suivant) |
| `CODE_COMPLET_a_copier.txt` | **Le code à copier-coller** dans un module VBA : les 9 modules réunis en un seul texte |
| `MC.csv`, `TTE.csv`, `SAN.csv`, `BNP.csv`, `AI.csv` | Secours : cours 2025 déjà téléchargés (Yahoo Finance, avec dividendes), pour la source *Fichiers CSV* |
| `Donnees_2_titres_BNP_AI.xlsx` | Données prêtes à l'emploi pour la version manuelle (étape 2) |

Actions retenues : LVMH (MC), TotalEnergies (TTE), Sanofi (SAN), BNP Paribas (BNP) et Air Liquide (AI).

## 2. Installation sur Mac (une seule fois, environ 5 minutes)

1. Téléchargez le dossier `livrable/` et gardez tous les fichiers ensemble.
2. Ouvrez `Frontiere_Efficiente.xlsx` dans Excel, puis **Fichier > Enregistrer sous…**. Choisissez le format **« Classeur Excel prenant en charge les macros (.xlsm) »** et le même dossier.
3. Ouvrez **Outils > Macro > Éditeur Visual Basic**. Cliquez sur `VBAProject (Frontiere_Efficiente.xlsm)`, puis faites **Insertion > Module**.
4. Ouvrez `CODE_COMPLET_a_copier.txt` avec TextEdit, faites **⌘A** puis **⌘C**, cliquez dans la fenêtre du module et faites **⌘V**. Il ne doit rester qu'**un seul** `Option Explicit`, tout en haut.
5. Faites **Débogage > Compiler VBAProject**. Si aucun message n'apparaît, tout est bon.
6. Revenez dans Excel : **Outils > Macro > Macros… > `CreerBoutons` > Exécuter**. Les boutons apparaissent sur *Parametres*. Enregistrez avec ⌘S.

> Sur Mac, l'import des fichiers `.bas` apparaît souvent grisé : c'est un défaut connu d'Excel pour Mac, d'où le copier-coller.
>
> Si Excel bloque les macros à l'ouverture, cliquez sur **Activer les macros**. Le réglage se trouve dans **Excel > Préférences > Sécurité**.

## 3. Utilisation

1. Sur la feuille *Parametres*, vérifiez les réglages :
   - **Source** : *Euronext*, valeur par défaut.
   - **Période** : du 01/01/2025 au 31/12/2025.
   - **Actions** : *Oui* ou *Non*, avec pour chacune son code, son nom, son **ISIN** et son marché (XPAR = Paris).
2. Cliquez sur **Lancer l'analyse**. La macro télécharge les cours des 5 actions sur Euronext, ce qui prend quelques secondes et demande une connexion Internet.
3. La macro crée les feuilles suivantes :

> **Euronext en bref.** Pour chaque action, la macro appelle l'export CSV du site, la même adresse que le bouton « Télécharger » de la page de l'action :
> `https://live.euronext.com/en/ajax/AwlHistoricalPrice/getFullDownloadAjax/<ISIN>-XPAR?format=csv&adjusted=N&startdate=…&enddate=…`
> - **Mac** : la macro utilise l'outil système `curl`. **Windows** : elle utilise l'objet `MSXML2.XMLHTTP`.
> - **Dividendes** : Euronext ne les fournit pas, ils viennent de la feuille *Dividendes*.
> - **Limite** : Euronext ne donne que les **deux dernières années**.
>
> **En secours**, si le téléchargement échoue (pas de connexion, Mac qui bloque `curl`…) :
> 1. Choisissez la source **Fichiers CSV**.
> 2. Placez à côté du classeur un fichier `CODE.csv` par action. Les fichiers Yahoo fournis conviennent, tout comme un export fait à la main depuis la page de l'action sur live.euronext.com (bouton de téléchargement de l'historique, format CSV, renommé par exemple `MC.csv`).


| Feuille | Contenu |
|---|---|
| **Cours** | cours de clôture alignés sur les dates communes |
| **Rendements** | rendements journaliers |
| **Statistiques** | E(R), Var(R), σ(R), valeurs annualisées, matrice de covariance, corrélations, **contrôle par les fonctions Excel** (l'écart doit être de 0) |
| **Deux titres** | le tableau de la vidéo (xA, xB, R(pf), var(pf), σ(pf)) et son graphique |
| **Frontiere** | portefeuille de variance minimale, tableau de la frontière (rendement, risque et poids de chaque action) et graphique complet |
| **Aleatoires** | 2 000 portefeuilles tirés au hasard (le nuage gris du graphique) |
| **Calculateur** | vous saisissez des poids, et le rendement et le risque se calculent en direct (formules Excel) |

## 4. Étape 2 : la version manuelle (conseil n°2 de la consigne)

Elle est à faire à la main dans un classeur séparé, avec `Donnees_2_titres_BNP_AI.xlsx`. Les colonnes sont A = Date, B = Cours BNP, C = Dividende BNP, D = Cours AI et E = Dividende AI, sur les lignes 2 à 256.

| Étape | Cellule | Formule (Excel en français) |
|---|---|---|
| Rendement BNP | G3, puis tirer jusqu'à G256 | `=(C3+B3-B2)/B2` |
| Rendement AI | H3, puis tirer jusqu'à H256 | `=(E3+D3-D2)/D2` |
| E(R) | | `=MOYENNE(G3:G256)` |
| Var(R) | | `=VAR.P.N(G3:G256)` |
| σ(R) | | `=ECARTYPE.PEARSON(G3:G256)` |
| Covariance | | `=COVARIANCE.PEARSON(G3:G256;H3:H256)` |
| Corrélation | | `=COEFFICIENT.CORRELATION(G3:G256;H3:H256)` |
| R(pf) | | `=xA*E(R_BNP) + xB*E(R_AI)` |
| var(pf) | | `=xA^2*Var_BNP + xB^2*Var_AI + 2*xA*xB*Cov` |
| σ(pf) | | `=RACINE(var(pf))` |

Pour le graphique : **Insertion > Nuage de points avec courbes et marqueurs**, avec σ(pf) en abscisse (X) et R(pf) en ordonnée (Y).

**Valeurs attendues**, à vérifier :

| | BNP | AI |
|---|---|---|
| E(R) journalier | 0,1745 % | 0,0240 % |
| Var(R) | 0,000310 | 0,000121 |
| σ(R) journalier | 1,7619 % | 1,0981 % |

- Covariance(BNP, AI) = 0,0000819
- Corrélation = 0,423

| xA (BNP) | xB (AI) | R(pf) | var(pf) | σ(pf) |
|---|---|---|---|---|
| 0 % | 100 % | 0,0240 % | 0,000121 | 1,0981 % |
| 10 % | 90 % | 0,0390 % | 0,000116 | 1,0747 % |
| 20 % | 80 % | 0,0541 % | 0,000116 | 1,0760 % |
| 30 % | 70 % | 0,0691 % | 0,000121 | 1,1018 % |
| 40 % | 60 % | 0,0842 % | 0,000132 | 1,1505 % |
| 50 % | 50 % | 0,0992 % | 0,000149 | 1,2193 % |
| 60 % | 40 % | 0,1143 % | 0,000170 | 1,3051 % |
| 70 % | 30 % | 0,1293 % | 0,000197 | 1,4048 % |
| 80 % | 20 % | 0,1444 % | 0,000230 | 1,5156 % |
| 90 % | 10 % | 0,1594 % | 0,000267 | 1,6352 % |
| 100 % | 0 % | 0,1745 % | 0,000310 | 1,7619 % |

Pour les 5 actions, le **portefeuille de variance minimale** est R(pf) = 0,0272 % et σ(pf) = 0,9176 % par jour. Il est composé d'environ 51 % d'AI, 34 % de TTE, 12 % de SAN, 1 % de BNP et 1 % de MC.

## 5. Organisation du code (pour l'oral)

| Module | Rôle |
|---|---|
| `modMain` | macros des boutons : `LancerAnalyse`, `EffacerResultats`, `CreerBoutons` ; gestion des erreurs |
| `modConfig` | noms des feuilles et des cellules, lecture et contrôle des paramètres (source, période, ISIN…) |
| `modEuronext` | téléchargement des cours sur Euronext (curl sur Mac, MSXML2.XMLHTTP sur Windows) |
| `modImport` | lecture des données (export Euronext, CSV Yahoo ou Google Sheets), période, dividendes de la feuille *Dividendes*, alignement des dates communes |
| `modStats` | rendements, moyennes, matrice de covariance, corrélation |
| `modPortefeuille` | rendement, variance et écart-type d'un portefeuille selon les poids |
| `modOptimisation` | portefeuille de variance minimale et frontière efficiente (voir ci-dessous) |
| `modSorties` | écriture des feuilles de résultats et du calculateur |
| `modGraphiques` | graphiques risque/rendement |

**Méthode d'optimisation**, sans le Solveur, qui est peu fiable sur Mac. Pour un « goût pour le rendement » λ, la macro minimise `var(pf) − λ·R(pf)` par **descente de gradient projetée**. À chaque itération, elle fait un petit pas qui réduit le critère, puis ramène les poids à des valeurs positives dont la somme fait 100 %.
- λ = 0 donne le portefeuille de variance minimale.
- Pour chaque rendement visé, la macro cherche le bon λ par **dichotomie** : elle obtient le portefeuille de risque minimal pour ce rendement.

Les résultats ont été vérifiés avec un solveur indépendant (Python, scipy) : l'écart est inférieur à 10⁻¹¹.

## 6. Ce qui a été testé, et ce qui reste à tester sur Mac

✅ **Testé dans LibreOffice**, qui exécute le VBA, et comparé à un calcul Python de référence (écart inférieur à 10⁻¹⁵) :
- tous les calculs, la feuille *Dividendes* et le choix de la période ;
- le mode Euronext, avec les **vrais fichiers renvoyés par Euronext** (récupérés via GitHub Actions) ;
- le mode *Fichiers CSV*, avec les fichiers Yahoo comme avec les exports Euronext ;
- les messages d'erreur : ISIN faux ou invalide, réponse vide d'Euronext, période de plus de 2 ans, dates inversées, source inconnue, dividende mal saisi, fichier absent, une seule action…

⚠️ **À tester dans Excel sur Mac**, car LibreOffice ne sait pas les exécuter :
- le **téléchargement via `curl`** ;
- les **graphiques** ;
- les **boutons**.

Si le téléchargement échoue, passez en source *Fichiers CSV* et envoyez une capture du message. Si un graphique ne se crée pas, la macro termine quand même les calculs et affiche un avertissement.

## 7. Dossiers techniques

| Dossier | Contenu |
|---|---|
| `donnees/` | données brutes : Yahoo (2025 et 12 derniers mois, `scripts/fetch_data.py`) et export Euronext 2025 (`scripts/sonder_euronext.sh`), téléchargées via GitHub Actions |
| `reference/reference.py` | calcul de référence en Python (mêmes formules, plus le contrôle par scipy) |
| `outils/` | génération du classeur modèle, test automatique dans LibreOffice, comparaison avec la référence |
