Attribute VB_Name = "modSorties"
Option Explicit
'==============================================================================
' modSorties : ecriture des resultats dans les feuilles du classeur
'==============================================================================

' Position des blocs de la feuille Statistiques (reutilisee par le calculateur)
Public Const STAT_LIGNE_ENTETE As Long = 3
Public Const STAT_LIGNE_MOYENNE As Long = 4
Public Const STAT_LIGNE_VARIANCE As Long = 5
Public Const STAT_LIGNE_ECART_TYPE As Long = 6
Public Const STAT_LIGNE_COV As Long = 13       ' premiere ligne de la matrice de covariance

Private Const COULEUR_ENTETE As Long = 6299648  ' RGB(0, 32, 96) : bleu fonce
Private Const COULEUR_SAISIE As Long = 10086143 ' RGB(255, 230, 153) : jaune clair

' Cree la feuille si besoin, sinon l'efface (cellules et graphiques).
Public Function PreparerFeuille(ByVal nom As String) As Worksheet
    Dim ws As Worksheet
    Set ws = FeuilleExistante(nom)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nom
    Else
        Do While ws.ChartObjects.Count > 0
            ws.ChartObjects(1).Delete
        Loop
        ws.Cells.Clear
    End If
    Set PreparerFeuille = ws
End Function

' Feuille "Cours" : dates et cours de cloture
Public Sub EcrireCours(ByRef p As TParametres, ByRef dates() As Date, ByRef prix() As Double)
    Dim ws As Worksheet, sortie() As Variant, t As Long, i As Long
    Set ws = PreparerFeuille(FEUILLE_COURS)
    ReDim sortie(0 To UBound(dates), 0 To p.NbTitres)
    sortie(0, 0) = "Date"
    For i = 1 To p.NbTitres
        sortie(0, i) = p.Codes(i)
    Next i
    For t = 1 To UBound(dates)
        sortie(t, 0) = dates(t)
        For i = 1 To p.NbTitres
            sortie(t, i) = prix(t, i)
        Next i
    Next t
    ws.Range("A1").Resize(UBound(dates) + 1, p.NbTitres + 1).Value = sortie
    ws.Range("A2").Resize(UBound(dates), 1).NumberFormat = "dd/mm/yyyy"
    ws.Range("B2").Resize(UBound(dates), p.NbTitres).NumberFormat = "0.00"
    FormaterEntete ws.Range("A1").Resize(1, p.NbTitres + 1)
    ws.Columns(1).ColumnWidth = 12
End Sub

' Feuille "Rendements" : R_t = (D_t + P_t - P_(t-1)) / P_(t-1)
Public Sub EcrireRendements(ByRef p As TParametres, ByRef dates() As Date, ByRef rend() As Double)
    Dim ws As Worksheet, sortie() As Variant, t As Long, i As Long, nbObs As Long
    Set ws = PreparerFeuille(FEUILLE_RENDEMENTS)
    nbObs = UBound(rend, 1)
    ReDim sortie(0 To nbObs, 0 To p.NbTitres)
    sortie(0, 0) = "Date"
    For i = 1 To p.NbTitres
        sortie(0, i) = p.Codes(i)
    Next i
    For t = 1 To nbObs
        sortie(t, 0) = dates(t + 1)            ' rendement entre la veille et ce jour
        For i = 1 To p.NbTitres
            sortie(t, i) = rend(t, i)
        Next i
    Next t
    ws.Range("A1").Resize(nbObs + 1, p.NbTitres + 1).Value = sortie
    ws.Range("A2").Resize(nbObs, 1).NumberFormat = "dd/mm/yyyy"
    ws.Range("B2").Resize(nbObs, p.NbTitres).NumberFormat = "0.000%"
    FormaterEntete ws.Range("A1").Resize(1, p.NbTitres + 1)
    ws.Columns(1).ColumnWidth = 12
End Sub

' Feuille "Statistiques" : E(R), Var(R), ecart-type, covariances, correlations,
' et verification par les fonctions d'Excel.
Public Sub EcrireStatistiques(ByRef p As TParametres, ByRef rend() As Double, _
                              ByRef moy() As Double, ByRef cov() As Double)
    Dim ws As Worksheet, i As Long, j As Long, n As Long, nbObs As Long
    Dim ligneCorr As Long, ligneVerif As Long, plage As String, plageJ As String

    Set ws = PreparerFeuille(FEUILLE_STATS)
    n = p.NbTitres
    nbObs = UBound(rend, 1)

    Titre ws, "Statistiques des rendements journaliers (" & nbObs & " rendements)"

    ' Bloc 1 : indicateurs par action
    ws.Cells(STAT_LIGNE_ENTETE, 1).Value = "Indicateur"
    ws.Cells(STAT_LIGNE_MOYENNE, 1).Value = "Rendement moyen journalier E(R)"
    ws.Cells(STAT_LIGNE_VARIANCE, 1).Value = "Variance journaliere Var(R)"
    ws.Cells(STAT_LIGNE_ECART_TYPE, 1).Value = "Ecart-type journalier sigma(R)"
    ws.Cells(7, 1).Value = "Rendement annualise (E(R) x " & JOURS_PAR_AN & ")"
    ws.Cells(8, 1).Value = "Ecart-type annualise (sigma x racine(" & JOURS_PAR_AN & "))"
    ws.Cells(9, 1).Value = "Nombre de rendements"
    For i = 1 To n
        ws.Cells(STAT_LIGNE_ENTETE, i + 1).Value = p.Codes(i)
        ws.Cells(STAT_LIGNE_MOYENNE, i + 1).Value = moy(i)
        ws.Cells(STAT_LIGNE_VARIANCE, i + 1).Value = cov(i, i)
        ws.Cells(STAT_LIGNE_ECART_TYPE, i + 1).Value = Sqr(cov(i, i))
        ws.Cells(7, i + 1).Value = moy(i) * JOURS_PAR_AN
        ws.Cells(8, i + 1).Value = Sqr(cov(i, i) * JOURS_PAR_AN)
        ws.Cells(9, i + 1).Value = nbObs
    Next i
    FormaterEntete ws.Cells(STAT_LIGNE_ENTETE, 1).Resize(1, n + 1)
    ws.Cells(STAT_LIGNE_MOYENNE, 2).Resize(1, n).NumberFormat = "0.0000%"
    ws.Cells(STAT_LIGNE_VARIANCE, 2).Resize(1, n).NumberFormat = "0.000000"
    ws.Cells(STAT_LIGNE_ECART_TYPE, 2).Resize(1, n).NumberFormat = "0.0000%"
    ws.Cells(7, 2).Resize(2, n).NumberFormat = "0.00%"

    ' Bloc 2 : matrice de variance-covariance
    ws.Cells(STAT_LIGNE_COV - 2, 1).Value = "Matrice de variance-covariance (journaliere)"
    ws.Cells(STAT_LIGNE_COV - 2, 1).Font.Bold = True
    EcrireMatrice ws, STAT_LIGNE_COV - 1, p, cov, False, "0.000000"

    ' Bloc 3 : matrice des correlations
    ligneCorr = STAT_LIGNE_COV + n + 2
    ws.Cells(ligneCorr, 1).Value = "Matrice des correlations"
    ws.Cells(ligneCorr, 1).Font.Bold = True
    EcrireMatrice ws, ligneCorr + 1, p, cov, True, "0.000"

    ' Bloc 4 : verification avec les fonctions d'Excel (doit donner les memes valeurs)
    ligneVerif = ligneCorr + n + 4
    ws.Cells(ligneVerif, 1).Value = "Verification avec les fonctions Excel (MOYENNE, VAR.P.N, ECARTYPE.PEARSON, COVARIANCE.PEARSON)"
    ws.Cells(ligneVerif, 1).Font.Bold = True
    ws.Cells(ligneVerif + 1, 1).Value = "Indicateur"
    ws.Cells(ligneVerif + 2, 1).Value = "E(R) par MOYENNE"
    ws.Cells(ligneVerif + 3, 1).Value = "Var(R) par VAR.P.N"
    ws.Cells(ligneVerif + 4, 1).Value = "sigma(R) par ECARTYPE.PEARSON"
    ws.Cells(ligneVerif + 5, 1).Value = "Ecart maximal avec la macro (doit etre ~0)"
    For i = 1 To n
        plage = PlageRendements(i, nbObs)
        ws.Cells(ligneVerif + 1, i + 1).Value = p.Codes(i)
        ws.Cells(ligneVerif + 2, i + 1).Formula = "=AVERAGE(" & plage & ")"
        ws.Cells(ligneVerif + 3, i + 1).Formula = "=VAR.P(" & plage & ")"
        ws.Cells(ligneVerif + 4, i + 1).Formula = "=STDEV.P(" & plage & ")"
        ws.Cells(ligneVerif + 5, i + 1).Formula = "=MAX(ABS(" & ws.Cells(ligneVerif + 2, i + 1).Address(False, False) & "-" & _
            ws.Cells(STAT_LIGNE_MOYENNE, i + 1).Address(False, False) & "),ABS(" & _
            ws.Cells(ligneVerif + 3, i + 1).Address(False, False) & "-" & _
            ws.Cells(STAT_LIGNE_VARIANCE, i + 1).Address(False, False) & "))"
    Next i
    FormaterEntete ws.Cells(ligneVerif + 1, 1).Resize(1, n + 1)
    ws.Cells(ligneVerif + 2, 2).Resize(1, n).NumberFormat = "0.0000%"
    ws.Cells(ligneVerif + 3, 2).Resize(1, n).NumberFormat = "0.000000"
    ws.Cells(ligneVerif + 4, 2).Resize(1, n).NumberFormat = "0.0000%"
    ws.Cells(ligneVerif + 5, 2).Resize(1, n).NumberFormat = "0.0E+00"

    ws.Cells(ligneVerif + 7, 1).Value = "Covariances par COVARIANCE.PEARSON"
    ws.Cells(ligneVerif + 7, 1).Font.Bold = True
    For i = 1 To n
        ws.Cells(ligneVerif + 8, i + 1).Value = p.Codes(i)
        ws.Cells(ligneVerif + 8 + i, 1).Value = p.Codes(i)
        For j = 1 To n
            plage = PlageRendements(i, nbObs)
            plageJ = PlageRendements(j, nbObs)
            ws.Cells(ligneVerif + 8 + i, j + 1).Formula = "=COVARIANCE.P(" & plage & "," & plageJ & ")"
        Next j
    Next i
    FormaterEntete ws.Cells(ligneVerif + 8, 1).Resize(1, n + 1)
    ws.Cells(ligneVerif + 9, 2).Resize(n, n).NumberFormat = "0.000000"

    ws.Columns(1).ColumnWidth = 44
    ws.Range(ws.Columns(2), ws.Columns(n + 1)).ColumnWidth = 12
End Sub

' Feuille "Deux titres" : le tableau de la video (xA, xB, R(pf), var(pf), sigma(pf))
Public Sub EcrireDeuxTitres(ByRef p As TParametres, ByRef moy() As Double, ByRef cov() As Double)
    Dim ws As Worksheet, a As Long, b As Long, k As Long, nbLignes As Long, ligne As Long
    Dim w() As Double, sortie() As Variant
    Const LIGNE_TABLEAU As Long = 10

    Set ws = PreparerFeuille(FEUILLE_DEUX_TITRES)
    a = p.IndiceTitreA
    b = p.IndiceTitreB
    Titre ws, "Portefeuille a deux titres : A = " & p.Codes(a) & ", B = " & p.Codes(b)

    ' Caracteristiques des deux titres
    ws.Range("B3").Value = "Titre A (" & p.Codes(a) & ")"
    ws.Range("C3").Value = "Titre B (" & p.Codes(b) & ")"
    ws.Range("A4").Value = "E(R)"
    ws.Range("A5").Value = "sigma(R)"
    ws.Range("A6").Value = "Var(R)"
    ws.Range("B4").Value = moy(a): ws.Range("C4").Value = moy(b)
    ws.Range("B5").Value = Sqr(cov(a, a)): ws.Range("C5").Value = Sqr(cov(b, b))
    ws.Range("B6").Value = cov(a, a): ws.Range("C6").Value = cov(b, b)
    ws.Range("E3").Value = "Covariance(A,B)"
    ws.Range("E4").Value = cov(a, b)
    ws.Range("E5").Value = "Correlation"
    ws.Range("E6").Value = Correlation(cov, a, b)
    FormaterEntete ws.Range("B3:C3")
    FormaterEntete ws.Range("E3")
    FormaterEntete ws.Range("E5")
    ws.Range("B4:C5").NumberFormat = "0.000%"
    ws.Range("B6:C6").NumberFormat = "0.000000"
    ws.Range("E4").NumberFormat = "0.000000"
    ws.Range("E6").NumberFormat = "0.00"

    ' Grille des poids
    nbLignes = CLng(1 / p.PasGrille) + 1
    ReDim sortie(0 To nbLignes, 1 To 5)
    sortie(0, 1) = "xA (" & p.Codes(a) & ")"
    sortie(0, 2) = "xB (" & p.Codes(b) & ")"
    sortie(0, 3) = "R(pf)"
    sortie(0, 4) = "var(pf)"
    sortie(0, 5) = "sigma(pf)"
    ReDim w(1 To p.NbTitres)
    For k = 1 To nbLignes
        w(a) = (k - 1) * p.PasGrille
        w(b) = 1 - w(a)
        sortie(k, 1) = w(a)
        sortie(k, 2) = w(b)
        sortie(k, 3) = RendementPortefeuille(w, moy)
        sortie(k, 4) = VariancePortefeuille(w, cov)
        sortie(k, 5) = EcartTypePortefeuille(w, cov)
    Next k
    ws.Cells(LIGNE_TABLEAU - 1, 1).Value = "Rendement et risque du portefeuille selon les poids"
    ws.Cells(LIGNE_TABLEAU - 1, 1).Font.Bold = True
    ws.Cells(LIGNE_TABLEAU, 1).Resize(nbLignes + 1, 5).Value = sortie
    FormaterEntete ws.Cells(LIGNE_TABLEAU, 1).Resize(1, 5)
    ligne = LIGNE_TABLEAU + 1
    ws.Cells(ligne, 1).Resize(nbLignes, 2).NumberFormat = "0%"
    ws.Cells(ligne, 3).Resize(nbLignes, 1).NumberFormat = "0.000%"
    ws.Cells(ligne, 4).Resize(nbLignes, 1).NumberFormat = "0.000000"
    ws.Cells(ligne, 5).Resize(nbLignes, 1).NumberFormat = "0.000%"
    ws.Range("A:E").ColumnWidth = 14

    GraphiqueDeuxTitres ws, ws.Cells(ligne, 5).Resize(nbLignes, 1), ws.Cells(ligne, 3).Resize(nbLignes, 1), _
                        p.Codes(a), p.Codes(b)
End Sub

' Feuille "Frontiere" : tableau de la frontiere efficiente (N actions) et graphique
Public Sub EcrireFrontiere(ByRef p As TParametres, ByRef poids() As Double, ByRef rendements() As Double, _
                           ByRef ecartsTypes() As Double, ByRef efficient() As Boolean, ByRef wMvp() As Double, _
                           ByRef moy() As Double, ByRef cov() As Double, ByVal nbDomines As Long, _
                           ByVal nbAleatoires As Long)
    Dim ws As Worksheet, k As Long, i As Long, nbPoints As Long, sortie() As Variant
    Dim premiereLigne As Long
    Const LIGNE_TABLEAU As Long = 8

    Set ws = PreparerFeuille(FEUILLE_FRONTIERE)
    nbPoints = UBound(rendements)
    Titre ws, "Frontiere efficiente (" & p.NbTitres & " actions, poids entre 0 et 100 %)"

    ' Portefeuille de variance minimale
    ws.Range("A3").Value = "Portefeuille de variance minimale"
    ws.Range("A3").Font.Bold = True
    ws.Range("A4").Value = "R(pf)"
    ws.Range("B4").Value = RendementPortefeuille(wMvp, moy)
    ws.Range("A5").Value = "sigma(pf)"
    ws.Range("B5").Value = EcartTypePortefeuille(wMvp, cov)
    ws.Range("B4:B5").NumberFormat = "0.0000%"
    For i = 1 To p.NbTitres
        ws.Cells(3, 3 + i).Value = p.Codes(i)
        ws.Cells(4, 3 + i).Value = wMvp(i)
    Next i
    FormaterEntete ws.Cells(3, 4).Resize(1, p.NbTitres)
    ws.Cells(4, 4).Resize(1, p.NbTitres).NumberFormat = "0.0%"

    ' Tableau de la frontiere
    ReDim sortie(0 To nbPoints, 1 To 7 + p.NbTitres)
    sortie(0, 1) = "N"
    sortie(0, 2) = "Efficient ?"
    sortie(0, 3) = "R(pf)"
    sortie(0, 4) = "var(pf)"
    sortie(0, 5) = "sigma(pf)"
    sortie(0, 6) = "R(pf) annualise"
    sortie(0, 7) = "sigma(pf) annualise"
    For i = 1 To p.NbTitres
        sortie(0, 7 + i) = "x " & p.Codes(i)
    Next i
    For k = 1 To nbPoints
        sortie(k, 1) = k
        If efficient(k) Then sortie(k, 2) = "Oui" Else sortie(k, 2) = "Non (domine)"
        sortie(k, 3) = rendements(k)
        sortie(k, 4) = ecartsTypes(k) ^ 2
        sortie(k, 5) = ecartsTypes(k)
        sortie(k, 6) = rendements(k) * JOURS_PAR_AN
        sortie(k, 7) = ecartsTypes(k) * Sqr(JOURS_PAR_AN)
        For i = 1 To p.NbTitres
            sortie(k, 7 + i) = poids(k, i)
        Next i
    Next k
    ws.Cells(LIGNE_TABLEAU - 1, 1).Value = "Frontiere de variance minimale : pour chaque rendement, le portefeuille le moins risque"
    ws.Cells(LIGNE_TABLEAU - 1, 1).Font.Bold = True
    ws.Cells(LIGNE_TABLEAU, 1).Resize(nbPoints + 1, 7 + p.NbTitres).Value = sortie
    FormaterEntete ws.Cells(LIGNE_TABLEAU, 1).Resize(1, 7 + p.NbTitres)
    premiereLigne = LIGNE_TABLEAU + 1
    ws.Cells(premiereLigne, 3).Resize(nbPoints, 1).NumberFormat = "0.0000%"
    ws.Cells(premiereLigne, 4).Resize(nbPoints, 1).NumberFormat = "0.000000"
    ws.Cells(premiereLigne, 5).Resize(nbPoints, 1).NumberFormat = "0.0000%"
    ws.Cells(premiereLigne, 6).Resize(nbPoints, 2).NumberFormat = "0.00%"
    ws.Cells(premiereLigne, 8).Resize(nbPoints, p.NbTitres).NumberFormat = "0.0%"
    If nbDomines > 0 Then ws.Cells(premiereLigne, 1).Resize(nbDomines, 7 + p.NbTitres).Font.Color = RGB(128, 128, 128)
    ws.Columns(1).ColumnWidth = 30
    ws.Range(ws.Columns(2), ws.Columns(7 + p.NbTitres)).ColumnWidth = 12

    GraphiqueFrontiere ws, premiereLigne, nbPoints, nbDomines, p, nbAleatoires
End Sub

' Feuille "Aleatoires" : nuage de portefeuilles tires au hasard (pour le graphique)
Public Sub EcrireAleatoires(ByRef rendements() As Double, ByRef ecartsTypes() As Double)
    Dim ws As Worksheet, sortie() As Variant, k As Long, nb As Long
    Set ws = PreparerFeuille(FEUILLE_ALEATOIRES)
    nb = UBound(rendements)
    ReDim sortie(0 To nb, 1 To 2)
    sortie(0, 1) = "sigma(pf)"
    sortie(0, 2) = "R(pf)"
    For k = 1 To nb
        sortie(k, 1) = ecartsTypes(k)
        sortie(k, 2) = rendements(k)
    Next k
    ws.Range("A1").Resize(nb + 1, 2).Value = sortie
    ws.Range("A2").Resize(nb, 2).NumberFormat = "0.000%"
    FormaterEntete ws.Range("A1:B1")
    ws.Columns("A:B").ColumnWidth = 12
End Sub

' Feuille "Calculateur" : l'utilisateur saisit des poids, Excel calcule R et sigma en direct.
Public Sub EcrireCalculateur(ByRef p As TParametres)
    Dim ws As Worksheet, i As Long, j As Long, n As Long
    Dim plagePoids As String, plageMoy As String, plageH As String, ligneRes As Long
    Const LIGNE_PREMIER As Long = 6          ' premiere ligne des actions
    Const COL_MATRICE As Long = 6            ' colonne F : debut de la matrice de covariance

    Set ws = PreparerFeuille(FEUILLE_CALCULATEUR)
    n = p.NbTitres
    Titre ws, "Calculateur : rendement et risque d'un portefeuille selon les poids"
    ws.Range("A2").Value = "Saisissez les poids dans les cases jaunes (total = 100 %). " & _
                           "Les resultats se mettent a jour automatiquement."

    ws.Cells(LIGNE_PREMIER - 1, 1).Value = "Action"
    ws.Cells(LIGNE_PREMIER - 1, 2).Value = "Poids x"
    ws.Cells(LIGNE_PREMIER - 1, 3).Value = "E(R)"
    ws.Cells(LIGNE_PREMIER - 1, 4).Value = "(Cov . x)"
    ws.Cells(LIGNE_PREMIER - 3, COL_MATRICE).Value = "Matrice de covariance (issue de la feuille Statistiques)"
    ws.Cells(LIGNE_PREMIER - 2, COL_MATRICE - 1).Value = "Poids x"
    plagePoids = ws.Cells(LIGNE_PREMIER, 2).Resize(n, 1).Address
    For i = 1 To n
        ws.Cells(LIGNE_PREMIER - 1 + i, 1).Value = p.Codes(i)
        ws.Cells(LIGNE_PREMIER - 1 + i, 2).Value = 1 / n
        ws.Cells(LIGNE_PREMIER - 1 + i, 3).Formula = "='" & FEUILLE_STATS & "'!" & _
            FeuilleExistante(FEUILLE_STATS).Cells(STAT_LIGNE_MOYENNE, i + 1).Address
        ' poids recopies en ligne pour le produit matriciel
        ws.Cells(LIGNE_PREMIER - 2, COL_MATRICE - 1 + i).Formula = "=INDEX(" & plagePoids & "," & i & ")"
        ws.Cells(LIGNE_PREMIER - 1, COL_MATRICE - 1 + i).Value = p.Codes(i)
        For j = 1 To n
            ws.Cells(LIGNE_PREMIER - 1 + i, COL_MATRICE - 1 + j).Formula = "='" & FEUILLE_STATS & "'!" & _
                FeuilleExistante(FEUILLE_STATS).Cells(STAT_LIGNE_COV - 1 + i, j + 1).Address
        Next j
        ' (Cov . x)_i = somme_j Cov(i, j) * x_j
        ws.Cells(LIGNE_PREMIER - 1 + i, 4).Formula = "=SUMPRODUCT(" & _
            ws.Cells(LIGNE_PREMIER - 1 + i, COL_MATRICE).Resize(1, n).Address & "," & _
            ws.Cells(LIGNE_PREMIER - 2, COL_MATRICE).Resize(1, n).Address & ")"
    Next i
    plageMoy = ws.Cells(LIGNE_PREMIER, 3).Resize(n, 1).Address
    plageH = ws.Cells(LIGNE_PREMIER, 4).Resize(n, 1).Address

    ligneRes = LIGNE_PREMIER + n + 1
    ws.Cells(ligneRes, 1).Value = "Somme des poids"
    ws.Cells(ligneRes, 2).Formula = "=SUM(" & plagePoids & ")"
    ws.Cells(ligneRes + 1, 1).Value = "Rendement du portefeuille R(pf)"
    ws.Cells(ligneRes + 1, 2).Formula = "=SUMPRODUCT(" & plagePoids & "," & plageMoy & ")"
    ws.Cells(ligneRes + 2, 1).Value = "Variance du portefeuille var(pf)"
    ws.Cells(ligneRes + 2, 2).Formula = "=SUMPRODUCT(" & plagePoids & "," & plageH & ")"
    ws.Cells(ligneRes + 3, 1).Value = "Ecart-type du portefeuille sigma(pf)"
    ws.Cells(ligneRes + 3, 2).Formula = "=SQRT(" & ws.Cells(ligneRes + 2, 2).Address & ")"
    ws.Cells(ligneRes + 4, 1).Value = "R(pf) annualise"
    ws.Cells(ligneRes + 4, 2).Formula = "=" & ws.Cells(ligneRes + 1, 2).Address & "*" & JOURS_PAR_AN
    ws.Cells(ligneRes + 5, 1).Value = "sigma(pf) annualise"
    ws.Cells(ligneRes + 5, 2).Formula = "=" & ws.Cells(ligneRes + 3, 2).Address & "*SQRT(" & JOURS_PAR_AN & ")"
    ws.Cells(ligneRes + 6, 1).Value = "Controle"
    ws.Cells(ligneRes + 6, 2).Formula = "=IF(ABS(" & ws.Cells(ligneRes, 2).Address & "-1)>0.0001," & _
        """Attention : la somme des poids doit faire 100 %""," & _
        "IF(MIN(" & plagePoids & ")<0,""Attention : poids negatif""," & """OK""))"

    FormaterEntete ws.Cells(LIGNE_PREMIER - 1, 1).Resize(1, 4)
    FormaterEntete ws.Cells(LIGNE_PREMIER - 1, COL_MATRICE).Resize(1, n)
    ws.Range(plagePoids).Interior.Color = COULEUR_SAISIE
    ws.Range(plagePoids).NumberFormat = "0.0%"
    ws.Cells(LIGNE_PREMIER - 2, COL_MATRICE).Resize(1, n).NumberFormat = "0.0%"
    ws.Range(plageMoy).NumberFormat = "0.0000%"
    ws.Range(plageH).NumberFormat = "0.0000000"
    ws.Cells(LIGNE_PREMIER, COL_MATRICE).Resize(n, n).NumberFormat = "0.000000"
    ws.Cells(ligneRes, 2).NumberFormat = "0.0%"
    ws.Cells(ligneRes + 1, 2).NumberFormat = "0.0000%"
    ws.Cells(ligneRes + 2, 2).NumberFormat = "0.000000"
    ws.Cells(ligneRes + 3, 2).NumberFormat = "0.0000%"
    ws.Cells(ligneRes + 4, 2).Resize(2, 1).NumberFormat = "0.00%"
    ws.Cells(ligneRes, 1).Resize(7, 1).Font.Bold = True
    ws.Columns(1).ColumnWidth = 36
    ws.Range(ws.Columns(2), ws.Columns(COL_MATRICE + n)).ColumnWidth = 12
End Sub

'------------------------------------------------------------------------------
' Outils de mise en forme
'------------------------------------------------------------------------------

Private Sub Titre(ByVal ws As Worksheet, ByVal texte As String)
    With ws.Range("A1")
        .Value = texte
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = COULEUR_ENTETE
    End With
End Sub

Public Sub FormaterEntete(ByVal plage As Range)
    With plage
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = COULEUR_ENTETE
        .HorizontalAlignment = xlCenter
    End With
End Sub

' Ecrit la matrice de covariance (ou de correlation) avec les codes en en-tete.
Private Sub EcrireMatrice(ByVal ws As Worksheet, ByVal ligneEntete As Long, ByRef p As TParametres, _
                          ByRef cov() As Double, ByVal enCorrelation As Boolean, ByVal formatNombre As String)
    Dim i As Long, j As Long, n As Long
    n = p.NbTitres
    For i = 1 To n
        ws.Cells(ligneEntete, i + 1).Value = p.Codes(i)
        ws.Cells(ligneEntete + i, 1).Value = p.Codes(i)
        For j = 1 To n
            If enCorrelation Then
                ws.Cells(ligneEntete + i, j + 1).Value = Correlation(cov, i, j)
            Else
                ws.Cells(ligneEntete + i, j + 1).Value = cov(i, j)
            End If
        Next j
    Next i
    FormaterEntete ws.Cells(ligneEntete, 1).Resize(1, n + 1)
    ws.Cells(ligneEntete + 1, 1).Resize(n, 1).Font.Bold = True
    ws.Cells(ligneEntete + 1, 2).Resize(n, n).NumberFormat = formatNombre
End Sub

' Reference (avec le nom de feuille) de la colonne des rendements de l'action i.
Private Function PlageRendements(ByVal i As Long, ByVal nbObs As Long) As String
    PlageRendements = "'" & FEUILLE_RENDEMENTS & "'!" & _
        FeuilleExistante(FEUILLE_RENDEMENTS).Cells(2, i + 1).Resize(nbObs, 1).Address
End Function
