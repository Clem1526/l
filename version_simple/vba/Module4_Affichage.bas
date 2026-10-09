Attribute VB_Name = "Module4_Affichage"
Option Explicit
'==============================================================================
'  MODULE 4 : L'AFFICHAGE DES RESULTATS (tableaux en couleurs et graphiques)
'==============================================================================
'  Feuilles creees :
'    Cours          : cours de cloture et rendements journaliers
'    Statistiques   : E(R), variance, ecart-type, covariances, correlations
'    Deux titres    : le tableau de la video (xA, xB, R(pf), var(pf), sigma(pf))
'    Frontiere      : le tableau et le graphique de la frontiere efficiente
'    Portefeuilles  : la liste de tous les portefeuilles testes
'==============================================================================

' Couleurs utilisees (valeur = Rouge + 256 x Vert + 65536 x Bleu)
Private Const BLEU_FONCE As Long = 6567967      ' RGB(31, 56, 100)  : titres et en-tetes
Private Const BLEU As Long = 11892015           ' RGB(47, 117, 181) : frontiere efficiente
Private Const VERT As Long = 3965440            ' RGB(0, 130, 60)   : valeurs positives
Private Const VERT_CLAIR As Long = 14348258     ' RGB(226, 239, 218): portefeuilles efficients
Private Const ROUGE As Long = 192               ' RGB(192, 0, 0)    : valeurs negatives
Private Const DORE As Long = 6740479            ' RGB(255, 217, 102): portefeuille de variance minimale
Private Const GRIS As Long = 10921638           ' RGB(166, 166, 166): portefeuilles domines
Private Const GRIS_CLAIR As Long = 15921906     ' RGB(242, 242, 242): fond des tableaux
Private Const ORANGE As Long = 3243501          ' RGB(237, 125, 49) : actions sur le graphique


'==============================================================================
' Feuille "Cours" : cours de cloture (a gauche) et rendements (a droite)
'==============================================================================
Public Sub AfficherCours()
    Dim ws As Worksheet, t As Long, i As Long, colR As Long
    Set ws = NouvelleFeuille("Cours", BLEU)

    colR = NbActions + 3                               ' colonne ou commencent les rendements
    ws.Cells(1, 1).Value = "Date"
    ws.Cells(1, colR - 1).Value = "Date"
    For i = 1 To NbActions
        ws.Cells(1, 1 + i).Value = "Cours " & Codes(i)
        ws.Cells(1, colR + i - 1).Value = "R " & Codes(i)
    Next i
    For t = 1 To NbJours
        ws.Cells(t + 1, 1).Value = Dates(t)
        For i = 1 To NbActions
            ws.Cells(t + 1, 1 + i).Value = Prix(t, i)
            ' Le rendement du jour t existe a partir du 2e jour
            If t >= 2 Then
                ws.Cells(t + 1, colR - 1).Value = Dates(t)
                ws.Cells(t + 1, colR + i - 1).Value = Rendements(t - 1, i)
            End If
        Next i
    Next t

    EnTete ws.Range(ws.Cells(1, 1), ws.Cells(1, 1 + NbActions))
    EnTete ws.Range(ws.Cells(1, colR - 1), ws.Cells(1, colR + NbActions - 1))
    ws.Columns(1).NumberFormat = "dd/mm/yyyy"
    ws.Columns(colR - 1).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(2, colR), ws.Cells(NbJours + 1, colR + NbActions - 1)).NumberFormat = "0.00%"
    ws.Columns.AutoFit
End Sub


'==============================================================================
' Feuille "Statistiques" (etapes 2 et 3)
'==============================================================================
Public Sub AfficherStatistiques()
    Dim ws As Worksheet, i As Long, j As Long, ligne As Long, plage As String
    Set ws = NouvelleFeuille("Statistiques", VERT)
    TitreFeuille ws, "Statistiques des rendements journaliers (" & (NbJours - 1) & " rendements)"

    ' --- Tableau 1 : une colonne par action
    ws.Range("A3").Value = "Indicateur"
    ws.Range("A4").Value = "Rendement moyen E(R)"
    ws.Range("A5").Value = "Variance Var(R)"
    ws.Range("A6").Value = "Ecart-type sigma(R)"
    ws.Range("A7").Value = "Rendement annualise (x 252)"
    ws.Range("A8").Value = "Ecart-type annualise (x racine 252)"
    For i = 1 To NbActions
        ws.Cells(3, i + 1).Value = Codes(i)
        ws.Cells(4, i + 1).Value = Moyennes(i)
        ws.Cells(5, i + 1).Value = Covariances(i, i)
        ws.Cells(6, i + 1).Value = EcartType(i)
        ws.Cells(7, i + 1).Value = Moyennes(i) * 252
        ws.Cells(8, i + 1).Value = EcartType(i) * Sqr(252)
        ' Rendement positif en vert, negatif en rouge
        If Moyennes(i) >= 0 Then ws.Cells(4, i + 1).Font.Color = VERT Else ws.Cells(4, i + 1).Font.Color = ROUGE
        If Moyennes(i) >= 0 Then ws.Cells(7, i + 1).Font.Color = VERT Else ws.Cells(7, i + 1).Font.Color = ROUGE
    Next i
    EnTete ws.Range(ws.Cells(3, 1), ws.Cells(3, NbActions + 1))
    ws.Range(ws.Cells(4, 2), ws.Cells(4, NbActions + 1)).NumberFormat = "0.000%"
    ws.Range(ws.Cells(5, 2), ws.Cells(5, NbActions + 1)).NumberFormat = "0.000000"
    ws.Range(ws.Cells(6, 2), ws.Cells(6, NbActions + 1)).NumberFormat = "0.000%"
    ws.Range(ws.Cells(7, 2), ws.Cells(8, NbActions + 1)).NumberFormat = "0.0%"

    ' --- Tableau 2 : matrice de variance-covariance
    ligne = 11
    ws.Cells(ligne - 1, 1).Value = "Matrice de variance-covariance"
    ws.Cells(ligne - 1, 1).Font.Bold = True
    For i = 1 To NbActions
        ws.Cells(ligne, i + 1).Value = Codes(i)
        ws.Cells(ligne + i, 1).Value = Codes(i)
        For j = 1 To NbActions
            ws.Cells(ligne + i, j + 1).Value = Covariances(i, j)
            If i = j Then ws.Cells(ligne + i, j + 1).Interior.Color = GRIS_CLAIR   ' diagonale = variances
        Next j
    Next i
    EnTete ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, NbActions + 1))
    ws.Range(ws.Cells(ligne + 1, 2), ws.Cells(ligne + NbActions, NbActions + 1)).NumberFormat = "0.000000"

    ' --- Tableau 3 : matrice des correlations, coloree (plus c'est orange, plus c'est correle)
    ligne = ligne + NbActions + 3
    ws.Cells(ligne - 1, 1).Value = "Matrice des correlations"
    ws.Cells(ligne - 1, 1).Font.Bold = True
    For i = 1 To NbActions
        ws.Cells(ligne, i + 1).Value = Codes(i)
        ws.Cells(ligne + i, 1).Value = Codes(i)
        For j = 1 To NbActions
            ws.Cells(ligne + i, j + 1).Value = Correlation(i, j)
            ws.Cells(ligne + i, j + 1).Interior.Color = CouleurCorrelation(Correlation(i, j))
        Next j
    Next i
    EnTete ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, NbActions + 1))
    ws.Range(ws.Cells(ligne + 1, 2), ws.Cells(ligne + NbActions, NbActions + 1)).NumberFormat = "0.00"

    ' --- Tableau 4 : verification avec les fonctions d'Excel (doit redonner les memes valeurs)
    ligne = ligne + NbActions + 3
    ws.Cells(ligne - 1, 1).Value = "Verification avec les fonctions Excel"
    ws.Cells(ligne - 1, 1).Font.Bold = True
    ws.Cells(ligne, 1).Value = "Fonction"
    ws.Cells(ligne + 1, 1).Value = "MOYENNE"
    ws.Cells(ligne + 2, 1).Value = "VAR.P.N"
    ws.Cells(ligne + 3, 1).Value = "ECARTYPE.PEARSON"
    For i = 1 To NbActions
        ' plage des rendements de l'action i dans la feuille Cours
        plage = "Cours!" & ThisWorkbook.Worksheets("Cours").Range(ThisWorkbook.Worksheets("Cours").Cells(3, NbActions + 2 + i), _
                ThisWorkbook.Worksheets("Cours").Cells(NbJours + 1, NbActions + 2 + i)).Address
        ws.Cells(ligne, i + 1).Value = Codes(i)
        ws.Cells(ligne + 1, i + 1).Formula = "=AVERAGE(" & plage & ")"
        ws.Cells(ligne + 2, i + 1).Formula = "=VAR.P(" & plage & ")"
        ws.Cells(ligne + 3, i + 1).Formula = "=STDEV.P(" & plage & ")"
    Next i
    EnTete ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, NbActions + 1))
    ws.Range(ws.Cells(ligne + 1, 2), ws.Cells(ligne + 1, NbActions + 1)).NumberFormat = "0.000%"
    ws.Range(ws.Cells(ligne + 2, 2), ws.Cells(ligne + 2, NbActions + 1)).NumberFormat = "0.000000"
    ws.Range(ws.Cells(ligne + 3, 2), ws.Cells(ligne + 3, NbActions + 1)).NumberFormat = "0.000%"

    ws.Columns(1).ColumnWidth = 36
    ws.Range(ws.Columns(2), ws.Columns(NbActions + 1)).ColumnWidth = 12
End Sub


'==============================================================================
' Feuille "Deux titres" : le tableau de la video, pour les titres A et B
'==============================================================================
Public Sub AfficherDeuxTitres()
    Dim ws As Worksheet, poids() As Double, k As Long, nbLignes As Long, ligne As Long
    Dim ligneMin As Long, risqueMin As Double

    Set ws = NouvelleFeuille("Deux titres", ORANGE)
    TitreFeuille ws, "Portefeuille de deux titres : A = " & Codes(TitreA) & " et B = " & Codes(TitreB)

    ' --- Caracteristiques des deux titres
    ws.Range("B3").Value = "Titre A : " & Codes(TitreA)
    ws.Range("C3").Value = "Titre B : " & Codes(TitreB)
    ws.Range("A4").Value = "E(R)":      ws.Range("B4").Value = Moyennes(TitreA):          ws.Range("C4").Value = Moyennes(TitreB)
    ws.Range("A5").Value = "sigma(R)":  ws.Range("B5").Value = EcartType(TitreA):         ws.Range("C5").Value = EcartType(TitreB)
    ws.Range("A6").Value = "Var(R)":    ws.Range("B6").Value = Covariances(TitreA, TitreA): ws.Range("C6").Value = Covariances(TitreB, TitreB)
    ws.Range("E3").Value = "Covariance(A,B)": ws.Range("E4").Value = Covariances(TitreA, TitreB)
    ws.Range("E5").Value = "Correlation":     ws.Range("E6").Value = Correlation(TitreA, TitreB)
    EnTete ws.Range("B3:C3"): EnTete ws.Range("E3"): EnTete ws.Range("E5")
    ws.Range("B4:C5").NumberFormat = "0.000%"
    ws.Range("B6:C6,E4").NumberFormat = "0.000000"
    ws.Range("E6").NumberFormat = "0.00"

    ' --- Tableau : xA varie de 0 a 100 %, xB = 100 % - xA
    ws.Range("A9:E9").Value = Array("xA (" & Codes(TitreA) & ")", "xB (" & Codes(TitreB) & ")", "R(pf)", "var(pf)", "sigma(pf)")
    EnTete ws.Range("A9:E9")
    ReDim poids(1 To NbActions)                       ' toutes les autres actions restent a 0 %
    nbLignes = Round(1 / PasGrille) + 1
    risqueMin = 1E+30
    For k = 1 To nbLignes
        ligne = 9 + k
        poids(TitreA) = (k - 1) * PasGrille
        poids(TitreB) = 1 - poids(TitreA)
        ws.Cells(ligne, 1).Value = poids(TitreA)
        ws.Cells(ligne, 2).Value = poids(TitreB)
        ws.Cells(ligne, 3).Value = RendementPortefeuille(poids)
        ws.Cells(ligne, 4).Value = RisquePortefeuille(poids) ^ 2
        ws.Cells(ligne, 5).Value = RisquePortefeuille(poids)
        If k Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 5)).Interior.Color = GRIS_CLAIR
        If RisquePortefeuille(poids) < risqueMin Then
            risqueMin = RisquePortefeuille(poids)
            ligneMin = ligne
        End If
    Next k
    ws.Range(ws.Cells(ligneMin, 1), ws.Cells(ligneMin, 5)).Interior.Color = DORE   ' le moins risque
    ws.Cells(ligneMin, 6).Value = "<- le moins risque"
    ws.Range(ws.Cells(10, 1), ws.Cells(9 + nbLignes, 2)).NumberFormat = "0%"
    ws.Range(ws.Cells(10, 3), ws.Cells(9 + nbLignes, 3)).NumberFormat = "0.000%"
    ws.Range(ws.Cells(10, 4), ws.Cells(9 + nbLignes, 4)).NumberFormat = "0.000000"
    ws.Range(ws.Cells(10, 5), ws.Cells(9 + nbLignes, 5)).NumberFormat = "0.000%"
    ws.Columns("A:F").ColumnWidth = 15

    ' --- Graphique : risque en abscisse, rendement en ordonnee
    ' (si Excel ne peut pas creer le graphique, on s'arrete la sans bloquer la macro)
    On Error GoTo SansGraphique
    Dim graphique As Chart
    Set graphique = NouveauGraphique(ws, ws.Range("H3"), "Frontiere efficiente : " & Codes(TitreA) & " et " & Codes(TitreB))
    AjouterSerie graphique, "Portefeuilles " & Codes(TitreA) & " / " & Codes(TitreB), _
                 ws.Range(ws.Cells(10, 5), ws.Cells(9 + nbLignes, 5)), _
                 ws.Range(ws.Cells(10, 3), ws.Cells(9 + nbLignes, 3)), BLEU, True
SansGraphique:
End Sub


'==============================================================================
' Feuilles "Portefeuilles" et "Frontiere" (etape 5)
'==============================================================================
'  Construction de la frontiere : on decoupe l'ecart entre le plus petit et le
'  plus grand rendement en "tranches" (NbTranches). Dans chaque tranche, on
'  garde le portefeuille le MOINS RISQUE. Les points situes au-dessus du
'  portefeuille de variance minimale forment la frontiere efficiente.
'==============================================================================
Public Sub AfficherFrontiere()
    Dim ws As Worksheet, wsListe As Worksheet
    Dim m As Long, i As Long, k As Long, ligne As Long, tranche As Long
    Dim rMin As Double, rMax As Double, largeur As Double
    Dim meilleur() As Long                             ' meilleur(k) = portefeuille le moins risque de la tranche k
    Dim sortie() As Variant, premiereEfficiente As Long, derniere As Long

    ' --- 1) Feuille "Portefeuilles" : tous les portefeuilles testes (ecrits d'un coup, c'est plus rapide)
    Set wsListe = NouvelleFeuille("Portefeuilles", GRIS)
    ReDim sortie(1 To NbPortefeuilles + 1, 1 To NbActions + 2)
    sortie(1, 1) = "sigma(pf)": sortie(1, 2) = "R(pf)"
    For i = 1 To NbActions
        sortie(1, 2 + i) = "x " & Codes(i)
    Next i
    For m = 1 To NbPortefeuilles
        sortie(m + 1, 1) = RisquesPf(m)
        sortie(m + 1, 2) = RendementsPf(m)
        For i = 1 To NbActions
            sortie(m + 1, 2 + i) = PoidsPortefeuilles(m, i)
        Next i
    Next m
    wsListe.Range(wsListe.Cells(1, 1), wsListe.Cells(NbPortefeuilles + 1, NbActions + 2)).Value = sortie
    EnTete wsListe.Range(wsListe.Cells(1, 1), wsListe.Cells(1, NbActions + 2))
    wsListe.Range(wsListe.Cells(2, 1), wsListe.Cells(NbPortefeuilles + 1, 2)).NumberFormat = "0.000%"
    wsListe.Range(wsListe.Cells(2, 3), wsListe.Cells(NbPortefeuilles + 1, NbActions + 2)).NumberFormat = "0%"

    ' --- 2) Le portefeuille le moins risque de chaque tranche de rendement
    rMin = RendementsPf(1): rMax = RendementsPf(1)
    For m = 2 To NbPortefeuilles
        If RendementsPf(m) < rMin Then rMin = RendementsPf(m)
        If RendementsPf(m) > rMax Then rMax = RendementsPf(m)
    Next m
    largeur = (rMax - rMin) / NbTranches
    If largeur = 0 Then largeur = 1                    ' cas extreme : toutes les actions ont le meme rendement
    ReDim meilleur(1 To NbTranches)
    For m = 1 To NbPortefeuilles
        tranche = Int((RendementsPf(m) - rMin) / largeur) + 1
        If tranche > NbTranches Then tranche = NbTranches   ' le rendement maximal va dans la derniere
        If meilleur(tranche) = 0 Then
            meilleur(tranche) = m
        ElseIf RisquesPf(m) < RisquesPf(meilleur(tranche)) Then
            meilleur(tranche) = m
        End If
    Next m

    ' --- 3) Feuille "Frontiere" : le portefeuille de variance minimale, puis le tableau
    Set ws = NouvelleFeuille("Frontiere", DORE)
    TitreFeuille ws, "Frontiere efficiente (" & NbActions & " actions, " & NbPortefeuilles & " portefeuilles testes)"
    ws.Range("A3").Value = "Portefeuille de variance minimale"
    ws.Range("A4").Value = "R(pf) = " & Format$(RendementsPf(IndiceVarianceMin), "0.000%") & _
                           "   sigma(pf) = " & Format$(RisquesPf(IndiceVarianceMin), "0.000%")
    ws.Range("A3:A4").Interior.Color = DORE
    ws.Range("A3").Font.Bold = True

    ws.Range("A6:E6").Value = Array("N", "Efficient ?", "R(pf)", "var(pf)", "sigma(pf)")
    For i = 1 To NbActions
        ws.Cells(6, 5 + i).Value = "x " & Codes(i)
    Next i
    EnTete ws.Range(ws.Cells(6, 1), ws.Cells(6, 5 + NbActions))

    ligne = 6
    For k = 1 To NbTranches
        If meilleur(k) > 0 Then                        ' une tranche peut etre vide
            m = meilleur(k)
            ligne = ligne + 1
            ws.Cells(ligne, 1).Value = ligne - 6
            ws.Cells(ligne, 3).Value = RendementsPf(m)
            ws.Cells(ligne, 4).Value = RisquesPf(m) ^ 2
            ws.Cells(ligne, 5).Value = RisquesPf(m)
            For i = 1 To NbActions
                ws.Cells(ligne, 5 + i).Value = PoidsPortefeuilles(m, i)
            Next i
            ' Couleur : vert = efficient, gris = domine, dore = variance minimale
            If RendementsPf(m) >= RendementsPf(IndiceVarianceMin) Then
                ws.Cells(ligne, 2).Value = "Oui"
                ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 5 + NbActions)).Interior.Color = VERT_CLAIR
                If premiereEfficiente = 0 Then premiereEfficiente = ligne
            Else
                ws.Cells(ligne, 2).Value = "Non (domine)"
                ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 5 + NbActions)).Font.Color = GRIS
            End If
            If m = IndiceVarianceMin Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 5 + NbActions)).Interior.Color = DORE
        End If
    Next k
    derniere = ligne
    ws.Range(ws.Cells(7, 3), ws.Cells(derniere, 3)).NumberFormat = "0.000%"
    ws.Range(ws.Cells(7, 4), ws.Cells(derniere, 4)).NumberFormat = "0.000000"
    ws.Range(ws.Cells(7, 5), ws.Cells(derniere, 5)).NumberFormat = "0.000%"
    ws.Range(ws.Cells(7, 6), ws.Cells(derniere, 5 + NbActions)).NumberFormat = "0%"
    ws.Columns(1).ColumnWidth = 6
    ws.Range(ws.Columns(2), ws.Columns(5 + NbActions)).ColumnWidth = 12

    ' --- 4) Graphique : nuage de tous les portefeuilles, frontiere, actions seules
    ' (si Excel ne peut pas creer le graphique, on s'arrete la sans bloquer la macro)
    On Error GoTo SansGraphique
    Dim graphique As Chart
    Set graphique = NouveauGraphique(ws, ws.Cells(3, 7 + NbActions), "Frontiere efficiente")
    AjouterSerie graphique, "Tous les portefeuilles testes", _
                 wsListe.Range(wsListe.Cells(2, 1), wsListe.Cells(NbPortefeuilles + 1, 1)), _
                 wsListe.Range(wsListe.Cells(2, 2), wsListe.Cells(NbPortefeuilles + 1, 2)), GRIS_CLAIR, False
    If premiereEfficiente > 7 Then
        AjouterSerie graphique, "Portefeuilles domines", ws.Range(ws.Cells(7, 5), ws.Cells(premiereEfficiente, 5)), _
                     ws.Range(ws.Cells(7, 3), ws.Cells(premiereEfficiente, 3)), GRIS, True
    End If
    AjouterSerie graphique, "Frontiere efficiente", ws.Range(ws.Cells(premiereEfficiente, 5), ws.Cells(derniere, 5)), _
                 ws.Range(ws.Cells(premiereEfficiente, 3), ws.Cells(derniere, 3)), BLEU, True
    AjouterSerie graphique, "Actions seules", _
                 ThisWorkbook.Worksheets("Statistiques").Range(ThisWorkbook.Worksheets("Statistiques").Cells(6, 2), ThisWorkbook.Worksheets("Statistiques").Cells(6, NbActions + 1)), _
                 ThisWorkbook.Worksheets("Statistiques").Range(ThisWorkbook.Worksheets("Statistiques").Cells(4, 2), ThisWorkbook.Worksheets("Statistiques").Cells(4, NbActions + 1)), ORANGE, False
SansGraphique:
End Sub


'==============================================================================
' Petits outils de mise en forme
'==============================================================================

' Cree une feuille vide (ou vide l'ancienne) avec un onglet de couleur.
Private Function NouvelleFeuille(ByVal nom As String, ByVal couleurOnglet As Long) As Worksheet
    Dim ws As Worksheet
    If FeuilleExiste(nom) Then
        Set ws = ThisWorkbook.Worksheets(nom)
        Do While ws.ChartObjects.Count > 0
            ws.ChartObjects(1).Delete
        Loop
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nom
    End If
    On Error Resume Next                               ' la couleur d'onglet est decorative
    ws.Tab.Color = couleurOnglet
    On Error GoTo 0
    Set NouvelleFeuille = ws
End Function

' Grand titre bleu en cellule A1.
Private Sub TitreFeuille(ByVal ws As Worksheet, ByVal texte As String)
    ws.Range("A1").Value = texte
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = BLEU_FONCE
End Sub

' En-tete de tableau : fond bleu fonce, texte blanc en gras, centre.
Private Sub EnTete(ByVal plage As Range)
    plage.Interior.Color = BLEU_FONCE
    plage.Font.Color = RGB(255, 255, 255)
    plage.Font.Bold = True
    plage.HorizontalAlignment = xlCenter
End Sub

' Couleur d'une correlation : blanc pour 0, de plus en plus orange quand elle se rapproche de 1.
Private Function CouleurCorrelation(ByVal c As Double) As Long
    If c < 0 Then c = 0
    If c > 1 Then c = 1
    CouleurCorrelation = RGB(255, 255 - Int(110 * c), 255 - Int(190 * c))
End Function

' Cree un graphique "nuage de points" avec titres d'axes.
Private Function NouveauGraphique(ByVal ws As Worksheet, ByVal position As Range, ByVal titre As String) As Chart
    Dim cadre As ChartObject
    Set cadre = ws.ChartObjects.Add(position.Left, position.Top, 520, 340)
    cadre.Chart.ChartType = xlXYScatter
    Do While cadre.Chart.SeriesCollection.Count > 0   ' Excel ajoute parfois une serie tout seul
        cadre.Chart.SeriesCollection(1).Delete
    Loop
    cadre.Chart.HasTitle = True
    cadre.Chart.ChartTitle.Text = titre
    cadre.Chart.Axes(xlCategory).HasTitle = True
    cadre.Chart.Axes(xlCategory).AxisTitle.Text = "Risque : ecart-type journalier sigma(pf)"
    cadre.Chart.Axes(xlValue).HasTitle = True
    cadre.Chart.Axes(xlValue).AxisTitle.Text = "Rendement journalier moyen R(pf)"
    cadre.Chart.HasLegend = True
    cadre.Chart.Legend.Position = xlLegendPositionBottom
    Set NouveauGraphique = cadre.Chart
End Function

' Ajoute une serie de points (x = risque, y = rendement), reliee par une ligne ou non.
Private Sub AjouterSerie(ByVal graphique As Chart, ByVal nom As String, ByVal x As Range, ByVal y As Range, _
                         ByVal couleur As Long, ByVal avecLigne As Boolean)
    Dim serie As Series
    Set serie = graphique.SeriesCollection.NewSeries
    serie.Name = nom
    serie.XValues = x
    serie.Values = y
    If avecLigne Then serie.ChartType = xlXYScatterLines Else serie.ChartType = xlXYScatter
    serie.MarkerStyle = xlMarkerStyleCircle
    If avecLigne Then serie.MarkerSize = 6 Else serie.MarkerSize = 4
    serie.MarkerBackgroundColor = couleur
    serie.MarkerForegroundColor = couleur
    If avecLigne Then serie.Format.Line.ForeColor.RGB = couleur
End Sub
