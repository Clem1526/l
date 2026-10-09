Attribute VB_Name = "modGraphiques"
Option Explicit
'==============================================================================
' modGraphiques : graphiques risque / rendement
' Si un graphique ne peut pas etre cree, les calculs ne sont pas perdus :
' un avertissement est affiche a la fin de l'analyse.
'==============================================================================

Public AvertissementGraphique As String

' Graphique de la feuille "Deux titres" (comme dans la video)
Public Sub GraphiqueDeuxTitres(ByVal ws As Worksheet, ByVal plageRisque As Range, ByVal plageRendement As Range, _
                               ByVal codeA As String, ByVal codeB As String)
    Dim ch As Chart
    On Error GoTo Echec
    Set ch = NouveauGraphique(ws, ws.Range("G3").Left, ws.Range("G3").Top, 480, 320)
    AjouterSerie ch, "Portefeuilles " & codeA & " / " & codeB, plageRisque, plageRendement, xlXYScatterLines
    MettreEnForme ch, "Frontiere efficiente : " & codeA & " et " & codeB, False
    Exit Sub
Echec:
    NoterEchec ws.Name, Err.Description
End Sub

' Graphique de la feuille "Frontiere" : nuage aleatoire, frontiere, actions, variance minimale
Public Sub GraphiqueFrontiere(ByVal ws As Worksheet, ByVal premiereLigne As Long, ByVal nbPoints As Long, _
                              ByVal nbDomines As Long, ByRef p As TParametres, ByVal nbAleatoires As Long)
    Dim ch As Chart, s As Series, wsStats As Worksheet, wsAlea As Worksheet
    Dim ligneMvp As Long, i As Long

    On Error GoTo Echec
    Set ch = NouveauGraphique(ws, ws.Cells(3, 9 + p.NbTitres).Left, ws.Range("A3").Top, 620, 400)
    ligneMvp = premiereLigne + nbDomines           ' le premier point efficient est la variance minimale

    ' 1) Nuage de portefeuilles aleatoires (en fond)
    If nbAleatoires > 0 Then
        Set wsAlea = FeuilleExistante(FEUILLE_ALEATOIRES)
        Set s = AjouterSerie(ch, "Portefeuilles possibles (aleatoires)", wsAlea.Range("A2").Resize(nbAleatoires, 1), _
                             wsAlea.Range("B2").Resize(nbAleatoires, 1), xlXYScatter)
        StyleMarqueur s, RGB(200, 200, 200), 2
    End If

    ' 2) Partie dominee (en pointilles), jusqu'au portefeuille de variance minimale
    If nbDomines > 0 Then
        Set s = AjouterSerie(ch, "Portefeuilles domines", ws.Cells(premiereLigne, 5).Resize(nbDomines + 1, 1), _
                             ws.Cells(premiereLigne, 3).Resize(nbDomines + 1, 1), xlXYScatterLines)
        StyleLigne s, RGB(150, 150, 150), True
    End If

    ' 3) Frontiere efficiente
    Set s = AjouterSerie(ch, "Frontiere efficiente", ws.Cells(ligneMvp, 5).Resize(nbPoints - nbDomines, 1), _
                         ws.Cells(ligneMvp, 3).Resize(nbPoints - nbDomines, 1), xlXYScatterLines)
    StyleLigne s, RGB(0, 70, 160), False

    ' 4) Les actions seules, avec leur code
    Set wsStats = FeuilleExistante(FEUILLE_STATS)
    Set s = AjouterSerie(ch, "Actions", wsStats.Cells(STAT_LIGNE_ECART_TYPE, 2).Resize(1, p.NbTitres), _
                         wsStats.Cells(STAT_LIGNE_MOYENNE, 2).Resize(1, p.NbTitres), xlXYScatter)
    StyleMarqueur s, RGB(220, 90, 0), 7
    EtiqueterPoints s, p.Codes

    ' 5) Portefeuille de variance minimale
    Set s = AjouterSerie(ch, "Variance minimale", ws.Cells(ligneMvp, 5), ws.Cells(ligneMvp, 3), xlXYScatter)
    StyleMarqueur s, RGB(0, 150, 60), 9

    MettreEnForme ch, "Frontiere efficiente (" & p.NbTitres & " actions)", True
    Exit Sub
Echec:
    NoterEchec ws.Name, Err.Description
End Sub

'------------------------------------------------------------------------------
' Outils
'------------------------------------------------------------------------------

Private Sub NoterEchec(ByVal nomFeuille As String, ByVal description As String)
    AvertissementGraphique = AvertissementGraphique & vbLf & "- feuille " & nomFeuille & " : " & description
End Sub

' Affiche le code de chaque action a cote de son point (purement decoratif).
Private Sub EtiqueterPoints(ByVal s As Series, ByVal codes As Variant)
    Dim i As Long
    On Error Resume Next
    For i = LBound(codes) To UBound(codes)
        s.Points(i).HasDataLabel = True
        s.Points(i).DataLabel.Text = codes(i)
    Next i
    On Error GoTo 0
End Sub

Private Function NouveauGraphique(ByVal ws As Worksheet, ByVal gauche As Double, ByVal haut As Double, _
                                  ByVal largeur As Double, ByVal hauteur As Double) As Chart
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(gauche, haut, largeur, hauteur)
    co.Chart.ChartType = xlXYScatter
    Do While co.Chart.SeriesCollection.Count > 0   ' Excel ajoute parfois des series automatiquement
        co.Chart.SeriesCollection(1).Delete
    Loop
    Set NouveauGraphique = co.Chart
End Function

Private Function AjouterSerie(ByVal ch As Chart, ByVal nom As String, ByVal x As Range, ByVal y As Range, _
                              ByVal typeSerie As XlChartType) As Series
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nom
    s.XValues = x
    s.Values = y
    s.ChartType = typeSerie
    Set AjouterSerie = s
End Function

Private Sub MettreEnForme(ByVal ch As Chart, ByVal titreGraphique As String, ByVal avecLegende As Boolean)
    On Error Resume Next                      ' mise en forme : ne doit jamais bloquer la macro
    ch.HasTitle = True
    ch.ChartTitle.Text = titreGraphique
    ch.HasLegend = avecLegende
    If avecLegende Then ch.Legend.Position = xlLegendPositionBottom
    With ch.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Risque : ecart-type journalier sigma(pf)"
        .TickLabels.NumberFormat = "0.0%"
        .HasMajorGridlines = False
    End With
    With ch.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Rendement journalier moyen R(pf)"
        .TickLabels.NumberFormat = "0.00%"
    End With
    On Error GoTo 0
End Sub

Private Sub StyleMarqueur(ByVal s As Series, ByVal couleur As Long, ByVal taille As Long)
    On Error Resume Next
    s.MarkerStyle = xlMarkerStyleCircle
    s.MarkerSize = taille
    s.MarkerBackgroundColor = couleur
    s.MarkerForegroundColor = couleur
    On Error GoTo 0
End Sub

Private Sub StyleLigne(ByVal s As Series, ByVal couleur As Long, ByVal pointilles As Boolean)
    On Error Resume Next
    s.Format.Line.ForeColor.RGB = couleur
    s.Format.Line.Weight = 2.25
    If pointilles Then s.Format.Line.DashStyle = msoLineDash
    s.MarkerStyle = xlMarkerStyleCircle
    s.MarkerSize = 4
    s.MarkerBackgroundColor = couleur
    s.MarkerForegroundColor = couleur
    On Error GoTo 0
End Sub
