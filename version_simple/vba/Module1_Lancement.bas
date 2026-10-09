Attribute VB_Name = "Module1_Lancement"
Option Explicit
'==============================================================================
'  PROJET VBA - FRONTIERE EFFICIENTE (theorie du portefeuille de Markowitz)
'==============================================================================
'  MODULE 1 : LANCEMENT
'  C'est ici que tout commence : la macro "LancerAnalyse" est reliee au bouton
'  de la feuille Parametres. Elle appelle les autres modules dans l'ordre des
'  etapes de la consigne :
'
'     Etape 1  : recuperer les cours journaliers          -> Module2_Donnees
'     Etape 2  : rendement moyen, variance, ecart-type    -> Module3_Calculs
'     Etape 3  : covariances entre les actions            -> Module3_Calculs
'     Etape 4  : rendement et risque selon les poids      -> Module3_Calculs
'     Etape 5  : tableau et graphique de la frontiere     -> Module3 + Module4
'
'  Les donnees partagees entre les modules sont stockees dans les variables
'  "Public" ci-dessous : chaque module peut les lire et les remplir.
'==============================================================================

'------------------------------------------------------------------------------
' Les actions choisies (lues dans la feuille Parametres)
'------------------------------------------------------------------------------
Public NbActions As Long            ' nombre d'actions analysees
Public Codes() As String            ' code de chaque action (ex. "MC")
Public Noms() As String             ' nom de l'entreprise (ex. "LVMH")
Public Isins() As String            ' code ISIN, utilise par Euronext (ex. "FR0000121014")

'------------------------------------------------------------------------------
' Les donnees de marche (remplies par le Module 2)
'------------------------------------------------------------------------------
Public NbJours As Long              ' nombre de jours de cotation
Public Dates() As Date              ' Dates(t)          : date du jour t
Public Prix() As Double             ' Prix(t, i)        : cours de cloture de l'action i le jour t
Public Dividendes() As Double       ' Dividendes(t, i)  : dividende detache le jour t (souvent 0)

'------------------------------------------------------------------------------
' Les resultats statistiques (remplis par le Module 3)
'------------------------------------------------------------------------------
Public Rendements() As Double       ' Rendements(t, i) : rendement de l'action i le jour t
Public Moyennes() As Double         ' Moyennes(i)      : rendement moyen E(Ri)
Public Covariances() As Double      ' Covariances(i,j) : Cov(Ri, Rj) ; la diagonale = les variances

'------------------------------------------------------------------------------
' Tous les portefeuilles testes (remplis par le Module 3)
'------------------------------------------------------------------------------
Public NbPortefeuilles As Long
Public PoidsPortefeuilles() As Double   ' PoidsPortefeuilles(m, i) : poids de l'action i dans le portefeuille m
Public RendementsPf() As Double         ' RendementsPf(m) : rendement du portefeuille m
Public RisquesPf() As Double            ' RisquesPf(m)    : ecart-type du portefeuille m
Public IndiceVarianceMin As Long        ' numero du portefeuille le moins risque

'------------------------------------------------------------------------------
' Les reglages (lus dans la feuille Parametres)
'------------------------------------------------------------------------------
Public SourceEuronext As Boolean    ' True = telechargement Euronext, False = fichiers CSV
Public DateDebut As Date
Public DateFin As Date
Public PasPoids As Double           ' pas des poids testes (ex. 5 %)
Public NbTranches As Long           ' nombre de points de la frontiere
Public PasGrille As Double          ' pas du tableau a deux titres (ex. 10 %)
Public TitreA As Long               ' numeros des deux actions du tableau a deux titres
Public TitreB As Long

' Numero d'erreur choisi pour les erreurs "de saisie" (message clair pour l'utilisateur).
' (-2147220504 est la valeur de vbObjectError + 1000)
Public Const ERREUR_SAISIE As Long = -2147220504


'==============================================================================
' LA MACRO PRINCIPALE (bouton "Lancer l'analyse")
'==============================================================================
Public Sub LancerAnalyse()
    Dim debut As Double
    debut = Timer                                  ' pour afficher la duree a la fin

    ' Si une erreur survient, on saute a "GestionErreur" pour afficher un message clair
    On Error GoTo GestionErreur
    Application.ScreenUpdating = False             ' l'ecran ne clignote pas pendant le calcul

    ' --- Etape 0 : lire les reglages et la liste des actions
    LireParametres

    ' --- Etape 1 : recuperer les cours (Euronext ou fichiers) et les dividendes
    RecupererCours
    AjouterDividendes

    ' --- Etapes 2 et 3 : rendements, moyennes, variances, covariances
    Application.StatusBar = "Calcul des statistiques..."
    CalculerRendements
    CalculerMoyennes
    CalculerCovariances

    ' --- Etapes 4 et 5 : tester toutes les repartitions possibles des poids
    Application.StatusBar = "Test de tous les portefeuilles possibles..."
    TesterTousLesPortefeuilles

    ' --- Affichage des resultats dans les feuilles
    Application.StatusBar = "Ecriture des resultats..."
    AfficherCours
    AfficherStatistiques
    AfficherDeuxTitres
    AfficherFrontiere

    ' --- Fin : on remet Excel dans son etat normal et on affiche un resume
    Application.StatusBar = False
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("Frontiere").Activate
    MsgBox "Analyse terminee en " & Format$(Timer - debut, "0.0") & " secondes." & vbLf & vbLf & _
           NbActions & " actions, " & (NbJours - 1) & " rendements journaliers" & vbLf & _
           "du " & Format$(Dates(1), "dd/mm/yyyy") & " au " & Format$(Dates(NbJours), "dd/mm/yyyy") & "." & vbLf & _
           NbPortefeuilles & " portefeuilles testes." & vbLf & vbLf & _
           "Portefeuille le moins risque : R = " & Format$(RendementsPf(IndiceVarianceMin), "0.000%") & _
           "   sigma = " & Format$(RisquesPf(IndiceVarianceMin), "0.000%"), _
           vbInformation, "Frontiere efficiente"
    Exit Sub

GestionErreur:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Err.Number = ERREUR_SAISIE Then
        ' Erreur prevue (fichier absent, saisie invalide...) : on affiche notre message
        MsgBox Err.Description, vbExclamation, "Frontiere efficiente - a corriger"
    Else
        ' Erreur imprevue : on affiche le message d'Excel
        MsgBox "Erreur inattendue (n " & Err.Number & ") :" & vbLf & Err.Description, _
               vbCritical, "Frontiere efficiente"
    End If
End Sub


'==============================================================================
' LECTURE DE LA FEUILLE PARAMETRES
'==============================================================================
Private Sub LireParametres()
    Dim ws As Worksheet
    Dim ligne As Long, n As Long, k As Long
    Dim texteSource As String

    Set ws = ThisWorkbook.Worksheets("Parametres")

    ' --- Les reglages (colonne C)
    texteSource = UCase$(Trim$(ws.Range("C4").Value))
    SourceEuronext = (InStr(texteSource, "EURONEXT") > 0)
    If Not SourceEuronext And InStr(texteSource, "CSV") = 0 Then
        Signaler "Cellule C4 : choisissez la source 'Euronext' ou 'Fichiers CSV'."
    End If

    If Not IsDate(ws.Range("C5").Value) Or Not IsDate(ws.Range("C6").Value) Then
        Signaler "Cellules C5 et C6 : saisissez la date de debut et la date de fin (ex. 01/01/2025)."
    End If
    DateDebut = ws.Range("C5").Value
    DateFin = ws.Range("C6").Value
    If DateFin <= DateDebut Then Signaler "La date de fin (C6) doit etre apres la date de debut (C5)."
    If SourceEuronext And DateDebut < Date - 730 Then
        Signaler "Euronext ne donne que les 2 dernieres annees : la date de debut (C5) doit etre " & _
                 "apres le " & Format$(Date - 730, "dd/mm/yyyy") & "."
    End If

    PasPoids = LirePourcentage(ws.Range("C7"), "Pas des poids testes")
    If Not IsNumeric(ws.Range("C8").Value) Or IsEmpty(ws.Range("C8").Value) Then
        Signaler "Cellule C8 : saisissez le nombre de points de la frontiere (ex. 20)."
    End If
    NbTranches = ws.Range("C8").Value
    If NbTranches < 3 Or NbTranches > 100 Then Signaler "Cellule C8 : le nombre de points doit etre entre 3 et 100."
    PasGrille = LirePourcentage(ws.Range("C9"), "Pas du tableau a deux titres")

    ' --- La liste des actions (a partir de la ligne 15) : on garde celles marquees "Oui"
    NbActions = 0
    ReDim Codes(1 To 30): ReDim Noms(1 To 30): ReDim Isins(1 To 30)
    For ligne = 15 To 44
        If UCase$(Trim$(ws.Cells(ligne, 2).Value)) = "OUI" And Trim$(ws.Cells(ligne, 3).Value) <> "" Then
            NbActions = NbActions + 1
            Codes(NbActions) = UCase$(Trim$(ws.Cells(ligne, 3).Value))
            Noms(NbActions) = Trim$(ws.Cells(ligne, 4).Value)
            Isins(NbActions) = UCase$(Trim$(ws.Cells(ligne, 5).Value))
            If SourceEuronext And Len(Isins(NbActions)) <> 12 Then
                Signaler "Ligne " & ligne & " : le code ISIN de " & Codes(NbActions) & _
                         " doit compter 12 caracteres (ex. FR0000121014 pour LVMH)."
            End If
        End If
    Next ligne
    If NbActions < 2 Then Signaler "Choisissez au moins 2 actions (colonne 'Utiliser ?' = Oui)."

    ' --- Les deux titres du tableau "Deux titres" (par defaut : les deux premiers)
    TitreA = NumeroAction(ws.Range("C10").Value, 1)
    TitreB = NumeroAction(ws.Range("C11").Value, 2)
    If TitreA = TitreB Then Signaler "Cellules C10 et C11 : choisissez deux actions differentes."
End Sub

' Lit un pourcentage (ex. 5 %) et verifie qu'il "tombe juste" sur 100 %.
Private Function LirePourcentage(ByVal cellule As Range, ByVal nom As String) As Double
    Dim valeur As Double
    If Not IsNumeric(cellule.Value) Or IsEmpty(cellule.Value) Then
        Signaler nom & " (cellule " & cellule.Address(False, False) & ") : saisissez un pourcentage, ex. 5 %."
    End If
    valeur = cellule.Value
    If valeur >= 1 Then valeur = valeur / 100        ' l'utilisateur a tape 5 au lieu de 5 %
    If valeur < 0.01 Or valeur > 0.5 Or Abs(Round(1 / valeur) * valeur - 1) > 0.000001 Then
        Signaler nom & " (cellule " & cellule.Address(False, False) & ") : choisissez un pas qui " & _
                 "divise 100 %, par exemple 5 %, 10 % ou 20 %."
    End If
    LirePourcentage = 1 / Round(1 / valeur)           ' valeur exacte (ex. 0,05)
End Function

' Renvoie le numero de l'action dont le code est donne (ou la valeur par defaut si vide).
Private Function NumeroAction(ByVal code As String, ByVal parDefaut As Long) As Long
    Dim i As Long
    code = UCase$(Trim$(code))
    NumeroAction = parDefaut
    If code = "" Then Exit Function
    For i = 1 To NbActions
        If Codes(i) = code Then
            NumeroAction = i
            Exit Function
        End If
    Next i
    Signaler "Le titre '" & code & "' (cellules C10 / C11) ne fait pas partie des actions choisies."
End Function

' Declenche une erreur "de saisie" : le message sera affiche tel quel a l'utilisateur.
Public Sub Signaler(ByVal message As String)
    Err.Raise ERREUR_SAISIE, "Frontiere efficiente", message
End Sub


'==============================================================================
' LES AUTRES BOUTONS
'==============================================================================

' Supprime les feuilles de resultats (les feuilles de reglages sont conservees).
Public Sub EffacerResultats()
    Dim nom As Variant
    If MsgBox("Supprimer les feuilles de resultats ?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Application.DisplayAlerts = False                 ' pas de demande de confirmation par feuille
    For Each nom In Array("Cours", "Statistiques", "Deux titres", "Frontiere", "Portefeuilles")
        If FeuilleExiste(CStr(nom)) Then ThisWorkbook.Worksheets(CStr(nom)).Delete
    Next nom
    Application.DisplayAlerts = True
End Sub

' A lancer UNE fois apres avoir colle le code : cree les deux boutons sur la feuille Parametres.
Public Sub CreerBoutons()
    Dim ws As Worksheet, bouton As Object
    Set ws = ThisWorkbook.Worksheets("Parametres")
    On Error Resume Next                               ' si les boutons n'existent pas encore
    ws.Buttons("btnLancer").Delete
    ws.Buttons("btnEffacer").Delete
    On Error GoTo 0

    Set bouton = ws.Buttons.Add(ws.Range("G4").Left, ws.Range("G4").Top, 180, 34)
    bouton.Name = "btnLancer"
    bouton.Caption = "Lancer l'analyse"
    bouton.OnAction = "LancerAnalyse"

    Set bouton = ws.Buttons.Add(ws.Range("G7").Left, ws.Range("G7").Top, 180, 34)
    bouton.Name = "btnEffacer"
    bouton.Caption = "Effacer les resultats"
    bouton.OnAction = "EffacerResultats"
    MsgBox "Les boutons ont ete crees sur la feuille Parametres.", vbInformation
End Sub

' Vrai si une feuille portant ce nom existe dans le classeur.
Public Function FeuilleExiste(ByVal nom As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nom Then FeuilleExiste = True
    Next ws
End Function
