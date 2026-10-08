Attribute VB_Name = "modConfig"
Option Explicit
'==============================================================================
' modConfig : constantes du classeur, lecture et controle des parametres
'==============================================================================

' --- Noms des feuilles ---
Public Const FEUILLE_PARAMETRES As String = "Parametres"
Public Const FEUILLE_COURS As String = "Cours"
Public Const FEUILLE_RENDEMENTS As String = "Rendements"
Public Const FEUILLE_STATS As String = "Statistiques"
Public Const FEUILLE_DEUX_TITRES As String = "Deux titres"
Public Const FEUILLE_FRONTIERE As String = "Frontiere"
Public Const FEUILLE_ALEATOIRES As String = "Aleatoires"
Public Const FEUILLE_CALCULATEUR As String = "Calculateur"

' --- Cellules de la feuille Parametres ---
Public Const CELLULE_DOSSIER As String = "C4"
Public Const CELLULE_NB_POINTS As String = "C5"
Public Const CELLULE_NB_ALEATOIRES As String = "C6"
Public Const CELLULE_PAS_GRILLE As String = "C7"
Public Const CELLULE_TITRE_A As String = "C8"
Public Const CELLULE_TITRE_B As String = "C9"

' Tableau des actions : colonne B = "Oui"/"Non", C = code (nom du fichier), D = nom
Public Const LIGNE_PREMIER_TITRE As Long = 13
Public Const COL_UTILISER As Long = 2
Public Const COL_CODE As Long = 3
Public Const COL_NOM As Long = 4
Public Const NB_MAX_TITRES As Long = 30

' Nombre de jours de bourse par an (pour les valeurs annualisees)
Public Const JOURS_PAR_AN As Long = 252

' Numero d'erreur reserve aux erreurs "utilisateur" (message clair, pas un bug)
Public Const ERR_UTILISATEUR As Long = -2147220504   ' = vbObjectError + 1000

' Ensemble des parametres saisis par l'utilisateur
Public Type TParametres
    Dossier As String              ' dossier contenant les fichiers CSV
    NbPointsFrontiere As Long      ' nombre de points sur la partie efficiente
    NbAleatoires As Long           ' nombre de portefeuilles aleatoires (0 = aucun)
    PasGrille As Double            ' pas de la grille du tableau a deux titres
    IndiceTitreA As Long           ' indices (dans Codes) des deux titres du tableau a 2 titres
    IndiceTitreB As Long
    NbTitres As Long
    Codes() As String              ' codes des actions retenues (1 To NbTitres)
    Noms() As String               ' noms des entreprises (1 To NbTitres)
End Type

' Declenche une erreur "utilisateur" : le message sera affiche tel quel.
Public Sub ErreurUtilisateur(ByVal message As String)
    Err.Raise ERR_UTILISATEUR, "Frontiere efficiente", message
End Sub

' Renvoie la feuille demandee, ou Nothing si elle n'existe pas.
Public Function FeuilleExistante(ByVal nom As String) As Worksheet
    On Error Resume Next
    Set FeuilleExistante = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
End Function

' Lit la feuille Parametres et verifie chaque saisie.
Public Function LireParametres() As TParametres
    Dim p As TParametres
    Dim ws As Worksheet
    Dim i As Long, j As Long, ligne As Long
    Dim utiliser As String, code As String, nom As String
    Dim codes() As String, noms() As String

    Set ws = FeuilleExistante(FEUILLE_PARAMETRES)
    If ws Is Nothing Then
        ErreurUtilisateur "La feuille '" & FEUILLE_PARAMETRES & "' est introuvable." & vbLf & _
                          "Ne la renommez pas et ne la supprimez pas."
    End If

    ' Dossier des fichiers : par defaut, celui du classeur
    p.Dossier = Trim$(CStr(ws.Range(CELLULE_DOSSIER).Value))
    If p.Dossier = "" Then
        If ThisWorkbook.Path = "" Then
            ErreurUtilisateur "Enregistrez d'abord le classeur (format .xlsm) dans le dossier " & _
                              "qui contient les fichiers CSV, puis relancez la macro."
        End If
        p.Dossier = ThisWorkbook.Path
    End If

    p.NbPointsFrontiere = LireEntier(ws.Range(CELLULE_NB_POINTS), "Nombre de points de la frontiere", 3, 200)
    p.NbAleatoires = LireEntier(ws.Range(CELLULE_NB_ALEATOIRES), "Nombre de portefeuilles aleatoires", 0, 20000)
    p.PasGrille = LirePas(ws.Range(CELLULE_PAS_GRILLE))

    ' Tableau des actions
    ReDim codes(1 To NB_MAX_TITRES)
    ReDim noms(1 To NB_MAX_TITRES)
    For i = 0 To NB_MAX_TITRES - 1
        ligne = LIGNE_PREMIER_TITRE + i
        utiliser = UCase$(Trim$(CStr(ws.Cells(ligne, COL_UTILISER).Value)))
        code = Trim$(CStr(ws.Cells(ligne, COL_CODE).Value))
        nom = Trim$(CStr(ws.Cells(ligne, COL_NOM).Value))
        If code <> "" And (utiliser = "OUI" Or utiliser = "O" Or utiliser = "X") Then
            For j = 1 To p.NbTitres
                If UCase$(codes(j)) = UCase$(code) Then
                    ErreurUtilisateur "L'action '" & code & "' est selectionnee deux fois " & _
                                      "dans la feuille Parametres (ligne " & ligne & ")."
                End If
            Next j
            p.NbTitres = p.NbTitres + 1
            codes(p.NbTitres) = code
            If nom = "" Then nom = code
            noms(p.NbTitres) = nom
        End If
    Next i

    If p.NbTitres < 2 Then
        ErreurUtilisateur "Selectionnez au moins 2 actions : ecrivez 'Oui' dans la colonne " & _
                          "'Utiliser ?' de la feuille Parametres, en face de chaque code."
    End If
    ReDim Preserve codes(1 To p.NbTitres)
    ReDim Preserve noms(1 To p.NbTitres)
    p.Codes = codes
    p.Noms = noms

    ' Les deux titres du tableau "Deux titres" (par defaut : les deux premiers)
    p.IndiceTitreA = IndiceDuTitre(p, CStr(ws.Range(CELLULE_TITRE_A).Value), 1)
    p.IndiceTitreB = IndiceDuTitre(p, CStr(ws.Range(CELLULE_TITRE_B).Value), 2)
    If p.IndiceTitreA = p.IndiceTitreB Then
        ErreurUtilisateur "Le titre A et le titre B du tableau a deux titres doivent etre differents."
    End If

    LireParametres = p
End Function

' Lit un nombre entier compris entre mini et maxi.
Private Function LireEntier(ByVal cellule As Range, ByVal libelle As String, _
                            ByVal mini As Long, ByVal maxi As Long) As Long
    Dim v As Variant
    v = cellule.Value
    If IsEmpty(v) Or Not IsNumeric(v) Then
        ErreurUtilisateur "Parametre '" & libelle & "' (cellule " & cellule.Address(False, False) & _
                          ") : saisissez un nombre entier entre " & mini & " et " & maxi & "."
    End If
    If CDbl(v) <> Int(CDbl(v)) Or CDbl(v) < mini Or CDbl(v) > maxi Then
        ErreurUtilisateur "Parametre '" & libelle & "' (cellule " & cellule.Address(False, False) & _
                          ") : saisissez un nombre entier entre " & mini & " et " & maxi & "."
    End If
    LireEntier = CLng(v)
End Function

' Lit le pas de la grille (ex. 10 %) : il doit diviser 100 %.
Private Function LirePas(ByVal cellule As Range) As Double
    Dim v As Variant, nbPas As Long
    v = cellule.Value
    If IsEmpty(v) Or Not IsNumeric(v) Then
        ErreurUtilisateur "Pas de la grille (cellule " & cellule.Address(False, False) & _
                          ") : saisissez un pourcentage, par exemple 10 %."
    End If
    If CDbl(v) >= 1 Then v = CDbl(v) / 100      ' l'utilisateur a tape 10 au lieu de 10 %
    If CDbl(v) < 0.01 Or CDbl(v) > 0.5 Then
        ErreurUtilisateur "Pas de la grille (cellule " & cellule.Address(False, False) & _
                          ") : choisissez une valeur entre 1 % et 50 %."
    End If
    nbPas = CLng(1 / CDbl(v))
    If Abs(nbPas * CDbl(v) - 1) > 0.000001 Then
        ErreurUtilisateur "Pas de la grille (cellule " & cellule.Address(False, False) & _
                          ") : le pas doit diviser 100 % (par exemple 5 %, 10 %, 20 % ou 25 %)."
    End If
    LirePas = 1 / nbPas
End Function

' Renvoie la position d'un code parmi les actions retenues (ou la valeur par defaut si vide).
Private Function IndiceDuTitre(ByRef p As TParametres, ByVal code As String, _
                               ByVal indiceParDefaut As Long) As Long
    Dim i As Long
    code = Trim$(code)
    If code = "" Then
        IndiceDuTitre = indiceParDefaut
        Exit Function
    End If
    For i = 1 To p.NbTitres
        If UCase$(p.Codes(i)) = UCase$(code) Then
            IndiceDuTitre = i
            Exit Function
        End If
    Next i
    ErreurUtilisateur "Le titre '" & code & "' choisi pour le tableau a deux titres ne fait pas " & _
                      "partie des actions selectionnees (colonne 'Utiliser ?' = Oui)."
End Function
