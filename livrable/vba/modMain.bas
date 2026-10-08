Attribute VB_Name = "modMain"
Option Explicit
'==============================================================================
' modMain : macros lancees par les boutons de la feuille Parametres
'
'   LancerAnalyse    : lit les fichiers CSV, calcule tout et cree les feuilles
'   EffacerResultats : supprime les feuilles de resultats
'   CreerBoutons     : (a lancer une fois) ajoute les boutons sur Parametres
'==============================================================================

Public Sub LancerAnalyse()
    Dim p As TParametres
    Dim dates() As Date, prix() As Double, dividendes() As Double, nbDatesIgnorees As Long
    Dim rend() As Double, moy() As Double, cov() As Double
    Dim poids() As Double, rendFrontiere() As Double, sigmaFrontiere() As Double, efficient() As Boolean
    Dim wMvp() As Double, nbDomines As Long
    Dim rendAlea() As Double, sigmaAlea() As Double
    Dim debut As Double, message As String

    On Error GoTo GestionErreur
    debut = Timer
    Application.ScreenUpdating = False
    Application.StatusBar = "Lecture des parametres..."

    ' 1) Parametres et donnees
    AvertissementGraphique = ""
    p = LireParametres()
    ImporterDonnees p, dates, prix, dividendes, nbDatesIgnorees

    ' 2) Statistiques : rendements, moyennes, variances, covariances
    Application.StatusBar = "Calcul des statistiques..."
    rend = CalculerRendements(prix, dividendes)
    moy = CalculerMoyennes(rend)
    cov = CalculerCovariances(rend, moy)

    ' 3) Frontiere efficiente et portefeuilles aleatoires
    CalculerFrontiere cov, moy, p.NbPointsFrontiere, poids, rendFrontiere, sigmaFrontiere, efficient, wMvp, nbDomines
    If p.NbAleatoires > 0 Then
        Application.StatusBar = "Tirage des portefeuilles aleatoires..."
        GenererPortefeuillesAleatoires cov, moy, p.NbAleatoires, rendAlea, sigmaAlea
    End If

    ' 4) Ecriture des resultats
    Application.StatusBar = "Ecriture des resultats..."
    EcrireCours p, dates, prix
    EcrireRendements p, dates, rend
    EcrireStatistiques p, rend, moy, cov
    EcrireDeuxTitres p, moy, cov
    If p.NbAleatoires > 0 Then
        EcrireAleatoires rendAlea, sigmaAlea
    Else
        SupprimerFeuille FEUILLE_ALEATOIRES
    End If
    EcrireFrontiere p, poids, rendFrontiere, sigmaFrontiere, efficient, wMvp, moy, cov, nbDomines, p.NbAleatoires
    EcrireCalculateur p

    Application.StatusBar = False
    Application.ScreenUpdating = True
    FeuilleExistante(FEUILLE_FRONTIERE).Activate

    message = "Analyse terminee en " & Format$(Timer - debut, "0.0") & " s." & vbLf & vbLf & _
              p.NbTitres & " actions, " & UBound(rend, 1) & " rendements journaliers" & vbLf & _
              "du " & Format$(dates(1), "dd/mm/yyyy") & " au " & Format$(dates(UBound(dates)), "dd/mm/yyyy") & "." & vbLf & vbLf & _
              "Portefeuille de variance minimale :" & vbLf & _
              "  R(pf) = " & Format$(RendementPortefeuille(wMvp, moy), "0.0000%") & _
              "   sigma(pf) = " & Format$(EcartTypePortefeuille(wMvp, cov), "0.0000%")
    If nbDatesIgnorees > 0 Then
        message = message & vbLf & vbLf & "Remarque : " & nbDatesIgnorees & _
                  " date(s) absente(s) d'au moins un fichier ont ete ignorees."
    End If
    If AvertissementGraphique <> "" Then
        message = message & vbLf & vbLf & "Attention, graphique(s) non cree(s) :" & AvertissementGraphique
    End If
    MsgBox message, vbInformation, "Frontiere efficiente"
    Exit Sub

GestionErreur:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Err.Number = ERR_UTILISATEUR Then
        MsgBox Err.Description, vbExclamation, "Frontiere efficiente - probleme a corriger"
    Else
        MsgBox "Erreur inattendue n" & Chr$(176) & " " & Err.Number & " :" & vbLf & Err.Description & vbLf & vbLf & _
               "Verifiez les fichiers CSV et les parametres, puis relancez.", vbCritical, "Frontiere efficiente"
    End If
End Sub

' Supprime toutes les feuilles de resultats (les parametres sont conserves).
Public Sub EffacerResultats()
    Dim noms As Variant, i As Long
    If MsgBox("Supprimer toutes les feuilles de resultats ?" & vbLf & _
              "(Les feuilles Parametres et Mode d'emploi sont conservees.)", _
              vbQuestion + vbYesNo, "Frontiere efficiente") = vbNo Then Exit Sub
    noms = Array(FEUILLE_COURS, FEUILLE_RENDEMENTS, FEUILLE_STATS, FEUILLE_DEUX_TITRES, _
                 FEUILLE_FRONTIERE, FEUILLE_ALEATOIRES, FEUILLE_CALCULATEUR)
    For i = LBound(noms) To UBound(noms)
        SupprimerFeuille CStr(noms(i))
    Next i
    FeuilleExistante(FEUILLE_PARAMETRES).Activate
End Sub

' Ajoute les boutons "Lancer l'analyse" et "Effacer les resultats" sur la feuille Parametres.
' A lancer une seule fois apres l'import des modules (Outils > Macro > Macros > CreerBoutons).
Public Sub CreerBoutons()
    Dim ws As Worksheet
    Set ws = FeuilleExistante(FEUILLE_PARAMETRES)
    If ws Is Nothing Then
        MsgBox "La feuille '" & FEUILLE_PARAMETRES & "' est introuvable.", vbExclamation
        Exit Sub
    End If
    SupprimerBouton ws, "btnLancer"
    SupprimerBouton ws, "btnEffacer"
    AjouterBouton ws, "btnLancer", "Lancer l'analyse", "LancerAnalyse", ws.Range("F4")
    AjouterBouton ws, "btnEffacer", "Effacer les resultats", "EffacerResultats", ws.Range("F7")
    MsgBox "Boutons crees sur la feuille " & FEUILLE_PARAMETRES & ".", vbInformation
End Sub

'------------------------------------------------------------------------------
' Outils
'------------------------------------------------------------------------------

Private Sub AjouterBouton(ByVal ws As Worksheet, ByVal nom As String, ByVal texte As String, _
                          ByVal macro As String, ByVal position As Range)
    Dim b As Object
    Set b = ws.Buttons.Add(position.Left, position.Top, 170, 32)   ' bouton de formulaire (compatible Mac)
    b.Name = nom
    b.Caption = texte
    b.OnAction = macro
End Sub

Private Sub SupprimerBouton(ByVal ws As Worksheet, ByVal nom As String)
    On Error Resume Next
    ws.Buttons(nom).Delete
    On Error GoTo 0
End Sub

Public Sub SupprimerFeuille(ByVal nom As String)
    Dim ws As Worksheet
    Set ws = FeuilleExistante(nom)
    If ws Is Nothing Then Exit Sub
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub
