Attribute VB_Name = "Module2_Donnees"
Option Explicit
'==============================================================================
'  MODULE 2 : LES DONNEES (etape 1 de la consigne)
'==============================================================================
'  Pour chaque action, on recupere l'historique des cours journaliers :
'    - soit en le telechargeant sur le site d'Euronext (source "Euronext"),
'    - soit en lisant un fichier CODE.csv place a cote du classeur (source
'      "Fichiers CSV"), au meme format que l'export Euronext.
'
'  Format d'un fichier Euronext (separateur ";" et point decimal) :
'      "Historical Data"                      <- 3 lignes d'introduction
'      "From 2025-01-01 to 2025-12-31"
'      FR0000121014
'      Date;Open;High;Low;Last;Close;...      <- ligne d'en-tete
'      31/12/2025;636.90;645.00;635.70;645.00;645.00;...   <- du plus recent
'      ...                                                  au plus ancien
'  On garde la colonne 1 (Date) et la colonne 6 (Close = cours de cloture).
'
'  Euronext ne donne pas les dividendes : on les lit dans la feuille Dividendes.
'==============================================================================

' Sur Mac, Excel ne sait pas telecharger une page web tout seul. On utilise
' donc l'outil "curl" du Mac, grace a 4 fonctions du systeme (bibliotheque C).
' Ces lignes ne servent que sur Mac ; sur Windows elles sont ignorees.
#If Mac Then
Private Declare PtrSafe Function popen Lib "/usr/lib/libc.dylib" (ByVal commande As String, ByVal mode As String) As LongPtr
Private Declare PtrSafe Function pclose Lib "/usr/lib/libc.dylib" (ByVal fichier As LongPtr) As Long
Private Declare PtrSafe Function fread Lib "/usr/lib/libc.dylib" (ByVal tampon As String, ByVal taille As LongPtr, ByVal nombre As LongPtr, ByVal fichier As LongPtr) As Long
Private Declare PtrSafe Function feof Lib "/usr/lib/libc.dylib" (ByVal fichier As LongPtr) As LongPtr
#End If


'==============================================================================
' Recupere les cours de toutes les actions et remplit Dates() et Prix()
'==============================================================================
Public Sub RecupererCours()
    Dim i As Long, texte As String

    For i = 1 To NbActions
        If SourceEuronext Then
            Application.StatusBar = "Telechargement de " & Codes(i) & " sur Euronext (" & i & "/" & NbActions & ")..."
            texte = TelechargerDepuisEuronext(i)
        Else
            Application.StatusBar = "Lecture du fichier " & Codes(i) & ".csv..."
            texte = LireFichier(ThisWorkbook.Path & Application.PathSeparator & Codes(i) & ".csv")
        End If
        LireCoursEuronext texte, i
    Next i

    ' Les dividendes sont remis a zero : ils seront ajoutes par AjouterDividendes
    ReDim Dividendes(1 To NbJours, 1 To NbActions)
End Sub


'==============================================================================
' Lit le texte d'un fichier Euronext et range les cours de l'action i
'==============================================================================
Private Sub LireCoursEuronext(ByVal texte As String, ByVal i As Long)
    Dim lignes() As String, champs() As String
    Dim k As Long, premiereLigne As Long, nbLignes As Long, t As Long
    Dim jour As Date

    ' 1) On decoupe le texte en lignes
    texte = Replace(texte, vbCr, "")                  ' on enleve les retours chariot (Windows)
    lignes = Split(texte, vbLf)

    ' 2) On cherche la ligne d'en-tete "Date;Open;..." : les cours commencent juste apres
    premiereLigne = -1
    For k = 0 To UBound(lignes)
        If Left$(lignes(k), 5) = "Date;" Then
            premiereLigne = k + 1
            Exit For
        End If
    Next k
    If premiereLigne = -1 Then
        Signaler "Donnees de " & Codes(i) & " illisibles : la ligne 'Date;Open;...;Close' est introuvable."
    End If

    ' 3) On compte les lignes de cours (on ignore les lignes vides)
    nbLignes = 0
    For k = premiereLigne To UBound(lignes)
        If InStr(lignes(k), ";") > 0 Then nbLignes = nbLignes + 1
    Next k
    If nbLignes < 3 Then
        Signaler "Aucun cours pour " & Codes(i) & " sur la periode choisie : verifiez le code ISIN et les dates."
    End If

    ' 4) La premiere action fixe le calendrier ; les suivantes doivent avoir les memes dates
    If i = 1 Then
        NbJours = nbLignes
        ReDim Dates(1 To NbJours)
        ReDim Prix(1 To NbJours, 1 To NbActions)
    ElseIf nbLignes <> NbJours Then
        Signaler Codes(i) & " a " & nbLignes & " jours de cotation au lieu de " & NbJours & _
                 " pour " & Codes(1) & " : choisissez des actions du meme marche."
    End If

    ' 5) On range les cours. Le fichier va du plus recent au plus ancien :
    '    la premiere ligne lue correspond donc au dernier jour (t = NbJours).
    t = NbJours
    For k = premiereLigne To UBound(lignes)
        If InStr(lignes(k), ";") > 0 Then
            champs = Split(lignes(k), ";")
            ' La date est au format JJ/MM/AAAA
            jour = DateSerial(Val(Mid$(champs(0), 7, 4)), Val(Mid$(champs(0), 4, 2)), Val(Left$(champs(0), 2)))
            If i = 1 Then
                Dates(t) = jour
            ElseIf Dates(t) <> jour Then
                Signaler "Les dates de " & Codes(i) & " ne correspondent pas a celles de " & Codes(1) & "."
            End If
            ' Colonne 6 = Close. Val() lit toujours le point comme separateur decimal.
            Prix(t, i) = Val(champs(5))
            If Prix(t, i) <= 0 Then Signaler "Cours manquant pour " & Codes(i) & " le " & champs(0) & "."
            t = t - 1
        End If
    Next k
End Sub


'==============================================================================
' Ajoute les dividendes de la feuille Dividendes (Euronext ne les fournit pas)
'==============================================================================
Public Sub AjouterDividendes()
    Dim ws As Worksheet, ligne As Long, i As Long, t As Long
    Dim code As String, jourDetachement As Date

    If Not FeuilleExiste("Dividendes") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Dividendes")

    ligne = 4                                          ' les dividendes commencent en ligne 4
    Do While Trim$(ws.Cells(ligne, 1).Value) <> ""
        code = UCase$(Trim$(ws.Cells(ligne, 1).Value))
        If Not IsDate(ws.Cells(ligne, 2).Value) Or Not IsNumeric(ws.Cells(ligne, 3).Value) Then
            Signaler "Feuille Dividendes, ligne " & ligne & " : date ou montant invalide."
        End If
        jourDetachement = ws.Cells(ligne, 2).Value

        ' On cherche l'action correspondante, puis le jour de detachement dans le calendrier
        For i = 1 To NbActions
            If Codes(i) = code Then
                For t = 2 To NbJours
                    If Dates(t) >= jourDetachement And Dates(t - 1) < jourDetachement Then
                        Dividendes(t, i) = Dividendes(t, i) + ws.Cells(ligne, 3).Value
                    End If
                Next t
            End If
        Next i
        ligne = ligne + 1
    Loop
End Sub


'==============================================================================
' Telechargement sur Euronext
'==============================================================================
' C'est l'adresse qu'utilise le bouton "Telecharger" de la page d'une action
' sur live.euronext.com. adjusted=N : cours reellement cotes.
Private Function TelechargerDepuisEuronext(ByVal i As Long) As String
    Dim adresse As String, texte As String
    adresse = "https://live.euronext.com/en/ajax/AwlHistoricalPrice/getFullDownloadAjax/" & _
              Isins(i) & "-XPAR?format=csv&decimal_separator=.&date_form=d/m/Y&adjusted=N" & _
              "&startdate=" & Format$(DateDebut, "yyyy-mm-dd") & "&enddate=" & Format$(DateFin, "yyyy-mm-dd")
    texte = TelechargerTexte(adresse)
    If InStr(texte, "Date;") = 0 Then
        Signaler "Le telechargement de " & Codes(i) & " sur Euronext a echoue." & vbLf & _
                 "Reponse : " & Left$(texte, 150) & vbLf & vbLf & _
                 "Verifiez la connexion Internet et l'ISIN, ou passez en source 'Fichiers CSV'."
    End If
    TelechargerDepuisEuronext = texte
End Function

' Telecharge le contenu d'une adresse web.
Public Function TelechargerTexte(ByVal adresse As String) As String
#If Mac Then
    ' --- Sur Mac : on lance "curl adresse" et on lit ce qu'il renvoie, morceau par morceau
    Dim fichier As LongPtr, morceau As String, nbLus As Long, resultat As String
    fichier = popen("/usr/bin/curl --silent --show-error --max-time 60 '" & adresse & "' 2>&1", "r")
    If fichier = 0 Then Signaler "Impossible de lancer curl : passez en source 'Fichiers CSV'."
    Do While feof(fichier) = 0
        morceau = Space$(4096)
        nbLus = fread(morceau, 1, 4095, fichier)
        If nbLus > 0 Then resultat = resultat & Left$(morceau, nbLus)
    Loop
    pclose fichier
    TelechargerTexte = resultat
#Else
    ' --- Sur Windows : objet de requete web fourni par le systeme
    Dim requete As Object
    Set requete = CreateObject("MSXML2.XMLHTTP")
    requete.Open "GET", adresse, False
    requete.send
    TelechargerTexte = requete.responseText
#End If
End Function

' Lit tout le contenu d'un fichier texte (source "Fichiers CSV").
Private Function LireFichier(ByVal chemin As String) As String
    Dim numero As Integer, ouvert As Boolean
#If Mac Then
    GrantAccessToMultipleFiles Array(chemin)           ' Mac : demande l'autorisation de lire le fichier
#End If
    numero = FreeFile
    On Error Resume Next                               ' on teste si le fichier s'ouvre...
    Open chemin For Input As #numero
    ouvert = (Err.Number = 0)
    On Error GoTo 0
    If Not ouvert Then Signaler "Fichier introuvable : " & chemin   ' ... sinon message clair
    LireFichier = Input$(LOF(numero), #numero)
    Close #numero
End Function
