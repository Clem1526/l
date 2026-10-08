Attribute VB_Name = "modImport"
Option Explicit
'==============================================================================
' modImport : lecture des fichiers CSV de cours et alignement des dates
'
' Formats acceptes (un fichier par action, nomme CODE.csv) :
'   - Date,Close,AdjClose,Dividend   (fichiers fournis, issus de Yahoo Finance)
'   - Date,Open,High,Low,Close,Adj Close,Volume   (export Yahoo Finance)
'   - Date;Close   (export Google Sheets en francais, virgule decimale)
' Les dates peuvent etre AAAA-MM-JJ ou JJ/MM/AAAA, en ordre croissant ou decroissant.
'==============================================================================

' Lit les fichiers de toutes les actions retenues et renvoie des tableaux
' alignes sur les dates communes a toutes les actions :
'   dates(1 To T), prix(1 To T, 1 To N), dividendes(1 To T, 1 To N)
Public Sub ImporterDonnees(ByRef p As TParametres, ByRef dates() As Date, ByRef prix() As Double, _
                           ByRef dividendes() As Double, ByRef nbDatesIgnorees As Long)
    Dim i As Long, t As Long, k As Long, nbDates As Long
    Dim chemins() As String
    Dim datesTitres() As Variant, prixTitres() As Variant, divTitres() As Variant
    Dim d() As Date, c() As Double, dv() As Double
    Dim communes() As Date, autres() As Date

    ReDim chemins(1 To p.NbTitres)
    ReDim datesTitres(1 To p.NbTitres)
    ReDim prixTitres(1 To p.NbTitres)
    ReDim divTitres(1 To p.NbTitres)

    For i = 1 To p.NbTitres
        chemins(i) = CheminFichier(p.Dossier, p.Codes(i))
    Next i
    DemanderAccesFichiers chemins

    ' 1) Lecture de chaque fichier
    For i = 1 To p.NbTitres
        Application.StatusBar = "Lecture du fichier " & p.Codes(i) & ".csv ..."
        LireFichierCours chemins(i), d, c, dv
        datesTitres(i) = CopieDates(d)       ' copies explicites : chaque action garde ses propres tableaux
        prixTitres(i) = CopieNombres(c)
        divTitres(i) = CopieNombres(dv)
    Next i

    ' 2) Dates presentes dans TOUS les fichiers
    d = datesTitres(1)
    communes = CopieDates(d)
    nbDates = UBound(communes)
    For i = 2 To p.NbTitres
        autres = datesTitres(i)
        nbDates = IntersecterDates(communes, nbDates, autres)
    Next i
    If nbDates < 3 Then
        ErreurUtilisateur "Les fichiers n'ont pas assez de dates en commun (" & nbDates & ")." & vbLf & _
                          "Verifiez qu'ils couvrent tous la meme periode."
    End If

    ' 3) Construction des tableaux alignes
    ReDim dates(1 To nbDates)
    ReDim prix(1 To nbDates, 1 To p.NbTitres)
    ReDim dividendes(1 To nbDates, 1 To p.NbTitres)
    For t = 1 To nbDates
        dates(t) = communes(t)
    Next t
    nbDatesIgnorees = 0
    For i = 1 To p.NbTitres
        d = datesTitres(i)
        c = prixTitres(i)
        dv = divTitres(i)
        If UBound(d) - nbDates > nbDatesIgnorees Then nbDatesIgnorees = UBound(d) - nbDates
        k = 1
        For t = 1 To nbDates
            Do While d(k) < dates(t)        ' on saute les dates absentes des autres fichiers
                k = k + 1
            Loop
            prix(t, i) = c(k)
            dividendes(t, i) = dv(k)
        Next t
    Next i
End Sub

' Construit le chemin complet du fichier CSV d'une action.
Public Function CheminFichier(ByVal dossier As String, ByVal code As String) As String
    Dim sep As String
    sep = Application.PathSeparator
    If Right$(dossier, 1) = sep Then dossier = Left$(dossier, Len(dossier) - 1)
    CheminFichier = dossier & sep & code & ".csv"
End Function

' Sur Mac, Excel doit obtenir l'autorisation de lire des fichiers hors de son dossier.
Private Sub DemanderAccesFichiers(ByRef chemins() As String)
#If Mac Then
    Dim liste() As Variant, i As Long
    ReDim liste(0 To UBound(chemins) - 1)
    For i = 1 To UBound(chemins)
        liste(i - 1) = chemins(i)
    Next i
    If Not GrantAccessToMultipleFiles(liste) Then
        ErreurUtilisateur "Excel n'a pas recu l'autorisation de lire les fichiers CSV." & vbLf & _
                          "Relancez la macro et cliquez sur 'Autoriser l'acces'."
    End If
#End If
End Sub

' Lit un fichier CSV et renvoie ses dates, cours de cloture et dividendes (ordre croissant).
Public Sub LireFichierCours(ByVal chemin As String, ByRef dates() As Date, _
                            ByRef cours() As Double, ByRef dividendes() As Double)
    Dim contenu As String, sep As String
    Dim lignes() As String, champs() As String
    Dim colDate As Long, colCours As Long, colDiv As Long, colMax As Long
    Dim i As Long, n As Long
    Dim d As Date, prix As Double, dv As Double
    Dim okDate As Boolean, okPrix As Boolean, okDiv As Boolean

    contenu = LireTexte(chemin)
    If Left$(contenu, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then contenu = Mid$(contenu, 4) ' BOM UTF-8
    contenu = Replace(contenu, vbCr, "")    ' fins de ligne Windows / Mac / Unix
    lignes = Split(contenu, vbLf)
    If UBound(lignes) < 3 Then
        ErreurUtilisateur "Le fichier suivant est vide ou trop court :" & vbLf & chemin
    End If

    If InStr(lignes(0), ";") > 0 Then sep = ";" Else sep = ","
    TrouverColonnes Split(lignes(0), sep), chemin, colDate, colCours, colDiv
    colMax = colDate
    If colCours > colMax Then colMax = colCours
    If colDiv > colMax Then colMax = colDiv

    ReDim dates(1 To UBound(lignes))
    ReDim cours(1 To UBound(lignes))
    ReDim dividendes(1 To UBound(lignes))
    For i = 1 To UBound(lignes)
        If Len(Trim$(lignes(i))) > 0 Then
            champs = Split(lignes(i), sep)
            If UBound(champs) >= colMax Then
                d = ConvertirDate(champs(colDate), okDate)
                prix = ConvertirNombre(champs(colCours), sep, okPrix)
                dv = 0
                If colDiv >= 0 Then
                    dv = ConvertirNombre(champs(colDiv), sep, okDiv)
                    If Not okDiv Then dv = 0
                End If
                ' Les lignes sans cours (ex. "null") sont ignorees
                If okDate And okPrix Then
                    If prix <= 0 Then
                        ErreurUtilisateur "Cours nul ou negatif a la ligne " & (i + 1) & " du fichier :" & vbLf & chemin
                    End If
                    n = n + 1
                    dates(n) = d
                    cours(n) = prix
                    dividendes(n) = dv
                End If
            End If
        End If
    Next i

    If n < 3 Then
        ErreurUtilisateur "Aucune donnee exploitable dans le fichier :" & vbLf & chemin & vbLf & _
                          "Format attendu : une colonne Date et une colonne Close."
    End If
    ReDim Preserve dates(1 To n)
    ReDim Preserve cours(1 To n)
    ReDim Preserve dividendes(1 To n)

    If dates(1) > dates(n) Then InverserOrdre dates, cours, dividendes   ' fichier du plus recent au plus ancien
    For i = 2 To n
        If dates(i) <= dates(i - 1) Then
            ErreurUtilisateur "Dates non triees ou en double (" & Format$(dates(i), "dd/mm/yyyy") & _
                              ") dans le fichier :" & vbLf & chemin
        End If
    Next i
End Sub

' Lit tout le contenu d'un fichier texte.
Private Function LireTexte(ByVal chemin As String) As String
    Dim numero As Integer, ouvert As Boolean
    numero = FreeFile
    On Error Resume Next
    Open chemin For Input Access Read As #numero
    ouvert = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If Not ouvert Then
        ErreurUtilisateur "Impossible d'ouvrir le fichier :" & vbLf & chemin & vbLf & vbLf & _
                          "Verifiez que le fichier existe, qu'il porte exactement le nom du code " & _
                          "(ex. MC.csv) et qu'il se trouve dans le dossier indique dans la feuille Parametres."
    End If
    LireTexte = Input$(LOF(numero), #numero)
    Close #numero
End Function

' Repere les colonnes Date, cours de cloture et dividende dans la ligne d'en-tete.
' Priorite au cours "Close" + colonne "Dividend" (formule du cours) ; a defaut, cours ajuste ;
' a defaut, cours "Close" seul (dividendes ignores).
Private Sub TrouverColonnes(ByRef entete() As String, ByVal chemin As String, ByRef colDate As Long, _
                            ByRef colCours As Long, ByRef colDiv As Long)
    Dim i As Long, nom As String, colAjuste As Long
    colDate = -1: colCours = -1: colDiv = -1: colAjuste = -1
    For i = 0 To UBound(entete)
        nom = LCase$(Trim$(Replace(entete(i), """", "")))
        If InStr(nom, "date") > 0 And colDate < 0 Then
            colDate = i
        ElseIf InStr(nom, "adj") > 0 Or InStr(nom, "ajust") > 0 Then
            colAjuste = i
        ElseIf Left$(nom, 5) = "close" Or (Left$(nom, 2) = "cl" And InStr(nom, "ture") > 0) Then
            colCours = i
        ElseIf InStr(nom, "divid") > 0 Then
            colDiv = i
        End If
    Next i
    ' Sans colonne "Dividend", on prend le cours ajuste, qui inclut deja les dividendes
    If colAjuste >= 0 And (colCours < 0 Or colDiv < 0) Then
        colCours = colAjuste
        colDiv = -1
    End If
    If colDate < 0 Or colCours < 0 Then
        ErreurUtilisateur "En-tete non reconnu dans le fichier :" & vbLf & chemin & vbLf & _
                          "La premiere ligne doit contenir au moins les colonnes 'Date' et 'Close'."
    End If
End Sub

' Convertit "2025-01-31" ou "31/01/2025" (heure eventuelle ignoree) en date.
Private Function ConvertirDate(ByVal texte As String, ByRef ok As Boolean) As Date
    texte = Trim$(Replace(texte, """", ""))
    ok = True
    If texte Like "####-##-##*" Then
        ConvertirDate = DateSerial(CInt(Left$(texte, 4)), CInt(Mid$(texte, 6, 2)), CInt(Mid$(texte, 9, 2)))
    ElseIf texte Like "##/##/####*" Then
        ConvertirDate = DateSerial(CInt(Mid$(texte, 7, 4)), CInt(Mid$(texte, 4, 2)), CInt(Left$(texte, 2)))
    Else
        ok = False
    End If
End Function

' Convertit un texte en nombre sans dependre des reglages regionaux d'Excel.
Private Function ConvertirNombre(ByVal texte As String, ByVal sep As String, ByRef ok As Boolean) As Double
    Dim i As Long
    texte = Replace(Replace(Replace(texte, """", ""), " ", ""), Chr$(160), "")
    If sep = ";" Then texte = Replace(texte, ",", ".")   ' virgule decimale (format francais)
    ok = False
    If Len(texte) = 0 Then Exit Function
    For i = 1 To Len(texte)
        If InStr("0123456789.-+eE", Mid$(texte, i, 1)) = 0 Then Exit Function
    Next i
    If Not (texte Like "*#*") Then Exit Function
    ok = True
    ConvertirNombre = Val(texte)            ' Val lit toujours le point comme separateur decimal
End Function

' Remet les observations dans l'ordre chronologique.
Private Sub InverserOrdre(ByRef dates() As Date, ByRef cours() As Double, ByRef dividendes() As Double)
    Dim i As Long, n As Long, d As Date, x As Double
    n = UBound(dates)
    For i = 1 To n \ 2
        d = dates(i): dates(i) = dates(n + 1 - i): dates(n + 1 - i) = d
        x = cours(i): cours(i) = cours(n + 1 - i): cours(n + 1 - i) = x
        x = dividendes(i): dividendes(i) = dividendes(n + 1 - i): dividendes(n + 1 - i) = x
    Next i
End Sub

' Garde dans "communes" (nbCommunes premieres cases) les dates aussi presentes dans "autres".
' Les deux listes sont triees : on les parcourt en parallele. Renvoie le nouveau nombre de dates.
Private Function IntersecterDates(ByRef communes() As Date, ByVal nbCommunes As Long, _
                                  ByRef autres() As Date) As Long
    Dim i As Long, j As Long, n As Long
    i = 1: j = 1
    Do While i <= nbCommunes And j <= UBound(autres)
        If communes(i) = autres(j) Then
            n = n + 1
            communes(n) = communes(i)
            i = i + 1: j = j + 1
        ElseIf communes(i) < autres(j) Then
            i = i + 1
        Else
            j = j + 1
        End If
    Loop
    IntersecterDates = n
End Function

' Copie independante d'un tableau de dates.
Private Function CopieDates(ByRef source() As Date) As Variant
    Dim copie() As Date, i As Long
    ReDim copie(LBound(source) To UBound(source))
    For i = LBound(source) To UBound(source)
        copie(i) = source(i)
    Next i
    CopieDates = copie
End Function

' Copie independante d'un tableau de nombres.
Private Function CopieNombres(ByRef source() As Double) As Variant
    Dim copie() As Double, i As Long
    ReDim copie(LBound(source) To UBound(source))
    For i = LBound(source) To UBound(source)
        copie(i) = source(i)
    Next i
    CopieNombres = copie
End Function
