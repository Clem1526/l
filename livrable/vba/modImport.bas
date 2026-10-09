Attribute VB_Name = "modImport"
Option Explicit
'==============================================================================
' modImport : recuperation des cours (Euronext ou fichiers CSV), dividendes,
'             alignement des dates
'
' Formats acceptes (Euronext ou un fichier par action, nomme CODE.csv) :
'   - export Euronext : 3 lignes d'introduction puis Date;Open;High;Low;Last;Close;...
'   - Date,Close,AdjClose,Dividend   (fichiers issus de Yahoo Finance)
'   - Date,Open,High,Low,Close,Adj Close,Volume   (export Yahoo Finance)
'   - Date;Close   (export Google Sheets en francais, virgule decimale)
' Les dates peuvent etre AAAA-MM-JJ ou JJ/MM/AAAA, en ordre croissant ou decroissant.
' Si le fichier n'a pas de colonne de dividendes, ceux de la feuille Dividendes sont ajoutes.
'==============================================================================

' Recupere les cours de toutes les actions retenues et renvoie des tableaux
' alignes sur les dates communes a toutes les actions :
'   dates(1 To T), prix(1 To T, 1 To N), dividendes(1 To T, 1 To N)
Public Sub ImporterDonnees(ByRef p As TParametres, ByRef dates() As Date, ByRef prix() As Double, _
                           ByRef dividendes() As Double, ByRef nbDatesIgnorees As Long)
    Dim i As Long, t As Long, k As Long, nbDates As Long
    Dim chemins() As String, texte As String, libelle As String, aDesDividendes As Boolean
    Dim datesTitres() As Variant, prixTitres() As Variant, divTitres() As Variant
    Dim d() As Date, c() As Double, dv() As Double
    Dim communes() As Date, autres() As Date
    Dim codesDiv() As String, datesDiv() As Date, montantsDiv() As Double, nbDiv As Long

    ReDim chemins(1 To p.NbTitres)
    ReDim datesTitres(1 To p.NbTitres)
    ReDim prixTitres(1 To p.NbTitres)
    ReDim divTitres(1 To p.NbTitres)

    If p.Source = SOURCE_CSV Then
        For i = 1 To p.NbTitres
            chemins(i) = CheminFichier(p.Dossier, p.Codes(i))
        Next i
        DemanderAccesFichiers chemins
    End If
    LireTableDividendes codesDiv, datesDiv, montantsDiv, nbDiv

    ' 1) Recuperation et lecture des cours de chaque action
    For i = 1 To p.NbTitres
        If p.Source = SOURCE_EURONEXT Then
            Application.StatusBar = "Telechargement des cours de " & p.Codes(i) & " sur Euronext (" & i & "/" & p.NbTitres & ")..."
            texte = TelechargerCoursEuronext(p.Codes(i), p.Isins(i), p.Marches(i), p.DateDebut, p.DateFin)
            libelle = "Euronext, action " & p.Codes(i) & " (ISIN " & p.Isins(i) & ")"
        Else
            Application.StatusBar = "Lecture du fichier " & p.Codes(i) & ".csv ..."
            texte = LireTexte(chemins(i))
            libelle = chemins(i)
        End If
        AnalyserTexteCours texte, libelle, d, c, dv, aDesDividendes
        GarderPeriode d, c, dv, p.DateDebut, p.DateFin, libelle
        If Not aDesDividendes Then AjouterDividendes p.Codes(i), d, dv, codesDiv, datesDiv, montantsDiv, nbDiv
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

' Analyse le contenu d'un fichier de cours et renvoie ses dates, cours de cloture
' et dividendes (ordre croissant). aDesDividendes = le fichier contient des dividendes
' (colonne Dividend) ou un cours ajuste qui les inclut deja.
Public Sub AnalyserTexteCours(ByVal contenu As String, ByVal libelle As String, ByRef dates() As Date, _
                              ByRef cours() As Double, ByRef dividendes() As Double, ByRef aDesDividendes As Boolean)
    Dim sep As String, lignes() As String, champs() As String, entete() As String
    Dim colDate As Long, colCours As Long, colDiv As Long, colMax As Long, ligneEntete As Long
    Dim i As Long, n As Long
    Dim d As Date, prix As Double, dv As Double
    Dim okDate As Boolean, okPrix As Boolean, okDiv As Boolean

    If Left$(contenu, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then contenu = Mid$(contenu, 4) ' BOM UTF-8
    contenu = Replace(contenu, vbCr, "")    ' fins de ligne Windows / Mac / Unix
    lignes = Split(contenu, vbLf)

    ' La ligne d'en-tete est la premiere qui contient "date" et un nom de colonne de cours
    ligneEntete = -1
    For i = 0 To UBound(lignes)
        If i > 20 Then Exit For
        If InStr(LCase$(lignes(i)), "date") > 0 And (InStr(LCase$(lignes(i)), "close") > 0 _
           Or InStr(LCase$(lignes(i)), "clot") > 0 Or InStr(LCase$(lignes(i)), "adj") > 0) Then
            ligneEntete = i
            Exit For
        End If
    Next i
    If ligneEntete < 0 Or UBound(lignes) < ligneEntete + 3 Then
        ErreurUtilisateur "Donnees vides ou en-tete non reconnu :" & vbLf & libelle & vbLf & _
                          "La ligne d'en-tete doit contenir au moins les colonnes 'Date' et 'Close'."
    End If

    If InStr(lignes(ligneEntete), ";") > 0 Then sep = ";" Else sep = ","
    entete = Split(lignes(ligneEntete), sep)
    TrouverColonnes entete, libelle, colDate, colCours, colDiv, aDesDividendes
    colMax = colDate
    If colCours > colMax Then colMax = colCours
    If colDiv > colMax Then colMax = colDiv

    ReDim dates(1 To UBound(lignes))
    ReDim cours(1 To UBound(lignes))
    ReDim dividendes(1 To UBound(lignes))
    For i = ligneEntete + 1 To UBound(lignes)
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
                        ErreurUtilisateur "Cours nul ou negatif a la ligne " & (i + 1) & " :" & vbLf & libelle
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
        ErreurUtilisateur "Aucune donnee exploitable :" & vbLf & libelle & vbLf & _
                          "Format attendu : une colonne Date et une colonne Close."
    End If
    ReDim Preserve dates(1 To n)
    ReDim Preserve cours(1 To n)
    ReDim Preserve dividendes(1 To n)

    If dates(1) > dates(n) Then InverserOrdre dates, cours, dividendes   ' du plus recent au plus ancien
    For i = 2 To n
        If dates(i) <= dates(i - 1) Then
            ErreurUtilisateur "Dates non triees ou en double (" & Format$(dates(i), "dd/mm/yyyy") & ") :" & vbLf & libelle
        End If
    Next i
End Sub

' Ne garde que les cotations comprises entre debut et fin (inclus).
Private Sub GarderPeriode(ByRef dates() As Date, ByRef cours() As Double, ByRef dividendes() As Double, _
                          ByVal debut As Date, ByVal fin As Date, ByVal libelle As String)
    Dim i As Long, n As Long
    For i = 1 To UBound(dates)
        If dates(i) >= debut And dates(i) <= fin Then
            n = n + 1
            dates(n) = dates(i)
            cours(n) = cours(i)
            dividendes(n) = dividendes(i)
        End If
    Next i
    If n < 3 Then
        ErreurUtilisateur "Moins de 3 cotations entre le " & Format$(debut, "dd/mm/yyyy") & " et le " & _
                          Format$(fin, "dd/mm/yyyy") & " :" & vbLf & libelle & vbLf & _
                          "Verifiez la periode choisie dans la feuille Parametres."
    End If
    ReDim Preserve dates(1 To n)
    ReDim Preserve cours(1 To n)
    ReDim Preserve dividendes(1 To n)
End Sub

' Lit la feuille Dividendes (code, date de detachement, montant par action).
Private Sub LireTableDividendes(ByRef codes() As String, ByRef datesDiv() As Date, _
                                ByRef montants() As Double, ByRef nb As Long)
    Dim ws As Worksheet, ligne As Long, code As String
    nb = 0
    ReDim codes(1 To 1): ReDim datesDiv(1 To 1): ReDim montants(1 To 1)
    Set ws = FeuilleExistante(FEUILLE_DIVIDENDES)
    If ws Is Nothing Then Exit Sub
    ligne = LIGNE_PREMIER_DIVIDENDE
    Do While Trim$(CStr(ws.Cells(ligne, 1).Value)) <> ""
        code = Trim$(CStr(ws.Cells(ligne, 1).Value))
        If Not IsDate(ws.Cells(ligne, 2).Value) Then
            ErreurUtilisateur "Feuille Dividendes, cellule " & ws.Cells(ligne, 2).Address(False, False) & _
                              " : saisissez la date de detachement (ex. 24/04/2025)."
        End If
        If IsEmpty(ws.Cells(ligne, 3).Value) Or Not IsNumeric(ws.Cells(ligne, 3).Value) Then
            ErreurUtilisateur "Feuille Dividendes, cellule " & ws.Cells(ligne, 3).Address(False, False) & _
                              " : saisissez le montant du dividende par action (ex. 7,5)."
        End If
        nb = nb + 1
        ReDim Preserve codes(1 To nb): ReDim Preserve datesDiv(1 To nb): ReDim Preserve montants(1 To nb)
        codes(nb) = UCase$(code)
        datesDiv(nb) = CDate(ws.Cells(ligne, 2).Value)
        montants(nb) = CDbl(ws.Cells(ligne, 3).Value)
        ligne = ligne + 1
    Loop
End Sub

' Ajoute les dividendes d'une action au jour de detachement (ou au premier jour de cotation suivant).
Private Sub AjouterDividendes(ByVal code As String, ByRef dates() As Date, ByRef dividendes() As Double, _
                              ByRef codes() As String, ByRef datesDiv() As Date, ByRef montants() As Double, _
                              ByVal nb As Long)
    Dim k As Long, t As Long
    For k = 1 To nb
        If codes(k) = UCase$(code) And datesDiv(k) > dates(1) And datesDiv(k) <= dates(UBound(dates)) Then
            t = 2
            Do While dates(t) < datesDiv(k)
                t = t + 1
            Loop
            dividendes(t) = dividendes(t) + montants(k)
        End If
    Next k
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
' a defaut, cours "Close" seul (dividendes pris dans la feuille Dividendes).
Private Sub TrouverColonnes(ByRef entete() As String, ByVal libelle As String, ByRef colDate As Long, _
                            ByRef colCours As Long, ByRef colDiv As Long, ByRef aDesDividendes As Boolean)
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
        aDesDividendes = True       ' deja inclus dans le cours ajuste
    Else
        aDesDividendes = (colDiv >= 0)
    End If
    If colDate < 0 Or colCours < 0 Then
        ErreurUtilisateur "En-tete non reconnu :" & vbLf & libelle & vbLf & _
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
