Attribute VB_Name = "modEuronext"
Option Explicit
'==============================================================================
' modEuronext : telechargement des cours historiques sur le site d'Euronext
'
' Euronext propose l'export CSV de l'historique d'une action (bouton
' "Telecharger" de la page de l'action sur live.euronext.com). La macro
' appelle directement la meme adresse :
'   https://live.euronext.com/en/ajax/AwlHistoricalPrice/getFullDownloadAjax/
'       <ISIN>-<MARCHE>?format=csv&decimal_separator=.&date_form=d/m/Y
'       &adjusted=N&startdate=AAAA-MM-JJ&enddate=AAAA-MM-JJ
' adjusted=N : cours reellement cotes (Euronext ne tient pas compte des
' dividendes, qui sont lus dans la feuille Dividendes).
' Limite d'Euronext : seules les deux dernieres annees sont disponibles.
'
' Sur Mac, Excel ne dispose pas d'objet HTTP : la macro utilise l'outil
' systeme "curl" (fonctions popen / fread de la bibliotheque C du Mac).
' Sur Windows, elle utilise l'objet MSXML2.XMLHTTP.
'==============================================================================

#If Mac Then
Private Declare PtrSafe Function c_popen Lib "/usr/lib/libc.dylib" Alias "popen" (ByVal commande As String, ByVal mode As String) As LongPtr
Private Declare PtrSafe Function c_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" (ByVal fichier As LongPtr) As Long
Private Declare PtrSafe Function c_fread Lib "/usr/lib/libc.dylib" Alias "fread" (ByVal tampon As String, ByVal taille As LongPtr, ByVal nombre As LongPtr, ByVal fichier As LongPtr) As Long
Private Declare PtrSafe Function c_feof Lib "/usr/lib/libc.dylib" Alias "feof" (ByVal fichier As LongPtr) As LongPtr
#End If

Private Const ADRESSE_EURONEXT As String = "https://live.euronext.com/en/ajax/AwlHistoricalPrice/getFullDownloadAjax/"

' Adresse de l'export CSV Euronext pour une action et une periode.
Public Function AdresseEuronext(ByVal isin As String, ByVal marche As String, _
                                ByVal debut As Date, ByVal fin As Date) As String
    AdresseEuronext = ADRESSE_EURONEXT & isin & "-" & marche & _
        "?format=csv&decimal_separator=.&date_form=d/m/Y&adjusted=N" & _
        "&startdate=" & Format$(debut, "yyyy-mm-dd") & "&enddate=" & Format$(fin, "yyyy-mm-dd")
End Function

' Telecharge l'historique d'une action et renvoie le contenu du fichier CSV.
Public Function TelechargerCoursEuronext(ByVal code As String, ByVal isin As String, ByVal marche As String, _
                                         ByVal debut As Date, ByVal fin As Date) As String
    Dim adresse As String, contenu As String
    adresse = AdresseEuronext(isin, marche, debut, fin)
    contenu = TelechargerTexte(adresse)

    If InStr(contenu, "Date;") = 0 Then
        ErreurUtilisateur "Le telechargement des cours de " & code & " (ISIN " & isin & ") sur Euronext a echoue." & vbLf & vbLf & _
                          "Reponse recue : " & Left$(contenu, 200) & vbLf & vbLf & _
                          "Verifiez votre connexion Internet et le code ISIN. Si le probleme persiste, " & _
                          "telechargez les fichiers sur live.euronext.com et choisissez la source 'Fichiers CSV'."
    End If
    ' Euronext renvoie un fichier sans aucune ligne de cours si l'ISIN, le marche ou la periode ne conviennent pas
    If Len(Trim$(Replace(Replace(Mid$(contenu, InStr(contenu, "Date;")), vbCr, ""), vbLf, ""))) < 200 Then
        ErreurUtilisateur "Euronext n'a renvoye aucune cotation pour " & code & " (ISIN " & isin & ", marche " & marche & ")" & _
                          " entre le " & Format$(debut, "dd/mm/yyyy") & " et le " & Format$(fin, "dd/mm/yyyy") & "." & vbLf & vbLf & _
                          "Verifiez le code ISIN et le marche (XPAR = Paris) dans la feuille Parametres."
    End If
    TelechargerCoursEuronext = contenu
End Function

' Telecharge le contenu d'une adresse web (texte).
Public Function TelechargerTexte(ByVal adresse As String) As String
#If Mac Then
    TelechargerTexte = TelechargerAvecCurl(adresse)
#Else
    TelechargerTexte = TelechargerAvecXmlHttp(adresse)
#End If
End Function

#If Mac Then
' Mac : execute "curl" et lit sa sortie.
Private Function TelechargerAvecCurl(ByVal adresse As String) As String
    Dim commande As String, fichier As LongPtr, tampon As String, nbLus As Long, resultat As String
    commande = "/usr/bin/curl --silent --show-error --location --max-time 60 " & _
               "--user-agent 'Mozilla/5.0' '" & adresse & "' 2>&1"
    fichier = c_popen(commande, "r")
    If fichier = 0 Then
        ErreurUtilisateur "Impossible de lancer le telechargement (curl) sur ce Mac." & vbLf & _
                          "Choisissez la source 'Fichiers CSV' et telechargez les fichiers sur live.euronext.com."
    End If
    Do While c_feof(fichier) = 0
        tampon = Space$(4096)
        nbLus = c_fread(tampon, 1, Len(tampon) - 1, fichier)
        If nbLus > 0 Then resultat = resultat & Left$(tampon, nbLus)
    Loop
    c_pclose fichier
    TelechargerAvecCurl = resultat
End Function
#End If

' Windows : requete HTTP classique.
Private Function TelechargerAvecXmlHttp(ByVal adresse As String) As String
    Dim requete As Object
    On Error GoTo Echec
    Set requete = CreateObject("MSXML2.XMLHTTP")
    requete.Open "GET", adresse, False
    requete.setRequestHeader "User-Agent", "Mozilla/5.0"
    requete.send
    If requete.Status <> 200 Then
        TelechargerAvecXmlHttp = "Erreur HTTP " & requete.Status
    Else
        TelechargerAvecXmlHttp = requete.responseText
    End If
    Exit Function
Echec:
    TelechargerAvecXmlHttp = "Erreur : " & Err.Description
End Function
