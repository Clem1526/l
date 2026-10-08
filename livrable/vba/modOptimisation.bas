Attribute VB_Name = "modOptimisation"
Option Explicit
'==============================================================================
' modOptimisation : portefeuille de variance minimale et frontiere efficiente
'
' Contraintes : poids entre 0 et 100 %, somme des poids = 100 % (pas de vente
' a decouvert). Le Solveur d'Excel n'est pas utilise (peu fiable sur Mac).
'
' Methode : pour un "gout pour le rendement" lambda donne, on minimise
'     Var(R_P) - lambda * E(R_P)
' par descente de gradient projetee. lambda = 0 donne le portefeuille de
' variance minimale ; plus lambda est grand, plus le portefeuille est rentable
' (et risque). Pour un rendement cible, on cherche le lambda qui l'atteint par
' dichotomie : on obtient le portefeuille de risque minimal pour ce rendement.
'==============================================================================

Private Const TOLERANCE_POIDS As Double = 0.000000000001    ' 1E-12 : arret de la descente
Private Const MAX_ITERATIONS As Long = 100000
Private Const MAX_DICHOTOMIE As Long = 100

' Calcule la frontiere de variance minimale, de la plus faible a la plus forte rentabilite.
'   poids(1 To K, 1 To N), rendements(1 To K), ecartsTypes(1 To K), efficient(1 To K)
'   wMvp(1 To N) : portefeuille de variance minimale ; nbDomines : points sous ce portefeuille
Public Sub CalculerFrontiere(ByRef cov() As Double, ByRef moy() As Double, ByVal nbPointsEfficients As Long, _
                             ByRef poids() As Double, ByRef rendements() As Double, ByRef ecartsTypes() As Double, _
                             ByRef efficient() As Boolean, ByRef wMvp() As Double, ByRef nbDomines As Long)
    Dim n As Long, k As Long, i As Long, nbPoints As Long
    Dim iMin As Long, iMax As Long, rMvp As Double, cible As Double, ecartMu As Double
    Dim w() As Double

    n = UBound(moy)
    iMin = 1: iMax = 1
    For i = 2 To n
        If moy(i) < moy(iMin) Then iMin = i
        If moy(i) > moy(iMax) Then iMax = i
    Next i
    ecartMu = moy(iMax) - moy(iMin)

    wMvp = PortefeuilleVarianceMinimale(cov, moy)
    rMvp = RendementPortefeuille(wMvp, moy)

    ' Partie dominee (sous le portefeuille de variance minimale), si elle existe
    nbDomines = 0
    If rMvp - moy(iMin) > 0.000001 * ecartMu Then
        nbDomines = nbPointsEfficients \ 2
        If nbDomines < 2 Then nbDomines = 2
    End If
    If moy(iMax) - rMvp <= 0.000001 * ecartMu Then nbPointsEfficients = 1

    nbPoints = nbDomines + nbPointsEfficients
    ReDim poids(1 To nbPoints, 1 To n)
    ReDim rendements(1 To nbPoints)
    ReDim ecartsTypes(1 To nbPoints)
    ReDim efficient(1 To nbPoints)

    ' Partie efficiente : du portefeuille de variance minimale a l'action la plus rentable
    w = Copie(wMvp)
    For k = 1 To nbPointsEfficients
        Application.StatusBar = "Calcul de la frontiere efficiente : point " & k & " / " & nbPointsEfficients
        If k = 1 Then
            w = Copie(wMvp)
        ElseIf k = nbPointsEfficients Then
            w = PortefeuilleUnTitre(n, iMax)
        Else
            cible = rMvp + (k - 1) * (moy(iMax) - rMvp) / (nbPointsEfficients - 1)
            w = PortefeuillePourRendementCible(cov, moy, cible, w, rMvp)
        End If
        StockerPoint nbDomines + k, w, cov, moy, True, poids, rendements, ecartsTypes, efficient
    Next k

    ' Partie dominee : de l'action la moins rentable au portefeuille de variance minimale
    w = Copie(wMvp)
    For k = nbDomines To 1 Step -1
        If k = 1 Then
            w = PortefeuilleUnTitre(n, iMin)
        Else
            cible = moy(iMin) + (k - 1) * (rMvp - moy(iMin)) / nbDomines
            w = PortefeuillePourRendementCible(cov, moy, cible, w, rMvp)
        End If
        StockerPoint k, w, cov, moy, False, poids, rendements, ecartsTypes, efficient
    Next k
End Sub

' Portefeuille de variance minimale (lambda = 0).
Public Function PortefeuilleVarianceMinimale(ByRef cov() As Double, ByRef moy() As Double) As Variant   ' tableau de Double
    Dim covN() As Double, moyN() As Double, w() As Double, i As Long, n As Long
    n = UBound(moy)
    Normaliser cov, moy, covN, moyN
    ReDim w(1 To n)
    For i = 1 To n
        w(i) = 1 / n                         ' depart : portefeuille equipondere
    Next i
    MinimiserCritere covN, moyN, 0, w
    PortefeuilleVarianceMinimale = w
End Function

' Portefeuille de risque minimal ayant le rendement "cible".
' wDepart : point de depart (accelere le calcul) ; rMvp : rendement du portefeuille de variance minimale.
Public Function PortefeuillePourRendementCible(ByRef cov() As Double, ByRef moy() As Double, ByVal cible As Double, _
                                               ByRef wDepart() As Double, ByVal rMvp As Double) As Variant   ' tableau de Double
    Dim covN() As Double, moyN() As Double, w() As Double
    Dim sens As Double, bas As Double, haut As Double, milieu As Double
    Dim tolerance As Double, ecart As Double, iter As Long

    Normaliser cov, moy, covN, moyN
    tolerance = 0.00000001 * EtendueMoyennes(moy)
    If tolerance <= 0 Then tolerance = 1E-15

    ' Cible au-dessus du portefeuille de variance minimale : lambda > 0 ; en dessous : lambda < 0
    If cible >= rMvp Then sens = 1 Else sens = -1

    ' 1) On encadre le bon lambda entre "bas" et "haut" en doublant "haut"
    w = Copie(wDepart)
    bas = 0
    haut = sens
    For iter = 1 To 80
        MinimiserCritere covN, moyN, haut, w
        If sens * (RendementPortefeuille(w, moy) - cible) >= 0 Then Exit For
        bas = haut
        haut = haut * 2
    Next iter

    ' 2) Dichotomie entre "bas" et "haut"
    For iter = 1 To MAX_DICHOTOMIE
        milieu = (bas + haut) / 2
        MinimiserCritere covN, moyN, milieu, w
        ecart = RendementPortefeuille(w, moy) - cible
        If Abs(ecart) < tolerance Then Exit For
        If sens * ecart < 0 Then bas = milieu Else haut = milieu
    Next iter

    PortefeuillePourRendementCible = w
End Function

' Portefeuille investi a 100 % dans l'action i.
Public Function PortefeuilleUnTitre(ByVal n As Long, ByVal i As Long) As Variant   ' tableau de Double
    Dim w() As Double
    ReDim w(1 To n)
    w(i) = 1
    PortefeuilleUnTitre = w
End Function

' Tire au hasard "nb" portefeuilles (poids positifs, somme = 100 %) pour le nuage de points.
Public Sub GenererPortefeuillesAleatoires(ByRef cov() As Double, ByRef moy() As Double, ByVal nb As Long, _
                                          ByRef rendements() As Double, ByRef ecartsTypes() As Double)
    Dim k As Long, i As Long, n As Long, total As Double, w() As Double
    n = UBound(moy)
    ReDim w(1 To n)
    ReDim rendements(1 To nb)
    ReDim ecartsTypes(1 To nb)
    Dim graine As Single
    graine = Rnd(-1)                        ' Rnd(-1) puis Randomize : graine fixe,
    Randomize 2025                          ' donc memes resultats a chaque execution
    For k = 1 To nb
        total = 0
        For i = 1 To n
            w(i) = -Log(1 - Rnd())          ' tirage uniforme sur l'ensemble des poids possibles
            total = total + w(i)
        Next i
        For i = 1 To n
            w(i) = w(i) / total
        Next i
        rendements(k) = RendementPortefeuille(w, moy)
        ecartsTypes(k) = EcartTypePortefeuille(w, cov)
    Next k
End Sub

'------------------------------------------------------------------------------
' Outils internes
'------------------------------------------------------------------------------

' Enregistre un point de la frontiere dans les tableaux de resultats.
Private Sub StockerPoint(ByVal k As Long, ByRef w() As Double, ByRef cov() As Double, ByRef moy() As Double, _
                         ByVal estEfficient As Boolean, ByRef poids() As Double, ByRef rendements() As Double, _
                         ByRef ecartsTypes() As Double, ByRef efficient() As Boolean)
    Dim i As Long
    For i = 1 To UBound(w)
        poids(k, i) = w(i)
    Next i
    rendements(k) = RendementPortefeuille(w, moy)
    ecartsTypes(k) = EcartTypePortefeuille(w, cov)
    efficient(k) = estEfficient
End Sub

' Met les donnees a une echelle proche de 1 pour que la descente converge bien
' (les variances journalieres sont de l'ordre de 0,0001). Ne change pas la solution.
Private Sub Normaliser(ByRef cov() As Double, ByRef moy() As Double, ByRef covN() As Double, ByRef moyN() As Double)
    Dim n As Long, i As Long, j As Long, echelleCov As Double, echelleMoy As Double
    n = UBound(moy)
    ReDim covN(1 To n, 1 To n)
    ReDim moyN(1 To n)
    echelleCov = SommeLigneMax(cov)
    If echelleCov <= 0 Then echelleCov = 1
    For i = 1 To n
        If Abs(moy(i)) > echelleMoy Then echelleMoy = Abs(moy(i))
    Next i
    If echelleMoy <= 0 Then echelleMoy = 1
    For i = 1 To n
        moyN(i) = moy(i) / echelleMoy
        For j = 1 To n
            covN(i, j) = cov(i, j) / echelleCov
        Next j
    Next i
End Sub

' Minimise Var - lambda * E sur l'ensemble des poids autorises (w : point de depart, puis resultat).
Private Sub MinimiserCritere(ByRef covN() As Double, ByRef moyN() As Double, ByVal lambda As Double, ByRef w() As Double)
    Dim n As Long, i As Long, j As Long, iter As Long
    Dim pas As Double, gradient As Double, ecart As Double
    Dim v() As Double, wNouveau() As Double

    n = UBound(w)
    ReDim v(1 To n)
    ReDim wNouveau(1 To n)
    pas = 1 / (2 * SommeLigneMax(covN))     ' pas assurant la convergence

    For iter = 1 To MAX_ITERATIONS
        ' Un pas dans la direction qui fait baisser le critere...
        For i = 1 To n
            gradient = -lambda * moyN(i)
            For j = 1 To n
                gradient = gradient + 2 * covN(i, j) * w(j)
            Next j
            v(i) = w(i) - pas * gradient
        Next i
        ' ... puis retour vers des poids valides (positifs, somme = 100 %)
        ProjeterSurSimplexe v, wNouveau
        ecart = 0
        For i = 1 To n
            If Abs(wNouveau(i) - w(i)) > ecart Then ecart = Abs(wNouveau(i) - w(i))
            w(i) = wNouveau(i)
        Next i
        If ecart < TOLERANCE_POIDS Then Exit Sub
    Next iter
End Sub

' Ramene un vecteur v au plus proche vecteur de poids valides w (w >= 0, somme = 1).
' On cherche theta tel que la somme des max(v_i - theta, 0) vaille 1.
Private Sub ProjeterSurSimplexe(ByRef v() As Double, ByRef w() As Double)
    Dim n As Long, i As Long, j As Long
    Dim u() As Double, tmp As Double, cumul As Double, theta As Double

    n = UBound(v)
    ReDim u(1 To n)
    For i = 1 To n
        u(i) = v(i)
    Next i
    For i = 2 To n                          ' tri decroissant (tri par insertion)
        tmp = u(i)
        j = i - 1
        Do While j >= 1
            If u(j) >= tmp Then Exit Do
            u(j + 1) = u(j)
            j = j - 1
        Loop
        u(j + 1) = tmp
    Next i

    cumul = 0
    For j = 1 To n
        cumul = cumul + u(j)
        If u(j) - (cumul - 1) / j > 0 Then theta = (cumul - 1) / j
    Next j
    For i = 1 To n
        w(i) = v(i) - theta
        If w(i) < 0 Then w(i) = 0
    Next i
End Sub

'' Copie independante d'un vecteur de poids (pour ne pas modifier l'original).
Private Function Copie(ByRef source() As Double) As Variant
    Dim resultat() As Double, i As Long
    ReDim resultat(LBound(source) To UBound(source))
    For i = LBound(source) To UBound(source)
        resultat(i) = source(i)
    Next i
    Copie = resultat
End Function

' Plus grande somme des valeurs absolues d'une ligne (majore la plus grande valeur propre).
Private Function SommeLigneMax(ByRef m() As Double) As Double
    Dim i As Long, j As Long, somme As Double, maxi As Double
    For i = 1 To UBound(m, 1)
        somme = 0
        For j = 1 To UBound(m, 2)
            somme = somme + Abs(m(i, j))
        Next j
        If somme > maxi Then maxi = somme
    Next i
    SommeLigneMax = maxi
End Function

Private Function EtendueMoyennes(ByRef moy() As Double) As Double
    Dim i As Long, mini As Double, maxi As Double
    mini = moy(1): maxi = moy(1)
    For i = 2 To UBound(moy)
        If moy(i) < mini Then mini = moy(i)
        If moy(i) > maxi Then maxi = moy(i)
    Next i
    EtendueMoyennes = maxi - mini
End Function
