Attribute VB_Name = "Module3_Calculs"
Option Explicit
'==============================================================================
'  MODULE 3 : LES CALCULS (etapes 2 a 5 de la consigne)
'==============================================================================
'  Formules du cours "Utiliser la theorie du portefeuille" :
'
'    Rendement du jour t   : R(t) = ( D(t) + P(t) - P(t-1) ) / P(t-1)
'    Rendement moyen       : E(R) = somme des R(t) / n
'    Variance              : Var(R) = somme des (R(t) - E(R))^2 / n
'    Ecart-type            : sigma(R) = racine( Var(R) )
'    Covariance            : Cov(Ri, Rj) = somme des (Ri(t) - E(Ri)) * (Rj(t) - E(Rj)) / n
'
'    Portefeuille de poids x1, x2, ... (somme = 100 %) :
'      Rendement   R(pf)   = somme des xi * E(Ri)
'      Variance    Var(pf) = somme sur i et j des xi * xj * Cov(Ri, Rj)
'      (avec 2 titres : xA^2 Var(A) + xB^2 Var(B) + 2 xA xB Cov(A,B))
'
'  Remarque : comme dans le cours, on divise par n (et non par n - 1).
'  Dans Excel, cela correspond a VAR.P.N, ECARTYPE.PEARSON et COVARIANCE.PEARSON.
'==============================================================================

' Variables utilisees seulement dans ce module, pour tester tous les portefeuilles
Private PoidsEnCours() As Double       ' le portefeuille en cours de construction
Private NbPas As Long                  ' 100 % = NbPas fois PasPoids (ex. 20 pas de 5 %)


'------------------------------------------------------------------------------
' Etape 2 : rendements journaliers de chaque action
'------------------------------------------------------------------------------
Public Sub CalculerRendements()
    Dim t As Long, i As Long
    ' Il y a un rendement de moins que de jours : le premier jour n'a pas de "veille"
    ReDim Rendements(1 To NbJours - 1, 1 To NbActions)
    For i = 1 To NbActions
        For t = 2 To NbJours
            Rendements(t - 1, i) = (Dividendes(t, i) + Prix(t, i) - Prix(t - 1, i)) / Prix(t - 1, i)
        Next t
    Next i
End Sub

'------------------------------------------------------------------------------
' Etape 2 : rendement moyen de chaque action
'------------------------------------------------------------------------------
Public Sub CalculerMoyennes()
    Dim t As Long, i As Long, n As Long, somme As Double
    n = NbJours - 1                                    ' nombre de rendements
    ReDim Moyennes(1 To NbActions)
    For i = 1 To NbActions
        somme = 0
        For t = 1 To n
            somme = somme + Rendements(t, i)
        Next t
        Moyennes(i) = somme / n
    Next i
End Sub

'------------------------------------------------------------------------------
' Etapes 2 et 3 : matrice de variance-covariance
' Covariances(i, i) est la variance de l'action i ; l'ecart-type est sa racine.
'------------------------------------------------------------------------------
Public Sub CalculerCovariances()
    Dim t As Long, i As Long, j As Long, n As Long, somme As Double
    n = NbJours - 1
    ReDim Covariances(1 To NbActions, 1 To NbActions)
    For i = 1 To NbActions
        For j = 1 To NbActions
            somme = 0
            For t = 1 To n
                somme = somme + (Rendements(t, i) - Moyennes(i)) * (Rendements(t, j) - Moyennes(j))
            Next t
            Covariances(i, j) = somme / n
        Next j
    Next i
End Sub

' Ecart-type de l'action i (racine de sa variance).
Public Function EcartType(ByVal i As Long) As Double
    EcartType = Sqr(Covariances(i, i))
End Function

' Coefficient de correlation entre les actions i et j.
Public Function Correlation(ByVal i As Long, ByVal j As Long) As Double
    Correlation = Covariances(i, j) / (EcartType(i) * EcartType(j))
End Function


'------------------------------------------------------------------------------
' Etape 4 : rendement et risque d'un portefeuille selon les poids
' "poids" est un tableau : poids(1) = part de l'action 1, poids(2) = part de l'action 2...
'------------------------------------------------------------------------------
Public Function RendementPortefeuille(ByRef poids() As Double) As Double
    Dim i As Long, total As Double
    For i = 1 To NbActions
        total = total + poids(i) * Moyennes(i)
    Next i
    RendementPortefeuille = total
End Function

Public Function RisquePortefeuille(ByRef poids() As Double) As Double
    Dim i As Long, j As Long, variance As Double
    For i = 1 To NbActions
        For j = 1 To NbActions
            variance = variance + poids(i) * poids(j) * Covariances(i, j)
        Next j
    Next i
    If variance < 0 Then variance = 0                  ' securite (arrondis)
    RisquePortefeuille = Sqr(variance)                 ' le risque est l'ecart-type
End Function


'==============================================================================
' Etape 5 : TESTER TOUS LES PORTEFEUILLES POSSIBLES
'==============================================================================
'  Idee : on essaie toutes les facons de repartir 100 % entre les actions, par
'  pas de PasPoids (ex. 5 %). Exemple avec 3 actions et un pas de 50 % :
'      (100%, 0, 0) (50%, 50%, 0) (50%, 0, 50%) (0, 100%, 0) (0, 50%, 50%) (0, 0, 100%)
'  Pour chacun, on calcule le rendement et le risque. La frontiere efficiente
'  est ensuite obtenue en gardant, pour chaque niveau de rendement, le
'  portefeuille le moins risque (voir Module4, AfficherFrontiere).
'
'  Pas de vente a decouvert : tous les poids sont entre 0 et 100 %.
'==============================================================================

Public Sub TesterTousLesPortefeuilles()
    Dim m As Long, nbPrevus As Double

    NbPas = Round(1 / PasPoids)
    nbPrevus = NombreDeCombinaisons(NbPas, NbActions)
    If nbPrevus > 300000 Then
        Signaler "Trop de portefeuilles a tester (" & nbPrevus & ") : augmentez le pas des poids " & _
                 "(cellule C7), par exemple 10 %."
    End If
    NbPortefeuilles = nbPrevus

    ReDim PoidsPortefeuilles(1 To NbPortefeuilles, 1 To NbActions)
    ReDim RendementsPf(1 To NbPortefeuilles)
    ReDim RisquesPf(1 To NbPortefeuilles)
    ReDim PoidsEnCours(1 To NbActions)

    NbPortefeuilles = 0                                ' compteur, augmente a chaque portefeuille
    ChoisirPoids 1, NbPas

    ' Le portefeuille de variance minimale est celui qui a le plus petit risque
    IndiceVarianceMin = 1
    For m = 2 To NbPortefeuilles
        If RisquesPf(m) < RisquesPf(IndiceVarianceMin) Then IndiceVarianceMin = m
    Next m
End Sub

' Choisit le poids de l'action numero "action", sachant qu'il reste "pasRestants" pas a distribuer.
' Cette procedure s'appelle elle-meme pour l'action suivante (c'est une procedure "recursive") :
' c'est l'equivalent d'autant de boucles For imbriquees qu'il y a d'actions.
Private Sub ChoisirPoids(ByVal action As Long, ByVal pasRestants As Long)
    Dim k As Long
    If action = NbActions Then
        ' Derniere action : elle prend tout ce qui reste, pour que la somme fasse 100 %
        PoidsEnCours(action) = pasRestants / NbPas
        EnregistrerPortefeuille
    Else
        For k = 0 To pasRestants
            PoidsEnCours(action) = k / NbPas
            ChoisirPoids action + 1, pasRestants - k
        Next k
    End If
End Sub

' Calcule le rendement et le risque du portefeuille en cours et le range dans la liste.
Private Sub EnregistrerPortefeuille()
    Dim i As Long
    NbPortefeuilles = NbPortefeuilles + 1
    For i = 1 To NbActions
        PoidsPortefeuilles(NbPortefeuilles, i) = PoidsEnCours(i)
    Next i
    RendementsPf(NbPortefeuilles) = RendementPortefeuille(PoidsEnCours)
    RisquesPf(NbPortefeuilles) = RisquePortefeuille(PoidsEnCours)
End Sub

' Nombre de facons de repartir "pas" pas entre "n" actions : combinaison C(pas + n - 1, n - 1).
Private Function NombreDeCombinaisons(ByVal pas As Long, ByVal n As Long) As Double
    Dim k As Long, resultat As Double
    resultat = 1
    For k = 1 To n - 1
        resultat = resultat * (pas + k) / k
    Next k
    NombreDeCombinaisons = Round(resultat)
End Function
