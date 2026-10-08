Attribute VB_Name = "modStats"
Option Explicit
'==============================================================================
' modStats : rendements, moyennes, variances, covariances, correlations
' Conventions du cours "Utiliser la theorie du portefeuille" :
'   R_t = (D_t + P_t - P_(t-1)) / P_(t-1)
'   moyenne arithmetique ; variance et covariance divisees par n (et non n - 1)
'==============================================================================

' Rendements journaliers : rend(1 To T-1, 1 To N)
Public Function CalculerRendements(ByRef prix() As Double, ByRef dividendes() As Double) As Variant   ' tableau de Double
    Dim rend() As Double, t As Long, i As Long
    ReDim rend(1 To UBound(prix, 1) - 1, 1 To UBound(prix, 2))
    For i = 1 To UBound(prix, 2)
        For t = 2 To UBound(prix, 1)
            rend(t - 1, i) = (dividendes(t, i) + prix(t, i) - prix(t - 1, i)) / prix(t - 1, i)
        Next t
    Next i
    CalculerRendements = rend
End Function

' Rendement moyen de chaque action : moy(1 To N)
Public Function CalculerMoyennes(ByRef rend() As Double) As Variant   ' tableau de Double
    Dim moy() As Double, t As Long, i As Long, nbObs As Long
    nbObs = UBound(rend, 1)
    ReDim moy(1 To UBound(rend, 2))
    For i = 1 To UBound(rend, 2)
        For t = 1 To nbObs
            moy(i) = moy(i) + rend(t, i)
        Next t
        moy(i) = moy(i) / nbObs
    Next i
    CalculerMoyennes = moy
End Function

' Matrice de variance-covariance : cov(i, j) = 1/n * somme (R_i - moy_i)(R_j - moy_j)
' La diagonale contient les variances.
Public Function CalculerCovariances(ByRef rend() As Double, ByRef moy() As Double) As Variant   ' tableau de Double
    Dim cov() As Double, t As Long, i As Long, j As Long, n As Long, nbObs As Long, somme As Double
    n = UBound(rend, 2)
    nbObs = UBound(rend, 1)
    ReDim cov(1 To n, 1 To n)
    For i = 1 To n
        For j = i To n
            somme = 0
            For t = 1 To nbObs
                somme = somme + (rend(t, i) - moy(i)) * (rend(t, j) - moy(j))
            Next t
            cov(i, j) = somme / nbObs
            cov(j, i) = cov(i, j)          ' la matrice est symetrique
        Next j
    Next i
    CalculerCovariances = cov
End Function

' Coefficient de correlation entre les actions i et j
Public Function Correlation(ByRef cov() As Double, ByVal i As Long, ByVal j As Long) As Double
    If cov(i, i) <= 0 Or cov(j, j) <= 0 Then
        Correlation = 0
    Else
        Correlation = cov(i, j) / Sqr(cov(i, i) * cov(j, j))
    End If
End Function
