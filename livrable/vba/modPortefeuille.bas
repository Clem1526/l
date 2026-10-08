Attribute VB_Name = "modPortefeuille"
Option Explicit
'==============================================================================
' modPortefeuille : rendement et risque d'un portefeuille selon les poids
'   R_P      = somme_i  x_i * E(R_i)
'   Var(R_P) = somme_i somme_j  x_i * x_j * Cov(R_i, R_j)
' (avec deux titres : xA^2 Var(A) + xB^2 Var(B) + 2 xA xB Cov(A,B))
'==============================================================================

Public Function RendementPortefeuille(ByRef poids() As Double, ByRef moy() As Double) As Double
    Dim i As Long, total As Double
    For i = 1 To UBound(moy)
        total = total + poids(i) * moy(i)
    Next i
    RendementPortefeuille = total
End Function

Public Function VariancePortefeuille(ByRef poids() As Double, ByRef cov() As Double) As Double
    Dim i As Long, j As Long, total As Double
    For i = 1 To UBound(cov, 1)
        For j = 1 To UBound(cov, 2)
            total = total + poids(i) * poids(j) * cov(i, j)
        Next j
    Next i
    VariancePortefeuille = total
End Function

Public Function EcartTypePortefeuille(ByRef poids() As Double, ByRef cov() As Double) As Double
    Dim v As Double
    v = VariancePortefeuille(poids, cov)
    If v < 0 Then v = 0                      ' securite contre les erreurs d'arrondi
    EcartTypePortefeuille = Sqr(v)
End Function
