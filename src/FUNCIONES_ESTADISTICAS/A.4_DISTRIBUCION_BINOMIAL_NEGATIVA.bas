
' FUNCIÓN DE PROBABILIDAD

Public Function p_BinomialN(x As Long, N As Long, p As Double) As Variant
' Función de probabilidad de la distribución Binomial Negativa
' N es el número de éxitos que se quiere obtener
' x es el número de fallos
' p la probabilidad de acertar en un ensayo
' Llama a la función N_Combinatorio

Dim q As Double, Eps As Double

Eps = 0.0000001

If p < 0 Or p > 1 Then
   p_BinomialN = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If x < 0 Then
   p_BinomialN = 0
   Exit Function
End If

If p < Eps Then
   p_BinomialN = 0
   Exit Function
End If

q = 1 - p
p_BinomialN = N_Combinatorio(N + x - 1, N - 1) * p ^ N * q ^ x

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function F_BinomialN(x As Double, N As Long, p As Double) As Variant
' Función de distribución de la distribución Binomial Negativa P(Psi<=x)
' Llama a la función p_BinomialN

Dim i As Long, ix As Long
Dim Eps As Double

Eps = 0.0000001

If p < 0 Or p > 1 Then
   F_BinomialN = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If x < 0 Then
   F_BinomialN = 0
   Exit Function
End If

If p < Eps Then
   F_BinomialN = 0
   Exit Function
End If

ix = Fix(x)
F_BinomialN = 0
For i = 0 To ix
    F_BinomialN = F_BinomialN + p_BinomialN(i, N, p)
Next i

End Function


