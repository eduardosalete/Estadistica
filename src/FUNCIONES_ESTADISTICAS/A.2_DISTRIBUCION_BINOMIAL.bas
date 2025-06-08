
' FUNCIÓN DE PROBABILIDAD

Public Function p_Binomial(x As Long, N As Long, p As Double) As Variant
' Función de probabilidad de la distribución Binomial
' N es el número total de pruebas
' x el número total de aciertos
' p la probabilidad de acertar en un ensayo
' Llama a la función N_Combinatorio

Dim q As Double, Eps As Double

Eps = 0.0000001

If p < 0 Or p > 1 Then
   p_Binomial = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If x < 0 Or x > N Then
   p_Binomial = 0
   Exit Function
End If

If p < Eps Then
   p_Binomial = 0
   Exit Function
End If

q = 1 - p
p_Binomial = N_Combinatorio(N, x) * p ^ x * q ^ (N - x)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Binomial(x As Double, N As Long, p As Double) As Variant
' Función de distribución de la distribución Binomial P(Psi<=x)
' Llama a la función p_Binomial
Dim i As Long, ix As Long
Dim Eps As Double

Eps = 0.0000001

If p < 0 Or p > 1 Then
   F_Binomial = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If x < 0 Then
   F_Binomial = 0
   Exit Function
End If

If x >= N Then
   F_Binomial = 1
   Exit Function
End If

If p < Eps Then
   F_Binomial = 0
   Exit Function
End If

If p > 1 Then
   F_Binomial = -1
   Exit Function
End If

ix = Fix(x)
F_Binomial = 0
For i = 0 To ix
    F_Binomial = F_Binomial + p_Binomial(i, N, p)
Next i

End Function


