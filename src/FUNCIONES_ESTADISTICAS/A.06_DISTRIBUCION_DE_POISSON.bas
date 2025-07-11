
' FUNCIÓN DE PROBABILIDAD

Public Function p_Poisson(x As Long, Lambda As Double) As Variant
' Función de probabilidad de la distribución de Poisson
' x el número total de ocurrencias
' Lambda es el parámetro de la distribución
' Llama a la función Mi_Factorial

Dim q As Double, Eps As Double

Eps = 0.0000001

If Lambda <= 0 Then
   p_Poisson = "Lambda debe ser >0"
   Exit Function
End If

If x < 0 Then
   p_Poisson = 0
   Exit Function
End If

If Lambda < Eps Then
   p_Poisson = 0
   Exit Function
End If

p_Poisson = Exp(-Lambda) * Lambda ^ x / Mi_Factorial(x)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Poisson(x As Double, Lambda As Double) As Variant
' Función de distribución de la distribución de Poisson P(Psi<=x)
' Llama a la función p_Poisson

Dim i As Long, ix As Long
Dim Eps As Double

Eps = 0.0000001

If Lambda <= 0 Then
   p_Poisson = "Lambda debe ser >0"
   Exit Function
End If

If x < 0 Then
   F_Poisson = 0
   Exit Function
End If

If Lambda < Eps Then
   F_Poisson = 0
   Exit Function
End If

ix = Fix(x)
F_Poisson = 0
For i = 0 To ix
    F_Poisson = F_Poisson + p_Poisson(i, Lambda)
Next i

End Function


