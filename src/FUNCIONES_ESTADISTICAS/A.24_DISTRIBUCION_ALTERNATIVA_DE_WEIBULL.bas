
' FUNCIÓN DE DENSIDAD

Public Function D_Weibull_A(x As Double, Lambda As Double, k As Double) As Variant
' Esta función calcula la función de densidad de la distribución alternativa de Weibull
Dim Eps As Double

Eps = 0.0000001

If k <= 0 Or Lambda <= 0 Then
   D_Weibull_A = "k y Lambda deben ser >0"
   Exit Function
End If

If x >= 0 Then
   D_Weibull_A = 0
   Exit Function
End If

If k < 1 And x < Eps Then
   ' Para prevenir infinitos si k<1
   D_Weibull_A = "+" & ChrW(8734)
   Exit Function
End If

D_Weibull_A = -k / x / Lambda * (-x / Lambda) ^ k * Exp(-(-x / Lambda) ^ k)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Weibull_A(x As Double, Lambda As Double, k As Double) As Variant
' Esta función calcula la función de distribución de la distribución alternativa de Weibull

If k <= 0 Or Lambda <= 0 Then
   FD_Weibull_A = "k y Lambda deben ser >0"
   Exit Function
End If

If x > 0 Then
   FD_Weibull_A = 0
   Exit Function
End If

FD_Weibull_A = Exp(-(-x / Lambda) ^ k)

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Weibull_A_Inv(Probabilidad As Double, Lambda As Double, k As Double) As Variant
' Esta función calcula la inversa de la función de distribución de la distribución alternativa de Weibull
Dim Eps As Double

Eps = 0.0000001

If k <= 0 Or Lambda <= 0 Then
   F_Weibull_A_Inv = "k y Lambda deben ser >0"
   Exit Function
End If

If Probabilidad < 0 Then
   F_Weibull_A_Inv = "Probabilidad negativa"
   Exit Function
End If

If Probabilidad > 1 Then
   F_Weibull_A_Inv = "Probabilidad > 1"
   Exit Function
End If

If Probabilidad <= Eps Then
   F_Weibull_A_Inv = "-" & ChrW(8734)
   Exit Function
End If

F_Weibull_A_Inv = -Lambda * (-Log(Probabilidad)) ^ (1 / k)

End Function


