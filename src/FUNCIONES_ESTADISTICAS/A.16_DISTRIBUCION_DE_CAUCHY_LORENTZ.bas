
' FUNCIÓN DE DENSIDAD

Public Function D_Cauchy(x As Double, x0 As Double, Gamma As Double) As Variant
' Esta función calcula la función de densidad de la distribución de Cauchy-Lorentz
Dim Pi As Double

Pi = 3.14159265358979

If Gamma <= 0 Then
   D_Cauchy = "Gamma debe ser >0"
   Exit Function
End If

D_Cauchy = 1 + ((x - x0) / Gamma) ^ 2
D_Cauchy = 1 / Gamma / Pi / D_Cauchy

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Cauchy(x As Double, x0 As Double, Gamma As Double) As Variant
' Esta función calcula la función de distribución de la distribución de Cauchy-Lorentz
Dim Pi As Double

If Gamma <= 0 Then
   FD_Cauchy = "Gamma debe ser >0"
   Exit Function
End If

Pi = 3.14159265358979
FD_Cauchy = Atn((x - x0) / Gamma) / Pi + 0.5

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Cauchy_Inv(Probabilidad As Double, x0 As Double, Gamma As Double) As Variant
' Esta función obtiene la inversa de la función de distribución de Cauchy-Lorentz
Dim Pi As Double, Eps As Double

Eps = 0.0000001
Pi = 3.14159265358979

If Gamma <= 0 Then
   F_Cauchy_Inv = "Gamma debe ser >0"
   Exit Function
End If

If Probabilidad < 0 Then
   F_Cauchy_Inv = "Probabilidad negativa"
   Exit Function
End If

If Probabilidad > 1 Then
   F_Cauchy_Inv = "Probabilidad > 1"
   Exit Function
End If

If Probabilidad <= Eps Then
   F_Cauchy_Inv = "-" & ChrW(8734)
   Exit Function
End If

If Probabilidad >= 1 - Eps Then
   F_Cauchy_Inv = "+" & ChrW(8734)
   Exit Function
End If

F_Cauchy_Inv = x0 + Gamma * Tan(Pi * (Probabilidad - 0.5))

End Function


