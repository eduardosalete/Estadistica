
' FUNCIÓN DE DENSIDAD

Public Function D_ParetoT(x As Double, Alfa As Double, x0 As Double, xM As Double) As Variant
' Esta función calcula la función de densidad de la distribución de Pareto truncada

If Alfa <= 0 Or x0 <= 0 Then
   D_ParetoT = "Alfa y x0 deben ser >0"
   Exit Function
End If

If xM <= x0 Then
   D_ParetoT = "xM debe ser mayor que x0"
   Exit Function
End If

If x < x0 Or x > xM Then
   D_ParetoT = 0
   Exit Function
End If

D_ParetoT = Alfa * x0 ^ Alfa / x ^ (Alfa + 1) / (1 - (x0 / xM) ^ Alfa)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_ParetoT(x As Double, Alfa As Double, x0 As Double, xM As Double) As Variant
' Esta función calcula la función de distribución de la distribución de Pareto truncada

If Alfa <= 0 Or x0 <= 0 Then
   FD_ParetoT = "Alfa y x0 deben ser >0"
   Exit Function
End If

If xM <= x0 Then
   FD_ParetoT = "xM debe ser mayor que x0"
   Exit Function
End If

If x < x0 Then
   FD_ParetoT = 0
   Exit Function
End If

If x > xM Then
   FD_ParetoT = 1
   Exit Function
End If

FD_ParetoT = (1 - (x0 / x) ^ Alfa) / (1 - (x0 / xM) ^ Alfa)

End Function


' COMPLEMENTO A LA FUNCIÓN DE DISTRIBUCIÓN

Public Function CFD_ParetoT(x As Double, Alfa As Double, x0 As Double, xM As Double) As Variant
' Esta función calcula el complemento a la función de distribución de la distribución de Pareto truncada

If Alfa <= 0 Or x0 <= 0 Then
   CFD_ParetoT = "Alfa y x0 deben ser >0"
   Exit Function
End If

If xM <= x0 Then
   FD_ParetoT = "xM debe ser mayor que x0"
   Exit Function
End If

If IsNumeric(FD_ParetoT) Then
   CFD_ParetoT = 1 - FD_ParetoT
Else
   CFD_ParetoT = "Error en datos"
End If

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_ParetoT_Inv(Probabilidad As Double, Alfa As Double, x0 As Double, xM As Double) As Variant
' Esta función calcula la inversa de la función de distribución de la distribución de Pareto truncada
Dim Beta As Double
Dim Eps As Double

Eps = 0.0000001

If Alfa <= 0 Or x0 <= 0 Then
   F_ParetoT_Inv = "Alfa y x0 deben ser >0"
   Exit Function
End If

If xM <= x0 Then
   F_ParetoT_Inv = "xM debe ser mayor que x0"
   Exit Function
End If

If Probabilidad < 0 Then
   F_ParetoT_Inv = "Probabilidad negativa"
   Exit Function
End If

If Probabilidad > 1 Then
   F_ParetoT_Inv = "Probabilidad > 1"
   Exit Function
End If

If Probabilidad >= 1 - Eps Then
   F_ParetoT_Inv = xM
   Exit Function
End If

Beta = 1 - (x0 / xM) ^ Alfa
F_ParetoT_Inv = x0 * (1 - Probabilidad * Beta) ^ (-1 / Alfa)

End Function


