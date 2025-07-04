
' FUNCIÓN DE DENSIDAD

Public Function D_Gumbel_A(x As Double, Mu As Double, Beta As Double) As Variant
' Esta función calcula la función de densidad de la distribución alternativa de Gumbel
Dim z As Double

If Beta <= 0 Then
   D_Gumbel_A = "Beta debe ser >0"
   Exit Function
End If

z = (x - Mu) / Beta
D_Gumbel_A = 1 / Beta * Exp(z - Exp(z))

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Gumbel_A(x As Double, Mu As Double, Beta As Double) As Variant
' Esta función calcula la función de distribución de la distribución alternativa de Gumbel
Dim z As Double

If Beta <= 0 Then
   FD_Gumbel_A = "Beta debe ser >0"
   Exit Function
End If

z = (x - Mu) / Beta
FD_Gumbel_A = 1 - Exp(-Exp(z))

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Gumbel_A_Inv(Probabilidad As Double, Mu As Double, Beta As Double) As Variant
' Esta función calcula la inversa de la función de distribución de la distribución alternativa de Gumbel
Dim z As Double
Dim Eps As Double

Eps = 0.0000001

If Beta <= 0 Then
   F_Gumbel_A_Inv = "Beta debe ser >0"
   Exit Function
End If

If Probabilidad < 0 Then
   F_Gumbel_A_Inv = "Probabilidad negativa"
   Exit Function
End If

If Probabilidad > 1 Then
   F_Gumbel_A_Inv = "Probabilidad > 1"
   Exit Function
End If

If Probabilidad < Eps Then
   F_Gumbel_A_Inv = "+" & ChrW(8734)
   Exit Function
End If

If Probabilidad >= 1 - Eps Then
   F_Gumbel_A_Inv = "-" & ChrW(8734)
   Exit Function
End If

z = Log(-Log(1 - Probabilidad))

F_Gumbel_A_Inv = Beta * z + Mu

End Function


