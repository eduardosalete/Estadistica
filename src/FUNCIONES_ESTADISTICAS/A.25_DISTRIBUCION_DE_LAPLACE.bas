
' FUNCIÓN DE DENSIDAD

Public Function D_Laplace(x As Double, Mu As Double, b As Double) As Variant
' Esta función calcula la función de densidad de la distribución de Laplace(Mu,b)
Dim Eps As Double

Eps = 0.0000001
If b <= Eps Then
   D_Laplace = "b debe ser > 0"
   Exit Function
End If

D_Laplace = 1 / 2 / b * Exp(-Abs(x - Mu) / b)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Laplace(x As Double, Mu As Double, b As Double) As Variant
' Esta función calcula la función de distribución de la distribución de Laplace(Mu,b)

Dim Eps As Double
Dim Signo As Integer

Eps = 0.0000001
If b <= Eps Then
   FD_Laplace = "b debe ser > 0"
   Exit Function
End If

Signo = 0
If x > Mu Then Signo = 1
If x < Mu Then Signo = -1

FD_Laplace = 0.5 + 0.5 * Signo * (1 - Exp(-Abs(x - Mu) / b))

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Laplace_Inv(Probabilidad As Double, Mu As Double, b As Double) As Variant
' Esta función obtiene la inversa de la función de distribución de Laplace(Mu,b)

Dim Eps As Double

Eps = 0.0000001
If b <= Eps Then
   F_Laplace_Inv = "b debe ser > 0"
   Exit Function
End If

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_Laplace_Inv = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If Probabilidad < Eps Then
   F_Laplace_Inv = "-" & ChrW(8734)
   Exit Function
End If

If Abs(Probabilidad - 0.5) < Eps Then
   F_Laplace_Inv = Mu
   Exit Function
End If

If Abs(Probabilidad - 1) < Eps Then
   F_Laplace_Inv = ChrW(8734)
   Exit Function
End If

If Probabilidad > 0.5 Then
   F_Laplace_Inv = Mu - b * Log(1 - 2 * (Probabilidad - 0.5))
Else
   F_Laplace_Inv = Mu + b * Log(1 + 2 * (Probabilidad - 0.5))
End If

End Function


