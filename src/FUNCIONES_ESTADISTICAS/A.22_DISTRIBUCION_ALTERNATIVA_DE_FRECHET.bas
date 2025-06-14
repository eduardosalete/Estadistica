
' FUNCIÓN DE DENSIDAD

Public Function D_Frechet_A(x As Double, Alfa As Double) As Variant
' Esta función calcula la función de densidad de la distribución alternativa de Fréchet
Dim Eps As Double, xx As Double
Dim a1 As Double, a2 As Double

Eps = 0.0000001

If Alfa <= Eps Then
   D_Frechet_A = "Alfa debe ser > 0"
   Exit Function
End If

If x >= 0 Then
   D_Frechet_A = 0
   Exit Function
End If

xx = -x

a1 = xx ^ (-Alfa)
a2 = a1 / xx

If a1 > 250 Then
  D_Frechet_A = 0
Else
  D_Frechet_A = Alfa * Exp(-a1) * a2
End If

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Frechet_A(x As Double, Alfa As Double) As Variant
' Esta función calcula la función de distribución de la distribución alternativa de Fréchet
Dim Eps As Double, aa As Double

Eps = 0.0000001

If Alfa <= Eps Then
   FD_Frechet_A = "Alfa debe ser > 0"
   Exit Function
End If

If x >= 0 Then
   FD_Frechet_A = 1
   Exit Function
End If

aa = (-x) ^ (-Alfa)

FD_Frechet_A = 1 - Exp(-aa)

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Frechet_A_Inv(Probabilidad As Double, Alfa As Double) As Variant
' Esta función obtiene la inversa de la función de distribución alternativa de Fréchet
Dim Eps As Double

Eps = 0.0000001

If Alfa <= Eps Then
   F_Frechet_A_Inv = "Alfa debe ser > 0"
   Exit Function
End If

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_Frechet_A_Inv = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If Probabilidad < Eps Then
   F_Frechet_A_Inv = "-" & ChrW(8734)
   Exit Function
End If

If Abs(Probabilidad - 1) < Eps Then
   F_Frechet_A_Inv = 0
   Exit Function
End If

F_Frechet_A_Inv = -(-Log(1 - Probabilidad)) ^ (-1 / Alfa)

End Function


