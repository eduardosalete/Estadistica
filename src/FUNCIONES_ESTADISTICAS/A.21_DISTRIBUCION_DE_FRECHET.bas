
' FUNCIÓN DE DENSIDAD

Public Function D_Frechet(x As Double, Alfa As Double) As Variant
' Esta función calcula la función de densidad de la distribución de Fréchet
Dim Eps As Double
Dim a1 As Double, a2 As Double, a3 As Double

Eps = 0.0000001

If Alfa <= Eps Then
   D_Frechet = "Alfa debe ser > 0"
   Exit Function
End If

If x <= 0 Then
   D_Frechet = 0
   Exit Function
End If

a1 = Alfa * x ^ (-1 - Alfa)
a2 = 1 / x ^ Alfa

If a2 < 500 Then
   a3 = 1 / Exp(a2)
Else
   a3 = 0
End If

D_Frechet = a1 * a3

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Frechet(x As Double, Alfa As Double) As Variant
' Esta función calcula la función de distribución de la distribución de Fréchet
Dim Eps As Double, aa As Double

Eps = 0.0000001

If Alfa <= Eps Then
   FD_Frechet = "Alfa debe ser > 0"
   Exit Function
End If

If x <= 0 Then
   FD_Frechet = 0
   Exit Function
End If

aa = x ^ (-Alfa)

FD_Frechet = Exp(-aa)

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Frechet_Inv(Probabilidad As Double, Alfa As Double) As Variant
' Esta función obtiene la inversa de la función de distribución de Fréchet
Dim Eps As Double

Eps = 0.0000001

If Alfa <= Eps Then
   F_Frechet_Inv = "Alfa debe ser > 0"
   Exit Function
End If

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_Frechet_Inv = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If Probabilidad < Eps Then
   F_Frechet_Inv = 0
   Exit Function
End If

If Abs(Probabilidad - 1) < Eps Then
   F_Frechet_Inv = ChrW(8734)
   Exit Function
End If

F_Frechet_Inv = (-1 / Log(Probabilidad)) ^ (1 / Alfa)

End Function


' FUNCIÓN F_Frechet_Media

Public Function F_Frechet_Media(Alfa As Double) As Variant
' Calcula la media de la distribución de Fréchet
' Llama a la función F_Gamma

If Alfa <= 0 Then
   F_Frechet_Media = "Alfa debe ser > 0"
   Exit Function
End If

If Alfa > 1 Then
   F_Frechet_Media = F_Gamma(1 - 1 / Alfa)
Else
   F_Frechet_Media = ChrW(8734)
End If

End Function


' FUNCIÓN F_Frechet_Moda

Public Function F_Frechet_Moda(Alfa As Double) As Variant
' Calcula la moda de la distribución de Fréchet

If Alfa <= 0 Then
   F_Frechet_Moda = ChrW(8734)
   Exit Function
End If

F_Frechet_Moda = (Alfa / (1 + Alfa)) ^ (1 / Alfa)

End Function


' FUNCIÓN F_Frechet_DesvTip

Public Function F_Frechet_DesvTip(Alfa As Double) As Variant
' Calcula la desviación típica de la distribución de Fréchet
' Llama a la función F_Gamma
Dim a1 As Double, a2 As Double

If Alfa <= 0 Then
   F_Frechet_DesvTip = "Alfa debe ser > 0"
   Exit Function
End If

If Alfa <= 2 Then
   F_Frechet_DesvTip = ChrW(8734)
   Exit Function
End If

a1 = F_Gamma(1 - 1 / Alfa)
a2 = F_Gamma(1 - 2 / Alfa)

F_Frechet_DesvTip = Sqr(a2 - a1 * a1)

End Function


' FUNCIÓN F_Frechet_Asimetria

Public Function F_Frechet_Asimetria(Alfa As Double) As Variant
' Calcula coeficiente de asimetría (Fisher) de la distribución de Fréchet
' Llama a la función F_Gamma
Dim a1 As Double, a2 As Double, a3 As Double

If Alfa <= 0 Then
   F_Frechet_Asimetria = "Alfa debe ser > 0"
   Exit Function
End If

If Alfa <= 3 Then
   F_Frechet_Asimetria = ChrW(8734)
   Exit Function
End If

a1 = F_Gamma(1 - 1 / Alfa)
a2 = F_Gamma(1 - 2 / Alfa)
a3 = F_Gamma(1 - 3 / Alfa)

F_Frechet_Asimetria = a3 - 3 * a2 * a1 + 2 * a1 ^ 3
F_Frechet_Asimetria = F_Frechet_Asimetria / (a2 - a1 * a1) ^ (3 / 2)

End Function


' FUNCIÓN F_Frechet_Curtosis

Public Function F_Frechet_Curtosis(Alfa As Double) As Variant
' Calcula la curtosis de la distribución de Fréchet
' Llama a la función F_Gamma
Dim a1 As Double, a2 As Double, a3 As Double, a4 As Double

If Alfa <= 0 Then
   F_Frechet_Curtosis = "Alfa debe ser > 0"
   Exit Function
End If

If Alfa <= 4 Then
   F_Frechet_Curtosis = ChrW(8734)
   Exit Function
End If

a1 = F_Gamma(1 - 1 / Alfa)
a2 = F_Gamma(1 - 2 / Alfa)
a3 = F_Gamma(1 - 3 / Alfa)
a4 = F_Gamma(1 - 4 / Alfa)

F_Frechet_Curtosis = a4 - 4 * a3 * a1 + 3 * a2 * a2
F_Frechet_Curtosis = F_Frechet_Curtosis / (a2 - a1 * a1) ^ 2
F_Frechet_Curtosis = F_Frechet_Curtosis - 3

End Function


