
' FUNCIÓN DE DENSIDAD

Public Function D_Weibull(x As Double, Lambda As Double, k As Double) As Variant
' Esta función calcula la función de densidad de la distribución de Weibull
Dim Eps As Double

Eps = 0.0000001

If k <= 0 Or Lambda <= 0 Then
   D_Weibull = "k y Lambda deben ser >0"
   Exit Function
End If

If x < 0 Then
   D_Weibull = 0
   Exit Function
End If

If k < 1 And x < Eps Then
   ' Para prevenir infinitos si k<1
   D_Weibull = "+" & ChrW(8734)
   Exit Function
End If

D_Weibull = k / Lambda * (x / Lambda) ^ (k - 1) * Exp(-(x / Lambda) ^ k)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Weibull(x As Double, Lambda As Double, k As Double) As Variant
' Esta función calcula la función de distribución de la distribución de Weibull
If k <= 0 Or Lambda <= 0 Then
   FD_Weibull = "k y Lambda deben ser >0"
   Exit Function
End If

If x < 0 Then
   FD_Weibull = 0
   Exit Function
End If

FD_Weibull = 1 - Exp(-(x / Lambda) ^ k)

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Weibull_Inv(Probabilidad As Double, Lambda As Double, k As Double) As Variant
' Esta función calcula la inversa de la función de distribución de la distribución de Weibull
Dim Eps As Double

Eps = 0.0000001

If k <= 0 Or Lambda <= 0 Then
   F_Weibull_Inv = "k y Lambda deben ser >0"
   Exit Function
End If

If Probabilidad < 0 Then
   F_Weibull_Inv = "Probabilidad negativa"
   Exit Function
End If

If Probabilidad > 1 Then
   F_Weibull_Inv = "Probabilidad > 1"
   Exit Function
End If

If Probabilidad >= 1 - Eps Then
   F_Weibull_Inv = "+" & ChrW(8734)
   Exit Function
End If

F_Weibull_Inv = Lambda * (-Log(1 - Probabilidad)) ^ (1 / k)

End Function


' FUNCIÓN F_Weibull_Media

Public Function F_Weibull_Media(Lambda As Double, k As Double) As Variant
' Calcula la media de la distribución de Weibull
' Llama a la función F_Gamma

If k <= 0 Or Lambda <= 0 Then
   F_Weibull_Media = "k y Lambda deben ser >0"
   Exit Function
End If

F_Weibull_Media = Lambda * F_Gamma(1 + 1 / k)

End Function


' FUNCIÓN F_Weibull_Moda

Public Function F_Weibull_Moda(Lambda As Double, k As Double) As Variant
' Calcula la moda de la distribución de Weibull

If k <= 0 Or Lambda <= 0 Then
   F_Weibull_Moda = "k y Lambda deben ser >0"
   Exit Function
End If

If k > 1 Then
   F_Weibull_Moda = Lambda * (1 - 1 / k) ^ (1 / k)
Else
   F_Weibull_Moda = 0
End If

End Function


' FUNCIÓN F_Weibull_Mediana

Public Function F_Weibull_Mediana(Lambda As Double, k As Double) As Variant
' Calcula la mediana de la distribución de Weibull

If k <= 0 Or Lambda <= 0 Then
   F_Weibull_Mediana = "k y Lambda deben ser >0"
   Exit Function
End If

F_Weibull_Mediana = Lambda * 2 ^ (1 / k)

End Function


' FUNCIÓN F_Weibull_DesvTip

Public Function F_Weibull_DesvTip(Lambda As Double, k As Double) As Variant
' Calcula la desviación típica de la distribución de Weibull
' Llama a la función F_Gamma

If k <= 0 Or Lambda <= 0 Then
   F_Weibull_DesvTip = "k y Lambda deben ser >0"
   Exit Function
End If

F_Weibull_DesvTip = Lambda ^ 2 * (F_Gamma(1 + 2 / k) - (F_Gamma(1 + 1 / k)) ^ 2)
F_Weibull_DesvTip = Sqr(F_Weibull_DesvTip)

End Function


' FUNCIÓN F_Weibull_Asimetria

Public Function F_Weibull_Asimetria(Lambda As Double, k As Double) As Variant
' Calcula el coeficiente de asimetría (Fisher) de la distribución de Weibull
' Llama a la función F_Gamma
' Llama a la función F_Weibull_Media
' Llama a la función F_Weibull_DesvTip
Dim Mu As Double, Sigma As Double

If k <= 0 Or Lambda <= 0 Then
   F_Weibull_Asimetria = "k y Lambda deben ser >0"
   Exit Function
End If

Mu = F_Weibull_Media(Lambda, k)
Sigma = F_Weibull_DesvTip(Lambda, k)

F_Weibull_Asimetria = F_Gamma(1 + 3 / k) * Lambda ^ 3
F_Weibull_Asimetria = F_Weibull_Asimetria - 3 * Mu * Sigma ^ 2 - Mu ^ 3
F_Weibull_Asimetria = F_Weibull_Asimetria / Sigma ^ 3

End Function


' FUNCIÓN F_Weibull_Curtosis

Public Function F_Weibull_Curtosis(Lambda As Double, k As Double) As Variant
' Calcula la curtosis de la distribución de Weibull
' Llama a la función F_Gamma
' Llama a la función F_Weibull_Media
' Llama a la función F_Weibull_DesvTip
' Llama a la función F_Weibull_Asimetria
Dim Mu As Double, Sigma As Double, Gamma As Double

If k <= 0 Or Lambda <= 0 Then
   F_Weibull_Curtosis = "k y Lambda deben ser >0"
   Exit Function
End If
Mu = F_Weibull_Media(Lambda, k)
Sigma = F_Weibull_DesvTip(Lambda, k)
Gamma = F_Weibull_Asimetria(Lambda, k)

F_Weibull_Curtosis = F_Gamma(1 + 4 / k) * Lambda ^ 4
F_Weibull_Curtosis = F_Weibull_Curtosis - 4 * Mu * Sigma ^ 3 * Gamma
F_Weibull_Curtosis = F_Weibull_Curtosis - 3 * Sigma ^ 4
F_Weibull_Curtosis = F_Weibull_Curtosis - 6 * Mu ^ 2 * Sigma ^ 2
F_Weibull_Curtosis = F_Weibull_Curtosis - Mu ^ 4
F_Weibull_Curtosis = F_Weibull_Curtosis / Sigma ^ 4
F_Weibull_Curtosis = F_Weibull_Curtosis + 3

End Function


