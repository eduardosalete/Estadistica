
' FUNCIÓN DE DENSIDAD

Public Function D_Gamma(xx As Double, Alfa As Double, Beta As Double, Optional Gamma_a As Double = 0) As Variant
' Esta función calcula la función de densidad de la distribución Gamma
' Llama a la función F_Gamma

Dim Eps As Double, x As Double

Eps = 0.0000001

If Alfa <= 0 Or Beta <= 0 Then
   D_Gamma = "Alfa y Beta deben ser positivos"
   Exit Function
End If

If Gamma_a <= Eps Then
   ' Gamma_a puede aportarse para ahorrar cálculo (especialmente si se
   ' va a repetir el cálculo muchas veces para el mismo valor de Alfa).
   ' Si no se proporciona lo calcula la función
   Gamma_a = F_Gamma(Alfa)
End If

If xx <= 0 Then
   x = 0
Else
   x = xx
End If

D_Gamma = Beta ^ (-Alfa) * x ^ (Alfa - 1) * Exp(-x / Beta) / Gamma_a

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Gamma(xx As Double, Alfa As Double, Beta As Double, Optional Gamma_a As Double = 0) As Variant
' Esta función calcula la función de distribución de la distribución Gamma
' Llama a la función F_Gamma
' Llama a la función F_Gamma_Inf

Dim Eps As Double, x As Double

Eps = 0.0000001

If Alfa <= 0 Or Beta <= 0 Then
   FD_Gamma = "Alfa y Beta deben ser positivos"
   Exit Function
End If

If Gamma_a <= Eps Then
   ' Gamma_a puede aportarse para ahorrar cálculo (especialmente si se
   ' va a repetir el cálculo muchas veces para el mismo valor de Alfa).
   ' Si no se proporciona lo calcula la función
   Gamma_a = F_Gamma(Alfa)
End If

If xx <= 0 Then
   x = 0
Else
   x = xx
End If

FD_Gamma = F_Gamma_Inf(Alfa, x / Beta) / Gamma_a

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Gamma_Inv(Probabilidad As Double, Alfa As Double, Beta As Double) As Variant
' Esta función obtiene la inversa de la función de distribución Gamma
' Llama a la función F_Gamma
' Llama a la función Mi_ecuacion_Est

Dim Mu As Double, Sigma As Double, G_Alfa As Double
Dim x1 As Double, x2 As Double, Factor As Double
Dim Hecho As String, aa As Variant
Dim Eps As Double

Eps = 0.0000001

If Alfa <= 0 Or Beta <= 0 Then
   F_Gamma_Inv = "Alfa y Beta deben ser positivos"
   Exit Function
End If

If Probabilidad <= Eps Then
   F_Gamma_Inv = 0
   Exit Function
End If

If Abs(Probabilidad - 1) <= Eps Then
   F_Gamma_Inv = ChrW(8734)
   Exit Function
End If

Mu = Alfa * Beta
Sigma = Sqr(Alfa) * Beta
G_Alfa = F_Gamma(Alfa)
Factor = 5

Hecho = "No"
Do While Hecho = "No"
   x1 = Mu - Factor * Sigma
   x2 = Mu + Factor * Sigma
   If x1 < 0 Then x1 = 0
   aa = Mi_ecuacion_Est("Gamma", Probabilidad, x1, x2, Alfa, Beta, G_Alfa, 0.000000001)
   If aa = "Rango Mal" Then
      ' El rango no contenía la raíz y lo aumentamos
      Factor = Factor + 1
   Else
      ' Se ha obtenido la raíz que se ha metido en la variable aa
      Hecho = "Si"
      F_Gamma_Inv = aa
   End If
Loop

End Function


