
' FUNCIÓN DE DENSIDAD

Public Function D_Pareto(x As Double, Alfa As Double, x0 As Double) As Variant
' Esta función calcula la función de densidad de la distribución de Pareto

If Alfa <= 0 Or x0 <= 0 Then
   D_Pareto = "Alfa y x0 deben ser >0"
   Exit Function
End If

If x < x0 Then
   D_Pareto = 0
   Exit Function
End If

D_Pareto = Alfa * x0 ^ Alfa / x ^ (Alfa + 1)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Pareto(x As Double, Alfa As Double, x0 As Double) As Variant
' Esta función calcula la función de distribución de la distribución de Pareto

If Alfa <= 0 Or x0 <= 0 Then
   FD_Pareto = "Alfa y x0 deben ser >0"
   Exit Function
End If

If x < x0 Then
   FD_Pareto = 0
   Exit Function
End If

FD_Pareto = 1 - (x0 / x) ^ Alfa

End Function


' COMPLEMENTO A LA FUNCIÓN DE DISTRIBUCIÓN

Public Function CFD_Pareto(x As Double, Alfa As Double, x0 As Double) As Variant
' Esta función calcula el complemento a la función de distribución de la distribución de Pareto

If Alfa <= 0 Or x0 <= 0 Then
   CFD_Pareto = "Alfa y x0 deben ser >0"
   Exit Function
End If

If x < x0 Then
   CFD_Pareto = 1
   Exit Function
End If

CFD_Pareto = (x0 / x) ^ Alfa

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Pareto_Inv(Probabilidad As Double, Alfa As Double, x0 As Double) As Variant
' Esta función calcula la inversa de la función de distribución de la distribución de Pareto
Dim Eps As Double
Eps = 0.0000001

If Alfa <= 0 Or x0 <= 0 Then
   F_Pareto_Inv = "Alfa y x0 deben ser >0"
   Exit Function
End If

If Probabilidad < 0 Then
   F_Pareto_Inv = "Probabilidad negativa"
   Exit Function
End If

If Probabilidad > 1 Then
   F_Pareto_Inv = "Probabilidad > 1"
   Exit Function
End If

If Probabilidad >= 1 - Eps Then
   F_Pareto_Inv = "+" & ChrW(8734)
   Exit Function
End If

F_Pareto_Inv = x0 * (1 - Probabilidad) ^ (-1 / Alfa)

End Function


' FUNCIÓN F_Pareto_Media

Public Function F_Pareto_Media(Alfa As Double, x0 As Double) As Variant
' Calcula la media de la distribución de Pareto

If Alfa <= 0 Or x0 <= 0 Then
   F_Pareto_Media = "Alfa y x0 deben ser >0"
   Exit Function
End If
If Alfa <= 1 Then
    F_Pareto_Media = "Indeterminada"
    Exit Function
End If

F_Pareto_Media = Alfa * x0 / (Alfa - 1)

End Function


' FUNCIÓN F_Pareto_Moda

Public Function F_Pareto_Moda(Alfa As Double, x0 As Double) As Variant
' Calcula la moda de la distribución de Pareto

If Alfa <= 0 Or x0 <= 0 Then
   F_Pareto_Moda = "Alfa y x0 deben ser >0"
   Exit Function
End If

F_Pareto_Moda = x0

End Function


' FUNCIÓN F_Pareto_Mediana

Public Function F_Pareto_Mediana(Alfa As Double, x0 As Double) As Variant
' Calcula la mediana de la distribución de Pareto

If Alfa <= 0 Or x0 <= 0 Then
   F_Pareto_Mediana = "Alfa y x0 deben ser >0"
   Exit Function
End If

F_Pareto_Mediana = x0 * 2 ^ (1 / Alfa)

End Function


' FUNCIÓN F_Pareto_DesvTip

Public Function F_Pareto_DesvTip(Alfa As Double, x0 As Double) As Variant
' Calcula la desviación típica de la distribución de Pareto

If Alfa <= 0 Or x0 <= 0 Then
   F_Pareto_DesvTip = "Alfa y x0 deben ser >0"
   Exit Function
End If
If Alfa <= 2 Then
    F_Pareto_DesvTip = "Indeterminada"
    Exit Function
End If

F_Pareto_DesvTip = Sqr(Alfa * x0 ^ 2 / (Alfa - 1) ^ 2 / (Alfa - 2))

End Function


' FUNCIÓN F_Pareto_Asimetria

Public Function F_Pareto_Asimetria(Alfa As Double, x0 As Double) As Variant
' Calcula el coeficiente de asimetría (Fisher) de la distribución de Pareto

If Alfa <= 0 Or x0 <= 0 Then
   F_Pareto_Asimetria = "Alfa y x0 deben ser >0"
   Exit Function
End If
If Alfa <= 3 Then
    F_Pareto_Asimetria = "Indeterminada"
    Exit Function
End If

F_Pareto_Asimetria = 2 * (1 + Alfa) / (Alfa - 3) * Sqr((Alfa - 2) / Alfa)

End Function


' FUNCIÓN F_Pareto_Curtosis

Public Function F_Pareto_Curtosis(Alfa As Double, x0 As Double) As Variant
' Calcula la curtosis de la distribución de Pareto

If Alfa <= 0 Or x0 <= 0 Then
   F_Pareto_Curtosis = "Alfa y x0 deben ser >0"
   Exit Function
End If
If Alfa <= 4 Then
    F_Pareto_Curtosis = "Indeterminada"
    Exit Function
End If

F_Pareto_Curtosis = 6 * (Alfa ^ 3 + Alfa ^ 2 - 6 * Alfa - 2)
F_Pareto_Curtosis = 3 + F_Pareto_Curtosis / Alfa / (Alfa - 3) / (Alfa - 4)

End Function


