
' FUNCIÓN DE DENSIDAD

Public Function D_Beta(x As Double, Alfa As Double, Beta As Double) As Variant
' Esta función calcula la función de densidad de la distribución Beta
' Llama a la función F_Beta
Dim Eps As Double

Eps = 0.0000001

If Alfa <= 0 Or Beta <= 0 Then
   D_Beta = "Alfa y Beta deben ser >0"
   Exit Function
End If

If x < 0 Or x > 1 Then
   D_Beta = 0
   Exit Function
End If

If Alfa < 1 And x < Eps Then
   ' Para prevenir infinitos si Alfa<1
   D_Beta = "+" & ChrW(8734)
   Exit Function
End If
If Beta < 1 And Abs(x - 1) < Eps Then
   ' Para prevenir infinitos si Beta<1
   D_Beta = "+" & ChrW(8734)
   Exit Function
End If

D_Beta = 1 / F_Beta(Alfa, Beta) * x ^ (Alfa - 1) * (1 - x) ^ (Beta - 1)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Beta(x As Double, Alfa As Double, Beta As Double) As Variant
' Esta función calcula la función de distribución de la distribución Beta
' Llama a la función F_Beta
' Llama a la función F_BetaI
Dim Eps As Double

Eps = 0.0000001

If Alfa <= 0 Or Beta <= 0 Then
   FD_Beta = "Alfa y Beta deben ser >0"
   Exit Function
End If

If x < Eps Then
   FD_Beta = 0
   Exit Function
End If

If x >= 1 Then
   FD_Beta = 1
   Exit Function
End If

FD_Beta = F_BetaI(Alfa, Beta, x, 1000) / F_Beta(Alfa, Beta)

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Beta_Inv(Probabilidad As Double, Alfa As Double, Beta As Double) As Variant
' Esta función obtiene la inversa de la función de distribución Beta
' Llama a la función FD_Beta
' Llama a la función Mi_ecuacion_Est

Dim x1 As Double, x2 As Double
Dim Hecho As String, aa As Variant
Dim Eps As Double
Dim F1 As Double, F2 As Double

Eps = 0.000000001

If Alfa <= 0 Or Beta <= 0 Then
   F_Beta_Inv = "Alfa y Beta deben ser >0"
   Exit Function
End If

If Probabilidad <= Eps Then
   F_Beta_Inv = 0
   Exit Function
End If

If Abs(Probabilidad - 1) <= Eps Then
   F_Beta_Inv = 1
   Exit Function
End If

x1 = Eps: F1 = FD_Beta(x1, Alfa, Beta)
x2 = 0.999999: F2 = FD_Beta(x2, Alfa, Beta)
Hecho = "No"
Do While Hecho = "No"
   If x1 < 0 Then x1 = 0
   aa = Mi_ecuacion_Est("Beta", Probabilidad, x1, x2, Alfa, Beta, , Eps)
   If aa = "Rango Mal" Then
      ' El rango no contenía la raíz y lo aumentamos
      If Probabilidad < 0.5 Then
         F_Beta_Inv = x1
      Else
         F_Beta_Inv = x2
      End If
      Exit Function
   Else
      ' Se ha obtenido la raíz que se ha metido en la variable aa
      Hecho = "Si"
      F_Beta_Inv = aa
   End If
Loop

End Function


' FUNCIÓN F_Beta_Media

Public Function F_Beta_Media(Alfa As Double, Beta As Double) As Variant
' Calcula la media de la distribución Beta

If Alfa <= 0 Or Beta <= 0 Then
   F_Beta_Media = "Alfa y Beta deben ser >0"
   Exit Function
End If

F_Beta_Media = Alfa / (Alfa + Beta)

End Function


' FUNCIÓN F_Beta_Moda

Public Function F_Beta_Moda(Alfa As Double, Beta As Double) As Variant
' Calcula la moda de la distribución Beta
If Alfa <= 0 Or Beta <= 0 Then
   F_Beta_Moda = "Alfa y Beta deben ser >0"
   Exit Function
End If

If Beta > 1 And Alfa < 1 Then
   F_Beta_Moda = 0
   Exit Function
End If

If Beta < 1 And Alfa > 1 Then
   F_Beta_Moda = 1
   Exit Function
End If

If Beta < 1 And Alfa < 1 Then
   F_Beta_Moda = "Indeterminada"
   Exit Function
End If

F_Beta_Moda = (Alfa - 1) / (Alfa + Beta - 2)

End Function


' FUNCIÓN F_Beta_DesvTip

Public Function F_Beta_DesvTip(Alfa As Double, Beta As Double) As Variant
' Calcula la desviación típica de la distribución Beta

If Alfa <= 0 Or Beta <= 0 Then
   F_Beta_DesvTip = "Alfa y Beta deben ser >0"
   Exit Function
End If

F_Beta_DesvTip = Alfa * Beta / (Alfa + Beta) ^ 2 / (Alfa + Beta + 1)
F_Beta_DesvTip = Sqr(F_Beta_DesvTip)

End Function


' FUNCIÓN F_Beta_Asimetria

Public Function F_Beta_Asimetria(Alfa As Double, Beta As Double) As Variant
' Calcula el coeficiente de asimetría (Fisher) de la distribución Beta

If Alfa <= 0 Or Beta <= 0 Then
   F_Beta_Asimetria = "Alfa y Beta deben ser >0"
   Exit Function
End If

F_Beta_Asimetria = 2 * (Beta - Alfa) * Sqr(Alfa + Beta + 1)
F_Beta_Asimetria = F_Beta_Asimetria / (Alfa + Beta + 2)
F_Beta_Asimetria = F_Beta_Asimetria / Sqr(Alfa * Beta)

End Function


' FUNCIÓN F_Beta_Curtosis

Public Function F_Beta_Curtosis(Alfa As Double, Beta As Double) As Variant
' Calcula la curtosis de la distribución Beta

If Alfa <= 0 Or Beta <= 0 Then
   F_Beta_Curtosis = "Alfa y Beta deben ser >0"
   Exit Function
End If

F_Beta_Curtosis = 6 * (Alfa - Beta) ^ 2 * (Alfa + Beta + 1)
F_Beta_Curtosis = F_Beta_Curtosis - 6 * Alfa * Beta * (Alfa + Beta + 2)
F_Beta_Curtosis = F_Beta_Curtosis / Alfa / Beta
F_Beta_Curtosis = F_Beta_Curtosis / (Alfa + Beta + 2)
F_Beta_Curtosis = F_Beta_Curtosis / (Alfa + Beta + 3)
F_Beta_Curtosis = F_Beta_Curtosis + 3

End Function


