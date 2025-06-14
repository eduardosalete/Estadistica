
' FUNCIÓN DE DENSIDAD

Public Function D_Pert(x As Double, a As Double, b As Double, c As Double) As Variant
' Esta función calcula la función de densidad de la distribución PERT
' Llama a la función F_Beta(x,y)

Dim BB As Double
Dim Alfa As Double, Beta As Double

If a >= b Or b >= c Then
   D_Pert = "Los parámetros deben ser a<b<c"
   Exit Function
End If

If x < a Or x > c Then
   D_Pert = 0
   Exit Function
End If

Alfa = 1 + 4 * (b - a) / (c - a)
Beta = 1 + 4 * (c - b) / (c - a)
BB = F_Beta(Alfa, Beta) * (c - a) ^ (Alfa + Beta - 1)

D_Pert = 1 / BB * (x - a) ^ (Alfa - 1) * (c - x) ^ (Beta - 1)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Pert(x As Double, a As Double, b As Double, c As Double) As Variant
' Esta función calcula la función de distribución de la distribución PERT
' Llama a la función F_Beta
' Llama a la función F_BetaI
Dim Alfa As Double, Beta As Double
Dim z As Double

If a >= b Or b >= c Then
   FD_Pert = "Los parámetros deben ser a<b<c"
   Exit Function
End If

If x < a Then
   FD_Pert = 0
   Exit Function
End If

If x > c Then
   FD_Pert = 1
   Exit Function
End If

Alfa = 1 + 4 * (b - a) / (c - a)
Beta = 1 + 4 * (c - b) / (c - a)
z = (x - a) / (c - a)

FD_Pert = F_BetaI(Alfa, Beta, z, 200) / F_Beta(Alfa, Beta)

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Pert_Inv(Probabilidad As Double, a As Double, b As Double, c As Double) As Variant
' Esta función obtiene la inversa de la función de distribución PERT
' Llama a la función F_Beta_Inv

Dim Eps As Double
Dim Alfa As Double, Beta As Double
Dim z As Double

Eps = 0.000000001

If a >= b Or b >= c Then
   F_Pert_Inv = "Los parámetros deben ser a<b<c"
   Exit Function
End If

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_Pert_Inv = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

Alfa = 1 + 4 * (b - a) / (c - a)
Beta = 1 + 4 * (c - b) / (c - a)

' Obtenemos z = (x - a) / (c - a)
z = F_Beta_Inv(Probabilidad, Alfa, Beta)
F_Pert_Inv = (c - a) * z + a

End Function


' FUNCIÓN F_Pert_Media

Public Function F_Pert_Media(a As Double, b As Double, c As Double) As Variant
' Calcula la media de la distribución PERT

If a >= b Or b >= c Then
   F_Pert_Media = "Los parámetros deben ser a<b<c"
   Exit Function
End If

F_Pert_Media = (a + 4 * b + c) / 6

End Function


' FUNCIÓN F_Pert_Moda

Public Function F_Pert_Moda(a As Double, b As Double, c As Double) As Variant
' Calcula la moda de la distribución PERT
If a >= b Or b >= c Then
   F_Pert_Moda = "Los parámetros deben ser a<b<c"
   Exit Function
End If

F_Pert_Moda = b

End Function


' FUNCIÓN F_Pert_DesvTip

Public Function F_Pert_DesvTip(a As Double, b As Double, c As Double) As Variant
' Calcula la desviación típica de la distribución PERT
' Llama a la función F_Pert_Media
Dim Mu As Double

If a >= b Or b >= c Then
   F_Pert_DesvTip = "Los parámetros deben ser a<b<c"
   Exit Function
End If

Mu = F_Pert_Media(a, b, c)
F_Pert_DesvTip = (Mu - a) * (c - Mu) / 7
F_Pert_DesvTip = Sqr(F_Pert_DesvTip)

End Function


' FUNCIÓN F_Pert_Asimetria

Public Function F_Pert_Asimetria(a As Double, b As Double, c As Double) As Variant
' Calcula el coeficiente de asimetría (Fisher) de la distribución PERT
' Llama a la función F_Beta_Asimetria
Dim Alfa As Double, Beta As Double

If a >= b Or b >= c Then
   F_Pert_Asimetria = "Los parámetros deben ser a<b<c"
   Exit Function
End If

Alfa = 1 + 4 * (b - a) / (c - a)
Beta = 1 + 4 * (c - b) / (c - a)

F_Pert_Asimetria = F_Beta_Asimetria(Alfa, Beta)

End Function


' FUNCIÓN F_Pert_Curtosis

Public Function F_Pert_Curtosis(a As Double, b As Double, c As Double) As Variant
' Calcula la curtosis de la distribución PERT
' Llama a la función F_Beta_Curtosis
Dim Alfa As Double, Beta As Double

If a >= b Or b >= c Then
   F_Pert_Curtosis = "Los parámetros deben ser a<b<c"
   Exit Function
End If

Alfa = 1 + 4 * (b - a) / (c - a)
Beta = 1 + 4 * (c - b) / (c - a)
F_Pert_Curtosis = F_Beta_Curtosis(Alfa, Beta)

End Function


