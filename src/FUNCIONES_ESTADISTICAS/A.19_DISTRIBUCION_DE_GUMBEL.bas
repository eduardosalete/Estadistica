
' FUNCIÓN DE DENSIDAD

Public Function D_Gumbel(x As Double, Mu As Double, Beta As Double) As Variant
' Esta función calcula la función de densidad de la distribución de Gumbel
Dim z As Double

If Beta <= 0 Then
   D_Gumbel = "Beta debe ser >0"
   Exit Function
End If

z = (x - Mu) / Beta
D_Gumbel = 1 / Beta * Exp(-(z + Exp(-z)))

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Gumbel(x As Double, Mu As Double, Beta As Double) As Variant
' Esta función calcula la función de distribución de la distribución de Gumbel
Dim z As Double

If Beta <= 0 Then
   FD_Gumbel = "Beta debe ser >0"
   Exit Function
End If

z = (x - Mu) / Beta
FD_Gumbel = Exp(-Exp(-z))

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Gumbel_Inv(Probabilidad As Double, Mu As Double, Beta As Double) As Variant
' Esta función calcula la inversa de la función de distribución de la distribución de Gumbel
Dim z As Double
Dim Eps As Double

Eps = 0.0000001

If Beta <= 0 Then
   F_Gumbel_Inv = "Beta debe ser >0"
   Exit Function
End If

If Probabilidad < 0 Then
   F_Gumbel_Inv = "Probabilidad negativa"
   Exit Function
End If

If Probabilidad > 1 Then
   F_Gumbel_Inv = "Probabilidad > 1"
   Exit Function
End If

If Probabilidad < Eps Then
   F_Gumbel_Inv = "-" & ChrW(8734)
   Exit Function
End If

If Probabilidad >= 1 - Eps Then
   F_Gumbel_Inv = "+" & ChrW(8734)
   Exit Function
End If

z = -Log(-Log(Probabilidad))

F_Gumbel_Inv = Beta * z + Mu

End Function


' FUNCIÓN F_Gumbel_Media

Public Function F_Gumbel_Media(Mu As Double, Beta As Double) As Variant
' Calcula la media de la distribución de Gumbel
Dim Gamma As Double

Gamma = 0.577215664901533   '  Constante de Euler Mascheroni

If Beta <= 0 Then
   F_Gumbel_Media = "Beta debe ser >0"
   Exit Function
End If

F_Gumbel_Media = Mu + Beta * Gamma

End Function


' FUNCIÓN F_Gumbel_Moda

Public Function F_Gumbel_Moda(Mu As Double, Beta As Double) As Variant
' Calcula la moda de la distribución de Gumbel

If Beta <= 0 Then
   F_Gumbel_Moda = "Beta debe ser >0"
   Exit Function
End If

F_Gumbel_Moda = Mu

End Function


' FUNCIÓN F_Gumbel_Mediana

Public Function F_Gumbel_Mediana(Mu As Double, Beta As Double) As Variant
' Calcula la mediana de la distribución de Gumbel

If Beta <= 0 Then
   F_Gumbel_Mediana = "Beta debe ser >0"
   Exit Function
End If

F_Gumbel_Mediana = Mu - Beta * Log(Log(2))

End Function


' FUNCIÓN F_Gumbel_DesvTip

Public Function F_Gumbel_DesvTip(Mu As Double, Beta As Double) As Variant
' Calcula la desviación típica de la distribución de Gumbel
Dim Pi As Double

Pi = 3.14159265358979

If Beta <= 0 Then
   F_Gumbel_DesvTip = "Beta debe ser >0"
   Exit Function
End If

F_Gumbel_DesvTip = Pi * Beta / Sqr(6)

End Function


' FUNCIÓN F_Gumbel_Asimetria

Public Function F_Gumbel_Asimetria(Mu As Double, Beta As Double) As Variant
' Calcula el coeficiente de asimetría (Fisher) de la distribución de Gumbel
Dim Zeta_3 As Double

Zeta_3 = 1.20205690315959 ' Constante de Apéry (no se usa realmente,
                          ' porque la función devuelve una constante

If Beta <= 0 Then
   F_Gumbel_Asimetria = "Beta debe ser >0"
   Exit Function
End If
 
F_Gumbel_Asimetria = 1.13954709940464  ' 12*sqr(6)*Zeta_3/Pi^3

End Function


' FUNCIÓN F_Gumbel_Curtosis

Public Function F_Gumbel_Curtosis(Mu As Double, Beta As Double) As Variant
' Calcula la curtosis de la distribución de Gumbel

If Beta <= 0 Then
   F_Gumbel_Curtosis = "Beta debe ser >0"
   Exit Function
End If

F_Gumbel_Curtosis = 12 / 5 + 3

End Function


