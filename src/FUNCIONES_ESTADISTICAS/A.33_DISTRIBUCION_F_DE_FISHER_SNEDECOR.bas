
' FUNCIÓN DE DENSIDAD

Public Function D_F_Snedecor(x As Double, n1 As Integer, n2 As Integer) As Variant
' Calcula la función de densidad de la distribución F de Snedecor
' con n1 y n2 grados de libertad
' Llama a la función F_Beta
Dim Beta As Double, Factor As Double, Eps As Double

Eps = 0.0000001

If n1 <= 0 Or n2 <= 0 Then
   D_F_Snedecor = "Los g.d.l. n1 y n2 deben ser >0"
   Exit Function
End If

If x < 0 Then
   D_F_Snedecor = "x debe ser >0"
   Exit Function
End If

If Abs(x) < Eps And n1 = 1 Then
   D_F_Snedecor = ChrW(8734)
   Exit Function
End If

If Abs(x) < Eps And n1 = 2 Then
   D_F_Snedecor = 1
   Exit Function
End If

If Abs(x) < Eps Then
   D_F_Snedecor = 0
   Exit Function
End If

Beta = F_Beta(n1 / 2, n2 / 2)
Factor = n1 * x / (n1 * x + n2)

D_F_Snedecor = 1 / Beta * Factor ^ (n1 / 2) * (1 - Factor) ^ (n2 / 2) / x

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_F_Snedecor(x As Double, n1 As Integer, n2 As Integer) As Variant
' Calcula la función de distribución de la distribución F de Snedecor
' con n1 y n2 grados de libertad
' Llama a la función F_Beta
' Llama a la función F_BetaI

Dim z As Double
Dim Beta As Double, Eps As Double

Eps = 0.0000001

If n1 <= 0 Or n2 <= 0 Then
   FD_F_Snedecor = "Los g.d.l. n1 y n2 deben ser >0"
   Exit Function
End If

If x < 0 Then
   FD_F_Snedecor = "x debe ser >0"
   Exit Function
End If

If Abs(x) < Eps Then
   FD_F_Snedecor = 0
   Exit Function
End If

z = n1 * x / (n1 * x + n2)
Beta = F_Beta(n1 / 2, n2 / 2)

FD_F_Snedecor = 1 / Beta * F_BetaI(n1 / 2, n2 / 2, z, 1000)

End Function


' FUNCIÓN DE DISTRIBUCIÓN (RELACIÓN BETA INCOMPLETA – HIPERGEOMÉTRICA GAUSS)

Public Function FD_F_SnedecorHG(x As Double, n1 As Integer, n2 As Integer) As Variant
' Calcula la función de distribución de la distribución F de Snedecor
' con n1 y n2 grados de libertad
' Utiliza la relación de la función Beta Incompleta
' con la función Hipergeométrica de Gauss:
' B(a,b,x)=a^(-1) x^a F(a,1-b,a+1|x)
' Llama a la función F_Beta
' Llama a la función F_HG_Gauss

Dim z As Double, Factor1 As Double, Factor2 As Double
Dim Beta As Double, Eps As Double
Dim a As Double, b As Double, c As Double

Eps = 0.0000001

If n1 <= 0 Or n2 <= 0 Then
   FD_F_SnedecorHG = "Los g.d.l. n1 y n2 deben ser >0"
   Exit Function
End If

If x < 0 Then
   FD_F_SnedecorHG = "x debe ser >0"
   Exit Function
End If

If Abs(x) < Eps Then
   FD_F_SnedecorHG = 0
   Exit Function
End If

z = n1 * x / (n1 * x + n2)

a = n1 / 2: b = 1 - n2 / 2: c = 1 + n1 / 2
Factor1 = 2 / n1 * z ^ (n1 / 2)
Factor2 = F_HG_Gauss(a, b, c, z)
Beta = F_Beta(n1 / 2, n2 / 2)

FD_F_SnedecorHG = 1 / Beta * Factor1 * Factor2

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_F_Snedecor_Inv(Probabilidad As Double, n1 As Integer, n2 As Integer) As Variant
' Esta función obtiene la inversa de la función de distribución F de Snedecor
' Llama a la función Mi_ecuacion_Est

Dim Sigma As Double
Dim x1 As Double, x2 As Double, Factor As Double
Dim Hecho As String, aa As Variant
Dim Ene1 As Double, Ene2 As Double
Dim Eps As Double
Dim Cambio As String, p As Double

Eps = 0.0000001
Cambio = "No"

If n1 <= 0 Or n2 <= 0 Then
   F_F_Snedecor_Inv = "Los g.d.l. n1 y n2 deben ser >0"
   Exit Function
End If

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_F_Snedecor_Inv = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If Probabilidad < Eps Then
   F_F_Snedecor_Inv = 0
   Exit Function
End If

If Abs(Probabilidad - 1) < Eps Then
   F_F_Snedecor_Inv = ChrW(8734)
   Exit Function
End If

If Probabilidad <= 0.5 Then
   p = Probabilidad
   Ene1 = n1                  ' Pasamos a números reales los parámetros
   Ene2 = n2
Else
   p = 1 - Probabilidad
   Ene1 = n2                  ' Pasamos a números reales los parámetros
   Ene2 = n1
   Cambio = "Si"
End If

If Ene2 > 4 Then
   Sigma = 2 * Ene1 ^ 2 * (Ene1 + Ene2 - 2) / Ene1 / (Ene2 - 2) ^ 2 / (Ene2 - 4)
Else
   Sigma = 1
End If

Factor = 10
Hecho = "No"

Do While Hecho = "No"
   x1 = 0
   x2 = Factor * Sigma

   aa = Mi_ecuacion_Est("F_Snedecor", p, x1, x2, Ene1, Ene2, , 0.000000001)
   If aa = "Rango Mal" Then
      ' El rango no contenía la raíz y lo aumentamos
      Factor = Factor + 1
   Else
      ' Se ha obtenido la raíz que se ha metido en la variable aa
      Hecho = "Si"
      F_F_Snedecor_Inv = aa
   End If
Loop

If Cambio = "Si" Then
   F_F_Snedecor_Inv = 1 / F_F_Snedecor_Inv
End If

End Function


' FUNCIÓN F_Snedecor_Media

Public Function F_Snedecor_Media(n1 As Integer, n2 As Integer) As Variant
' Calcula la media de la distribución F de Snedecor
If n1 <= 0 Or n2 <= 0 Then
   F_Snedecor_Media = "Los g.d.l. n1 y n2 deben ser >0"
   Exit Function
End If
If n2 <= 2 Then
   F_Snedecor_Media = "Indeterminado"
Else
   F_Snedecor_Media = n2 / (n2 - 2)
End If

End Function


' FUNCIÓN F_Snedecor_Moda

Public Function F_Snedecor_Moda(n1 As Integer, n2 As Integer) As Variant
' Calcula la moda de la distribución F de Snedecor

If n1 <= 0 Or n2 <= 0 Then
   F_Snedecor_Moda = "Los g.d.l. n1 y n2 deben ser >0"
   Exit Function
End If

If n2 <= 2 Then
   F_Snedecor_Moda = "Indeterminado"
Else
   F_Snedecor_Moda = (n1 - 2) * n2 / n1 / (n2 + 2)
End If

End Function


' FUNCIÓN F_Snedecor_DesvTip

Public Function F_Snedecor_DesvTip(n1 As Integer, n2 As Integer) As Variant
' Calcula la desviación típica de la distribución F de Snedecor

If n1 <= 0 Or n2 <= 0 Then
   F_Snedecor_DesvTip = "Los g.d.l. n1 y n2 deben ser >0"
   Exit Function
End If
If n2 <= 4 Then
   F_Snedecor_DesvTip = "Indeterminado"
Else
   F_Snedecor_DesvTip = 2 * n2 ^ 2 * (n2 + n1 - 2) / (n2 - 4) / (n2 - 2) ^ 2 / n1
   F_Snedecor_DesvTip = Sqr(F_Snedecor_DesvTip)
End If

End Function


' FUNCIÓN F_Snedecor_Asimetria

Public Function F_Snedecor_Asimetria(n1 As Integer, n2 As Integer) As Variant
' Calcula coeficiente de asimetría (Fisher) de la distribución F de Snedecor
Dim R2 As Double

R2 = 1.4142135623731

If n1 <= 0 Or n2 <= 0 Then
   F_Snedecor_Asimetria = "Los g.d.l. n1 y n2 deben ser >0"
   Exit Function
End If
If n2 <= 6 Then
   F_Snedecor_Asimetria = "Indeterminado"
Else
   F_Snedecor_Asimetria = 2 * R2 * (2 * n1 + n2 - 2) * Sqr(n2 - 4)
   F_Snedecor_Asimetria = F_Snedecor_Asimetria / (n2 - 6) / Sqr(n1 * (n1 + n2 - 2))
End If

End Function


' FUNCIÓN F_Snedecor_Curtosis

Public Function F_Snedecor_Curtosis(n1 As Integer, n2 As Integer) As Variant
' Calcula la curtosis de la distribución F de Snedecor

If n1 <= 0 Or n2 <= 0 Then
   F_Snedecor_Curtosis = "Los g.d.l. n1 y n2 deben ser >0"
   Exit Function
End If
If n2 <= 8 Then
   F_Snedecor_Curtosis = "Indeterminado"
Else
   F_Snedecor_Curtosis = n1 * (5 * n2 - 22) * (n1 + n2 - 2)
   F_Snedecor_Curtosis = F_Snedecor_Curtosis + (n2 - 4) * (n2 - 2) ^ 2
   F_Snedecor_Curtosis = 12 * F_Snedecor_Curtosis
   F_Snedecor_Curtosis = 3 + F_Snedecor_Curtosis / n1 / (n2 - 6) / (n2 - 8) / (n1 + n2 - 2)
End If

End Function


