
' FUNCIÓN DE DENSIDAD

Public Function D_M_B(x As Double, Sigma As Double) As Variant
' Esta función calcula la función de densidad de la distribución de Maxwell-Boltzmann
Dim Eps As Double, s2 As Double, s3 As Double
Dim R2Pi As Double

R2Pi = 0.797884560802865
Eps = 0.0000001

If Sigma <= Eps Then
   D_M_B = "Sigma debe ser > 0"
   Exit Function
End If

If x < 0 Then
   D_M_B = 0
   Exit Function
End If

s2 = Sigma * Sigma
s3 = s2 * Sigma

D_M_B = x * x / s3 * R2Pi * Exp(-x * x / 2 / s2)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_M_B(x As Double, Sigma As Double) As Variant
' Esta función calcula la función de distribución de la distribución de Maxwell-Boltzmann
' Llama a la función de error F_erf()
Dim Eps As Double, s2 As Double
Dim R2 As Double, R2Pi As Double
Dim a1 As Double, a2 As Double

R2Pi = 0.797884560802865
R2 = 1.4142135623731
Eps = 0.0000001

If Sigma <= Eps Then
   FD_M_B = "Sigma debe ser > 0"
   Exit Function
End If

If x <= 0 Then
   FD_M_B = 0
   Exit Function
End If

s2 = Sigma * Sigma
a1 = F_erf(x / Sigma / R2)
a2 = R2Pi * x / Sigma * Exp(-x * x / 2 / s2)

FD_M_B = a1 - a2

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_M_B_Inv(Probabilidad As Double, Sigma As Double) As Variant
' Esta función obtiene la inversa de la función de distribución de Maxwell-Boltzman
' Llama a la función Mi_ecuacion_Est

Dim x1 As Double, x2 As Double, Factor As Double
Dim Hecho As String, aa As Variant
Dim Ene As Double, Eps As Double

Eps = 0.0000001

If Sigma <= Eps Then
   F_M_B_Inv = "Sigma debe ser > 0"
   Exit Function
End If

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_M_B_Inv = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If Probabilidad < Eps Then
   F_M_B_Inv = 0
   Exit Function
End If

If Abs(Probabilidad - 1) < Eps Then
   F_M_B_Inv = ChrW(8734)
   Exit Function
End If

Factor = 10
Hecho = "No"

Do While Hecho = "No"
   x1 = -Factor * Sigma
   x2 = Factor * Sigma
   aa = Mi_ecuacion_Est("M_B", Probabilidad, x1, x2, Sigma, , , 0.000000001)
   If aa = "Rango Mal" Then
      ' El rango no contenía la raíz y lo aumentamos
      Factor = Factor + 1
   Else
      ' Se ha obtenido la raíz que se ha metido en la variable aa
      Hecho = "Si"
      F_M_B_Inv = aa
   End If
Loop

End Function


' FUNCIÓN F_M_B_Media

Public Function F_M_B_Media(Sigma As Double) As Variant
' Calcula la media de la distribución de Maxwell-Boltzmann
Dim Pi As Double

Pi = 3.14159265358979
If Sigma <= 0 Then
   F_M_B_Media = "Sigma debe ser > 0"
   Exit Function
End If

F_M_B_Media = 2 * Sigma * Sqr(2 / Pi)

End Function


' FUNCIÓN F_M_B_Moda

Public Function F_M_B_Moda(Sigma As Double) As Variant
' Calcula la moda de la distribución de Maxwell-Boltzmann

If Sigma <= 0 Then
   F_M_B_Moda = "Sigma debe ser > 0"
   Exit Function
End If

F_M_B_Moda = Sigma * Sqr(2)

End Function


' FUNCIÓN F_M_B_DesvTip

Public Function F_M_B_DesvTip(Sigma As Double) As Variant
' Calcula la desviación típica de la distribución de Maxwell-Boltzmann
Dim Pi As Double

Pi = 3.14159265358979
If Sigma <= 0 Then
   F_M_B_DesvTip = "Sigma debe ser > 0"
   Exit Function
End If

F_M_B_DesvTip = Sigma * Sqr((3 * Pi - 8) / Pi)

End Function


' FUNCIÓN F_M_B_Asimetria

Public Function F_M_B_Asimetria(Sigma As Double) As Variant
' Calcula coeficiente de asimetría (Fisher) de la distribución de Maxwell-Boltzmann
Dim Pi As Double

Pi = 3.14159265358979
If Sigma <= 0 Then
   F_M_B_Asimetria = "Sigma debe ser > 0"
   Exit Function
End If

F_M_B_Asimetria = 2 * Sqr(2) * (16 - 5 * Pi) / Sqr((3 * Pi - 8) ^ 3)

End Function


' FUNCIÓN F_M_B_Curtosis

Public Function F_M_B_Curtosis(Sigma As Double) As Variant
' Calcula la curtosis de la distribución de Maxwell-Boltzmann
Dim Pi As Double

Pi = 3.14159265358979
If Sigma <= 0 Then
   F_M_B_Curtosis = "Sigma debe ser > 0"
   Exit Function
End If

F_M_B_Curtosis = 3 + 4 * (-3 * Pi * Pi + 40 * Pi - 96) / (3 * Pi - 8) ^ 2

End Function


