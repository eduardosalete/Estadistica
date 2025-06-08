
' FUNCIÓN DE DENSIDAD

Public Function D_Rayleigh(x As Double, Sigma As Double) As Variant
' Esta función calcula la función de densidad de la distribución de Rayleigh
Dim Eps As Double, s2 As Double

Eps = 0.0000001
If Sigma <= Eps Then
   D_Rayleigh = "Sigma debe ser > 0"
   Exit Function
End If

If x < 0 Then
   D_Rayleigh = 0
   Exit Function
End If

s2 = Sigma * Sigma
D_Rayleigh = x / s2 * Exp(-x * x / 2 / s2)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Rayleigh(x As Double, Sigma As Double) As Variant
' Esta función calcula la función de distribución de la distribución de Rayleigh
Dim Eps As Double, s2 As Double

Eps = 0.0000001
If Sigma <= Eps Then
   FD_Rayleigh = "Sigma debe ser > 0"
   Exit Function
End If

If x < 0 Then
   FD_Rayleigh = 0
   Exit Function
End If

s2 = Sigma * Sigma
FD_Rayleigh = 1 - Exp(-x * x / 2 / s2)

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Rayleigh_Inv(Probabilidad As Double, Sigma As Double) As Variant
' Esta función obtiene la inversa de la función de distribución de Rayleigh
Dim Eps As Double

Eps = 0.0000001
If Sigma <= Eps Then
   F_Rayleigh_Inv = "Sigma debe ser > 0"
   Exit Function
End If

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_Rayleigh_Inv = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If Abs(Probabilidad - 1) < Eps Then
   F_Rayleigh_Inv = ChrW(8734)
   Exit Function
End If

F_Rayleigh_Inv = Sigma * Sqr(-2 * Log(1 - Probabilidad))

End Function


' FUNCIÓN F_Rayleigh_Media

Public Function F_Rayleigh_Media(Sigma As Double) As Variant
' Calcula la media de la distribución de Rayleigh
Dim Pi As Double

Pi = 3.14159265358979
If Sigma <= 0 Then
   F_Rayleigh_Media = "Sigma debe ser > 0"
   Exit Function
End If

F_Rayleigh_Media = Sigma * Sqr(Pi / 2)

End Function


' FUNCIÓN F_Rayleigh_Moda

Public Function F_Rayleigh_Moda(Sigma As Double) As Variant
' Calcula la moda de la distribución de Rayleigh

If Sigma <= 0 Then
   F_Rayleigh_Moda = "Sigma debe ser > 0"
   Exit Function
End If

F_Rayleigh_Moda = Sigma

End Function


' FUNCIÓN F_Rayleigh_DesvTip

Public Function F_Rayleigh_DesvTip(Sigma As Double) As Variant
' Calcula la desviación típica de la distribución de Rayleigh
Dim Pi As Double

Pi = 3.14159265358979
If Sigma <= 0 Then
   F_Rayleigh_DesvTip = "Sigma debe ser > 0"
   Exit Function
End If

F_Rayleigh_DesvTip = Sigma * Sqr((4 - Pi) / 2)

End Function


' FUNCIÓN F_Rayleigh_Asimetria

Public Function F_Rayleigh_Asimetria(Sigma As Double) As Variant
' Calcula coeficiente de asimetría (Fisher) de la distribución de Rayleigh
Dim Pi As Double

Pi = 3.14159265358979
If Sigma <= 0 Then
   F_Rayleigh_Asimetria = "Sigma debe ser > 0"
   Exit Function
End If

F_Rayleigh_Asimetria = 2 * Sqr(Pi) * (Pi - 3) / Sqr((4 - Pi) ^ 3)

End Function


' FUNCIÓN F_Rayleigh_Curtosis

Public Function F_Rayleigh_Curtosis(Sigma As Double) As Variant
' Calcula la curtosis de la distribución de Rayleigh
Dim Pi As Double

Pi = 3.14159265358979
If Sigma <= 0 Then
   F_Rayleigh_Curtosis = "Sigma debe ser > 0"
   Exit Function
End If

F_Rayleigh_Curtosis = 3 - (6 * Pi * Pi - 24 * Pi + 16) / (4 - Pi) ^ 2

End Function


