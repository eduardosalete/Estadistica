
' FUNCIÓN DE DENSIDAD

Public Function D_Triangular(x As Double, a As Double, b As Double, c As Double) As Variant
' Esta función calcula la función de densidad de la distribución Triangular
' a valor inferior, b valor superior, c moda
Dim Denominador1 As Double, Denominador2 As Double

If b <= a Then
   D_Triangular = "Los parámetros deben ser a < b"
   Exit Function
End If
If c < a Or c > b Then
   D_Triangular = "Los parámetros deben ser a <= c <= b"
   Exit Function
End If

If x <= a Or x >= b Then
   ' La función es nula fuera del intervalo (a, b)
   D_Triangular = 0
   Exit Function
End If

If x <= c Then
   ' Estamos en la rama ascendente
   Denominador1 = (b - a) * (c - a)
   D_Triangular = 2 * (x - a) / Denominador1
Else
   ' Estamos en la rama descendente
   Denominador2 = (b - a) * (b - c)
   D_Triangular = 2 * (b - x) / Denominador2
End If

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Triangular(x As Double, a As Double, b As Double, c As Double) As Variant
' Esta función calcula la función de distribución de la distribución Triangular
' a valor inferior, b valor superior, c moda
Dim Denominador1 As Double, Denominador2 As Double

If b <= a Then
   FD_Triangular = "Los parámetros deben ser a < b"
   Exit Function
End If
If c < a Or c > b Then
   FD_Triangular = "Los parámetros deben ser a <= c <= b"
   Exit Function
End If

If x <= a Then
   ' La función es nula
   FD_Triangular = 0
   Exit Function
End If

If x >= b Then
   ' La función vale la unidad
   FD_Triangular = 1
   Exit Function
End If

If x <= c Then
   ' Estamos en la rama ascendente
   Denominador1 = (b - a) * (c - a)
   FD_Triangular = (x - a) ^ 2 / Denominador1
Else
   ' Estamos en la rama descendente
   Denominador2 = (b - a) * (b - c)
   FD_Triangular = 1 - (b - x) ^ 2 / Denominador2
End If

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Triangular_Inv(Probabilidad As Double, a As Double, b As Double, c As Double) As Variant
' Esta función calcula la función de distribución de la distribución Triangular
' a valor inferior, b valor superior, c moda
Dim Denominador1 As Double, Denominador2 As Double, Eps As Double
Dim ProbModa As Double

Eps = 0.0000001

If b <= a Then
   F_Triangular_Inv = "Los parámetros deben ser a < b"
   Exit Function
End If
If c < a Or c > b Then
   F_Triangular_Inv = "Los parámetros deben ser a <= c <= b"
   Exit Function
End If

ProbModa = 1 / (b - a) * (c - a) ' Masa total asociada a la moda

If Probabilidad <= Eps Then
   F_Triangular_Inv = a
   Exit Function
End If

If Abs(Probabilidad - 1) <= Eps Then
   F_Triangular_Inv = b
   Exit Function
End If

If Probabilidad >= 1 Then
   F_Triangular_Inv = ChrW(8734)
   Exit Function
End If

If Probabilidad <= ProbModa Then
   ' Estamos en la rama ascendente
   Denominador1 = (b - a) * (c - a)
   F_Triangular_Inv = a + Sqr(Probabilidad * Denominador1)
Else
   ' Estamos en la rama descendente
   Denominador2 = (b - a) * (b - c)
   F_Triangular_Inv = b - Sqr((1 - Probabilidad) * Denominador2)
End If

End Function


