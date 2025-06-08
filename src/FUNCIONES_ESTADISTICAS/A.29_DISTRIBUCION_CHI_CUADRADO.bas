
' FUNCIÓN DE DENSIDAD

Public Function D_Chi_Cuadrado(x As Double, n As Integer) As Variant
' Calcula la función de densidad de la distribución Chi-Cuadrado
' con n grados de libertad
' Llama a la función F_Gamma
Dim Eps As Double

Eps = 0.0000001

If n <= 0 Then
   D_Chi_Cuadrado = " Los g.d.l. (n) debe ser >0"
   Exit Function
End If

If x < Eps Then
   If n = 1 Then
      D_Chi_Cuadrado = ChrW(8734)
   ElseIf n = 2 Then
      D_Chi_Cuadrado = 0.5
   Else
      D_Chi_Cuadrado = 0
   End If
   Exit Function
End If

D_Chi_Cuadrado = 2 ^ (-n / 2) / F_Gamma(n / 2) * x ^ (n / 2 - 1) * Exp(-x / 2)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Chi_Cuadrado(x As Double, n As Integer) As Variant
' Calcula la función de distribución de la distribución Chi-Cuadrado
' con n grados de libertad
' Llama a la función F_Gamma
' Llama a la función F_Gamma_Inf
Dim a As Double, Eps As Double

Eps = 0.0000001

If n <= 0 Then
   FD_Chi_Cuadrado = " Los g.d.l. (n) debe ser >0"
   Exit Function
End If

If x < Eps Then
   FD_Chi_Cuadrado = 0
   Exit Function
End If

a = n / 2
FD_Chi_Cuadrado = 1 / F_Gamma(a) * F_Gamma_Inf(a, x / 2)

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Chi_Cuadrado_Inv(Probabilidad As Double, n As Integer) As Variant
' Esta función obtiene la inversa de la función de distribución Chi-Cuadrado
' Llama a la función Mi_ecuacion_Est
Dim Sigma As Double
Dim x1 As Double, x2 As Double, Factor As Double
Dim Hecho As String, aa As Variant
Dim Ene As Double, Eps As Double

Eps = 0.0000001

If n <= 0 Then
   F_Chi_Cuadrado_Inv = " Los g.d.l. (n) debe ser >0"
   Exit Function
End If

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_Chi_Cuadrado_Inv = "La probabilidad debe ser <=1 y >=0"
   Exit Function
End If

If Probabilidad < Eps Then
   F_Chi_Cuadrado_Inv = 0
   Exit Function
End If

If Abs(Probabilidad - 1) < Eps Then
   F_Chi_Cuadrado_Inv = ChrW(8734)
   Exit Function
End If

Ene = n                  ' Pasamos a número real el parámetro
Sigma = Sqr(2 * n)
Factor = 5
Hecho = "No"

Do While Hecho = "No"
   x1 = -Factor * Sigma
   x2 = Factor * Sigma
   aa = Mi_ecuacion_Est("Chi_Cuadrado", Probabilidad, x1, x2, Ene, , , 0.000000001)
   If aa = "Rango Mal" Then
      ' El rango no contenía la raíz y lo aumentamos
      Factor = Factor + 1
   Else
      ' Se ha obtenido la raíz que se ha metido en la variable aa
      Hecho = "Si"
      F_Chi_Cuadrado_Inv = aa
   End If
Loop

End Function


