
' FUNCIÓN DE DENSIDAD

Public Function D_Irwin_Hall(xx As Double, n As Long) As Variant
' Esta función calcula la función de densidad de la distribución de Irwin-Hall
' Llama a la función auxiliar Mi_Factorial
' Llama a la función auxiliar D_I_H
' Llama a la función D_Normal_MS y ésta a D_Normal_01
' Cuando n>=25 utiliza la aproximación en el intervalo [0,n]
'
Dim i As Long, Eps As Double, x As Double
Dim Mu As Double, Sigma As Double

Eps = 0.0000001

If n <= 0 Then
   D_Irwin_Hall = "n debe ser >0"
   Exit Function
End If

x = xx

If x < Eps Or x > n Then
   D_Irwin_Hall = 0
   Exit Function
End If
If x > n / 2 Then
   x = n - x
End If

If x <= 1 + Eps Then
   D_Irwin_Hall = x ^ (n - 1) / Mi_Factorial(n - 1)
   Exit Function
End If

If n < 25 Then
   ' Aplicamos la expresión de Irwin-Hall directamente
   D_Irwin_Hall = 0
   For i = 0 To n
       If Abs(x - i) > Eps Then
         'D_Irwin_Hall = D_Irwin_Hall + (-1) ^ i * (x - i) ^ (n - 1) / Mi_Factorial(i) / Mi_Factorial(n - i) * Sgn(x - i)
          D_Irwin_Hall = D_Irwin_Hall + (-1) ^ i * D_I_H(x, i, n)
       End If
   Next
   D_Irwin_Hall = Abs(D_Irwin_Hall * n / 2)
Else
   ' Aplicamos la aproximación a la distribución Normal
   Mu = n / 2
   Sigma = Sqr(n / 12)
   D_Irwin_Hall = D_Normal_MS(x, Mu, Sigma)
End If

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Irwin_Hall(x As Double, n As Long) As Variant
' Esta función calcula la función de distribución de la distribución de Irwin-Hall
' Llama a la función auxiliar Mi_Factorial
' Llama a la función auxiliar D_I_H
' Llama a la función FD_Normal_MS y ésta a FD_Normal_01_G y FD_Normal_01_H
' Cuando n>=25 utiliza la aproximación normal en el intervalo [0,n]

Dim i As Long, Eps As Double
Dim Mu As Double, Sigma As Double

Eps = 0.0000001

If n <= 0 Then
   FD_Irwin_Hall = "n debe ser >0"
   Exit Function
End If

If x < Eps Then
   FD_Irwin_Hall = 0
   Exit Function
End If

If x <= 1 + Eps Then
   FD_Irwin_Hall = x ^ n / Mi_Factorial(n)
   Exit Function
End If

If x >= n Then
   FD_Irwin_Hall = 1
   Exit Function
End If

If x > n - 1 + Eps Then
   FD_Irwin_Hall = (n - x) ^ n / Mi_Factorial(n)
   FD_Irwin_Hall = 1 - FD_Irwin_Hall
   Exit Function
End If

If n <= 25 Then
   ' Aplicamos la expresión de Irwin-Hall directamente
   FD_Irwin_Hall = 1 / 2
   For i = 0 To n
       If Abs(x - i) > Eps Then
           FD_Irwin_Hall = FD_Irwin_Hall + (-1) ^ i * D_I_H(x, i, n) * (x - i) / 2
       End If
   Next
Else
   ' Aplicamos la aproximación a la distribución Normal
   Mu = n / 2
   Sigma = Sqr(n / 12)
   FD_Irwin_Hall = FD_Normal_MS(x, Mu, Sigma, 2)
End If

End Function


' FUNCIÓN AUXILIAR D_I_H

Public Function D_I_H(x As Double, i As Long, n As Long) As Double
' Función auxiliar que calcula el término
' (x - i) ^ (n - 1) / Mi_Factorial(i) / Mi_Factorial(n - i)* Sgn(x - i)
' Evitando overflow para potencias grandes
' Lo calculamos como
' [(x - i) ^ i / Mi_Factorial(i)] * [(x - i) ^ (n-i) / Mi_Factorial(n-i)] / (x-i) *sgn(x-i)
Dim j As Long, xij As Double

D_I_H = 1
If i = 0 Then
    For j = 1 To n
       xij = x / j
       D_I_H = D_I_H * xij
    Next
    D_I_H = D_I_H / x
    Exit Function
Else
    For j = 1 To i
       xij = (x - i) / j
       D_I_H = D_I_H * xij
    Next
    For j = 1 To n - i
       xij = (x - i) / j
       D_I_H = D_I_H * xij
    Next
    D_I_H = D_I_H * Sgn(x - i) / (x - i)
End If

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Irwin_Hall_Inv(Probabilidad As Double, n As Long) As Variant
' Esta función calcula la inversa de la función de distribución de la distribución de Irwin-Hall
' Llama a la función Mi_ecuacion_Est
Dim Eps As Double, x1 As Double, x2 As Double, nD As Double
Dim aa As Variant

Eps = 0.0000001

If n <= 0 Then
   F_Irwin_Hall_Inv = "n debe ser >0"
   Exit Function
End If

If Probabilidad < 0 Then
   F_Irwin_Hall_Inv = "Probabilidad negativa"
   Exit Function
End If

If Probabilidad > 1 Then
   F_Irwin_Hall_Inv = "Probabilidad > 1"
   Exit Function
End If

If Probabilidad < Eps Then
   F_Irwin_Hall_Inv = 0
   Exit Function
End If

If Probabilidad >= 1 - Eps Then
   F_Irwin_Hall_Inv = n
   Exit Function
End If

nD = n
x1 = 0
x2 = nD

aa = Mi_ecuacion_Est("I-H", Probabilidad, x1, x2, nD, , , 0.000000001)

If aa = "Rango Mal" Then
   ' Esto no debería ocurrir nunca
   F_Irwin_Hall_Inv = "No ha sido posible obtener el resultado"
Else
   F_Irwin_Hall_Inv = aa
End If

End Function


' FUNCIÓN F_Irwin_Hall_Media

Public Function F_Irwin_Hall_Media(n As Long) As Variant
' Calcula la media de la distribución de Irwin-Hall

If n <= 0 Then
   F_Irwin_Hall_Media = "n debe ser >0"
   Exit Function
End If

F_Irwin_Hall_Media = n / 2

End Function


' FUNCIÓN F_Irwin_Hall_Moda

Public Function F_Irwin_Hall_Moda(n As Long) As Variant
' Calcula la moda de la distribución de Irwin-Hall
If n <= 0 Then
   F_Irwin_Hall_Moda = "n debe ser >0"
   Exit Function
End If

If n = 1 Then
   F_Irwin_Hall_Moda = "Cualquier valor entre 0 y 1"
Else
   F_Irwin_Hall_Moda = n / 2
End If

End Function


' FUNCIÓN F_Irwin_Hall_Mediana

Public Function F_Irwin_Hall_Mediana(n As Long) As Variant
' Calcula la mediana de la distribución de Irwin-Hall
If n <= 0 Then
   F_Irwin_Hall_Mediana = "n debe ser >0"
   Exit Function
End If

F_Irwin_Hall_Mediana = n / 2

End Function


' FUNCIÓN F_Irwin_Hall_DesvTip

Public Function F_Irwin_Hall_DesvTip(n As Long) As Variant
' Calcula la desviación típica de la distribución de Irwin-Hall

If n <= 0 Then
   F_Irwin_Hall_DesvTip = "n debe ser >0"
   Exit Function
End If

F_Irwin_Hall_DesvTip = Sqr(n / 12)

End Function


' FUNCIÓN F_Irwin_Hall_Asimetria

Public Function F_Irwin_Hall_Asimetria(n As Long) As Variant
' Calcula el coeficiente de asimetría (Fisher) de la distribución de Irwin-Hall

If n <= 0 Then
   F_Irwin_Hall_Asimetria = "n debe ser >0"
   Exit Function
End If

F_Irwin_Hall_Asimetria = 0

End Function


' FUNCIÓN F_Irwin_Hall_Curtosis

Public Function F_Irwin_Hall_Curtosis(n As Long) As Variant
' Calcula la curtosis de la distribución de Irwin-Hall

If n <= 0 Then
   F_Irwin_Hall_Curtosis = "n debe ser >0"
   Exit Function
End If

F_Irwin_Hall_Curtosis = 3 - 6 / 5 / n

End Function


