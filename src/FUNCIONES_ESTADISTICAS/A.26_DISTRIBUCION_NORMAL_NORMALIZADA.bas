
' FUNCIÓN DE DENSIDAD

Public Function D_Normal_01(x As Double) As Double
' Esta función calcula la función de densidad de la distribución N(0,1)
Dim R2Pi As Double

R2Pi = 2.506628274631

D_Normal_01 = 1 / R2Pi * Exp(-x * x / 2)
End Function


' FUNCIÓN DE DISTRIBUCIÓN (EXACTA)

Public Function FD_Normal_01_G(xx As Double) As Double
' Esta función calcula la función de distribución "exacta" de la distribución N(0,1)
' Llama a la función F_Gamma_Inf

Dim i As Integer, Positivo As Integer
Dim G1_2 As Double, x As Double

G1_2 = 1.77245385090552

Positivo = 1: x = xx
If x < 0 Then
   x = -xx
   Positivo = 0
End If

FD_Normal_01_G = (F_Gamma_Inf(0.5, x * x / 2) / G1_2 + 1) / 2

If Positivo = 0 Then
   FD_Normal_01_G = 1 - FD_Normal_01_G
End If
End Function


' FUNCIÓN DE DISTRIBUCIÓN DE DOS COLAS (EXACTA)

Public Function FD_Normal_01_G2(xx As Double) As Double
' Esta función calcula la función de distribución de dos colas "exacta" de la distribución N(0,1)
' Llama a la función FD_Normal_01_G

FD_Normal_01_G2 = 2 * FD_Normal_01_G(xx) - 1
End Function


' FUNCIÓN DE DISTRIBUCIÓN (HASTINGS)

Public Function FD_Normal_01_H(xx As Double) As Double
' Esta función calcula la función de distribución de la distribución N(0,1)
' Utiliza la fórmula aproximada de Hastings
' (Abramowitz y Stegun, pág. 932.
' Error < 7.5 E-8)
' Llama a la función D_Normal_01

Dim i As Integer, Positivo As Integer
Dim b(0 To 6) As Double, Phi As Double, t As Double, tt As Double, x As Double

b(0) = 0.2316419
b(1) = 0.31938153
b(2) = -0.356563782
b(3) = 1.781477937
b(4) = -1.821255978
b(5) = 1.330274429

Positivo = 1: x = xx
If xx < 0 Then
   x = -xx
   Positivo = 0
End If

Phi = D_Normal_01(x)
t = 1 / (1 + b(0) * x)

FD_Normal_01_H = 0
tt = Phi
For i = 1 To 5
  tt = tt * t
  FD_Normal_01_H = FD_Normal_01_H + b(i) * tt
Next
FD_Normal_01_H = 1 - FD_Normal_01_H

If Positivo = 0 Then
   FD_Normal_01_H = 1 - FD_Normal_01_H
End If
End Function


' FUNCIÓN DE DISTRIBUCIÓN DE DOS COLAS (HASTINGS)

Public Function FD_Normal_01_H2(xx As Double) As Double
' Esta función calcula la función de distribución de dos colas de la distribución N(0,1)
' Utiliza la fórmula aproximada de Hastings
' (Abramowitz y Stegun, pág. 932.
' Error < 7.5 E-8)
' Llama a la función FD_Normal_01_H

FD_Normal_01_H2 = 2 * FD_Normal_01_H(xx) - 1
End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Normal_01_Inv(Probabilidad As Double, Optional Procedimiento As Double = 2) As Variant
' Esta función obtiene la inversa de la función de distribución N(0,1)
' Si Procedimiento=1 emplea la relación con la función Gamma Incompleta Inferior
' Si Procedimiento=2 emplea la aproximación de Hastings
' Llama a la función Mi_ecuacion_Est

Dim Mu As Double, Sigma As Double
Dim x1 As Double, x2 As Double, Factor As Double
Dim Hecho As String, aa As Variant
Dim Eps As Double

Eps = 0.0000001

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_Normal_01_Inv = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If Probabilidad < Eps Then
   F_Normal_01_Inv = 0
   Exit Function
End If

If Abs(Probabilidad - 1) < Eps Then
   F_Normal_01_Inv = ChrW(8734)
   Exit Function
End If

Mu = 0: Sigma = 1
Factor = 5
Hecho = "No"
Do While Hecho = "No"
   x1 = Mu - Factor * Sigma
   x2 = Mu + Factor * Sigma
   aa = Mi_ecuacion_Est("Normal_01", Probabilidad, x1, x2, Procedimiento, , , 0.000000001)
   If aa = "Rango Mal" Then
      ' El rango no contenía la raíz y lo aumentamos
      Factor = Factor + 1
   Else
      ' Se ha obtenido la raíz que se ha metido en la variable aa
      Hecho = "Si"
      F_Normal_01_Inv = aa
   End If
Loop

End Function


