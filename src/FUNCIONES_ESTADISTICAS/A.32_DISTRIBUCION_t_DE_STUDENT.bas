
' FUNCIÓN DE DENSIDAD

Public Function D_t_Student(x As Double, n As Integer) As Variant
' Calcula la función de densidad de la distribución t de Student
' con n grados de libertad
Dim N1 As Double, Pi As Double
Dim FGamma As Double
Dim m As Double, R2 As Double
Dim m2 As Double, m3 As Double, m5 As Double, m7 As Double

Pi = 3.14159265358979
N1 = (n + 1) / 2
R2 = Sqr(2)

If n <= 0 Then
   D_t_Student = " Los g.d.l. (n) debe ser >0"
   Exit Function
End If

If n > 342 Then
   ' Empleamos la fórmula de Puiseux que da el desarrollo de Gamma((n+1)/2))/Gamma(n/2)
   ' en el entorno del infinito
   m = Sqr(1 / n): m2 = m * m: m3 = m * m2: m5 = m3 * m2: m7 = m5 * m2
   FGamma = Sqr(n / 2) - m / 4 / R2 + m3 / 32 / R2
   FGamma = FGamma + 5 * m5 / 128 / R2 - 21 * m7 / 2048 / R2
Else
   FGamma = F_Gamma(N1) / F_Gamma(n / 2)
End If

D_t_Student = FGamma / Sqr(n * Pi) / (1 + x * x / n) ^ N1

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_t_Student(x As Double, n As Integer) As Variant
' Calcula la función de distribución de la distribución t de Student
' con n grados de libertad
Dim a As Double
Dim N1 As Double, Pi As Double
Dim a1 As Double, a2 As Double
Dim FGamma As Double
Dim xx As Double, m As Double, R2 As Double
Dim m2 As Double, m3 As Double, m5 As Double, m7 As Double

Pi = 3.14159265358979
N1 = (n + 1) / 2
xx = -x * x / n
R2 = Sqr(2)

If n <= 0 Then
   FD_t_Student = " Los g.d.l. (n) debe ser >0"
   Exit Function
End If

If n > 342 Then
   ' Empleamos la fórmula de Puiseux que da el desarrollo de Gamma((n+1)/2))/Gamma(n/2)
   ' en el entorno del infinito
   m = Sqr(1 / n): m2 = m * m: m3 = m * m2: m5 = m3 * m2: m7 = m5 * m2
   FGamma = Sqr(n / 2) - m / 4 / R2 + m3 / 32 / R2
   FGamma = FGamma + 5 * m5 / 128 / R2 - 21 * m7 / 2048 / R2
Else
   FGamma = F_Gamma(N1) / F_Gamma(n / 2)
End If
   
a1 = x * FGamma / Sqr(n * Pi)

If n = 2 Then
   ' caso particular n=2
   ' Aplicamos que 1F2(a,b;b|z)=(1-z)^(-1)
   a2 = (1 - xx) ^ (-0.5)
Else
   a2 = F_HG_Gauss(0.5, N1, 1.5, xx)
End If

FD_t_Student = 0.5 + a1 * a2

End Function


' FUNCIÓN DE DISTRIBUCIÓN DE DOS COLAS

Public Function FD_t_Student2(x As Double, n As Integer) As Variant
' Calcula la función de distribución de dos colas de la distribución t de Student
' con n grados de libertad
' Llama a la función FD_t_Student

If n <= 0 Then
   FD_t_Student2 = " Los g.d.l. (n) debe ser >0"
   Exit Function
End If

FD_t_Student2 = 2 * FD_t_Student(x, n) - 1

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_t_Student_Inv(Probabilidad As Double, n As Integer) As Variant
' Esta función obtiene la inversa de la función de distribución t de Student
' Llama a la función Mi_ecuacion_Est

Dim Sigma As Double
Dim x1 As Double, x2 As Double, Factor As Double
Dim Hecho As String, aa As Variant
Dim Ene As Double, Eps As Double

Eps = 0.0000001

If n <= 0 Then
   F_t_Student_Inv = " Los g.d.l. (n) debe ser >0"
   Exit Function
End If

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_t_Student_Inv = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If Probabilidad < Eps Then
   F_t_Student_Inv = "-" & ChrW(8734)
   Exit Function
End If

If Abs(Probabilidad - 1) < Eps Then
   F_t_Student_Inv = ChrW(8734)
   Exit Function
End If

Ene = n                  ' Pasamos a número real el parámetro
If n > 2 Then
   Sigma = Sqr(n / (n - 2))
Else
   Sigma = 1
End If

Factor = 5
Hecho = "No"

Do While Hecho = "No"
   x1 = -Factor * Sigma
   x2 = Factor * Sigma
   aa = Mi_ecuacion_Est("t_Student", Probabilidad, x1, x2, Ene, , , 0.000000001)
   If aa = "Rango Mal" Then
      ' El rango no contenía la raíz y lo aumentamos
      Factor = Factor + 1
   Else
      ' Se ha obtenido la raíz que se ha metido en la variable aa
      Hecho = "Si"
      F_t_Student_Inv = aa
   End If
Loop

End Function


