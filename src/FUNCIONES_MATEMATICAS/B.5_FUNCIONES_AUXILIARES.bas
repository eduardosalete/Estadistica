
' RUTINA Mi_es_Entero

Public Sub Mi_es_Entero(x As Double, N As Integer, Tipo As Integer)
' Esta rutina estudia si un número x es entero y devuelve
' n=[x]
' Tipo =  1 si es entero positivo.
'      =  0 si vale 0.
'      = -1 si es entero negativo.
'      =  2 si no es entero.
Dim Eps As Double

Eps = 0.0000001

Tipo = 2
If Abs(Fix(Abs(x)) - Abs(x)) < Eps Then
   N = x
   If N < 0 Then
      Tipo = -1
   ElseIf N = 0 Then
      Tipo = 0
   ElseIf N > 0 Then
      Tipo = 1
   End If
Else
   N = Fix(x)
   Tipo = 2
End If

End Sub


' FUNCIÓN NÚMERO COMBINATORIO

Public Function N_Combinatorio(N As Long, i As Long) As Double
' Número combinatorio N sobre i

Dim j As Long

If i = 0 Or i = N Then
   N_Combinatorio = 1
   Exit Function
End If

N_Combinatorio = 1
For j = 1 To i
    N_Combinatorio = N_Combinatorio / j * (N - j + 1)
Next

End Function


' FUNCIÓN FACTORIAL

Public Function Mi_Factorial(N As Long) As Double
' Calcula n! como número en doble precisión

Dim i As Long

Mi_Factorial = 1
If N = 0 Then
   Exit Function
End If

For i = 1 To N
   Mi_Factorial = i * Mi_Factorial
Next

End Function


' FUNCIÓN Mi_ecuacion_Est

Public Function Mi_ecuacion_Est(Distr As String, _
                Probabilidad As Double, _
                x1 As Double, x2 As Double, _
                Optional a1 As Double, Optional a2 As Double, Optional a3 As Double, _
                Optional Eps As Double = 0.0000001) As Variant
' Esta función resuelve por bisección la ecuación F(Distr, x)=0
' x1 y x2 son los valores extremos del intervalo en donde está la raíz
' Busca la raíz en el intervalo [Mu-Veces_s x Sigma, Mu+Veces_s x Sigma]
' F depende de la distribución de que se trate
' Llama a la función Mi_Funcion_Est

Dim f1 As Double, f2 As Double, ff As Double
Dim xx As Double, Error As Double, Iteracion As Long
Dim xx1(1 To 1000) As Double, xx2(1 To 1000) As Double, fff(1 To 1000) As Double

f1 = Mi_Funcion_Est(Distr, x1, Probabilidad, a1, a2, a3)
f2 = Mi_Funcion_Est(Distr, x2, Probabilidad, a1, a2, a3)
Iteracion = 0

If f1 * f2 > 0 Then
   Mi_ecuacion_Est = "Rango Mal"
   Exit Function
End If

ff = 1
Do While Abs(ff) > Eps ' And Iteracion < 100
   xx = (x1 + x2) / 2
   ff = Mi_Funcion_Est(Distr, xx, Probabilidad, a1, a2, a3)
   If ff * f1 > 0 Then
      x1 = xx
      f1 = ff
   Else
      x2 = xx
   End If
   Iteracion = Iteracion + 1
   xx1(Iteracion) = x1: xx2(Iteracion) = x2: fff(Iteracion) = ff
Loop

Mi_ecuacion_Est = xx
End Function


' FUNCIÓN Mi_Funcion_Est

Public Function Mi_Funcion_Est(Distr As String, x As Double, _
                Probabilidad As Double, _
                Optional a1 As Double, Optional a2 As Double, _
                Optional a3 As Double, Optional a4 As Double, _
                Optional a5 As Double, Optional a6 As Double) As Double

' Esta rutina calcula el valor de la función de distribución,
' Dependiendo de la distribución de que se trate
' Los parámetros pueden ser los que definen la distribución o
' valores auxiliares para aumentar la velocidad del cálculo

' Llama a la función FD_Gamma
' Llama a la función FD_Beta
' Llama a la función FD_Irwin_Hall
' Llama a la función FD_Normal_01_G
' Llama a la función FD_Normal_01_H
' Llama a la función FD_Chi_Cuadrado
' Llama a la función FD_t_Student
' Llama a la función FD_F_Snedecor
' Llama a la función FD_M_B de Maxwell-Boltzmann

Dim N As Integer, n1 As Integer, n2 As Integer
Dim nih As Long

Select Case Distr
   Case "Gamma"
      ' Alfa=a1, Beta=a2, Gamma(Alfa)=a3
      Mi_Funcion_Est = FD_Gamma(x, a1, a2, a3) - Probabilidad
   Case "Beta"
      ' Alfa=a1, Beta=a2,
      Mi_Funcion_Est = FD_Beta(x, a1, a2) - Probabilidad
   Case "I-H"
      nih = a1
      Mi_Funcion_Est = FD_Irwin_Hall(x, nih) - Probabilidad
   Case "Normal_01"
      If Abs(a1 - 1) <= 0.01 Then
         ' Cálculo con la función Gamma Incompleta Inferior
         Mi_Funcion_Est = FD_Normal_01_G(x) - Probabilidad
      Else
         ' Cálculo con la aproximación de Hastings
         Mi_Funcion_Est = FD_Normal_01_H(x) - Probabilidad
      End If
    Case "Chi_Cuadrado"
         N = a1
         Mi_Funcion_Est = FD_Chi_Cuadrado(x, N) - Probabilidad
    Case "t_Student"
         N = a1
         Mi_Funcion_Est = FD_t_Student(x, N) - Probabilidad
    Case "F_Snedecor"
         n1 = a1: n2 = a2
         Mi_Funcion_Est = FD_F_Snedecor(x, n1, n2) - Probabilidad
    Case "M_B"
         Mi_Funcion_Est = FD_M_B(x, a1) - Probabilidad
End Select

End Function


' FUNCIÓN S_Zeta_Entera

Public Function S_Zeta_Entera(N As Integer) As Variant
Dim Zeta(1 To 48) As Double
Dim i As Integer

If N = 1 Then
   S_Zeta_Entera = "+" & ChrW(8734)
   Exit Function
End If

If N > 47 Then
   S_Zeta_Entera = 1
   Exit Function
End If

' Primeros 2 a 47 valores
' (el primero de todos es la serie armónica que vale infinito)
i = 1
i = i + 1: Zeta(i) = 1.64493406684823  ' 2
i = i + 1: Zeta(i) = 1.20205690315959  ' 3
i = i + 1: Zeta(i) = 1.08232323371114  ' 4
i = i + 1: Zeta(i) = 1.03692775514337  ' 5
i = i + 1: Zeta(i) = 1.01734306198445  ' 6
i = i + 1: Zeta(i) = 1.00834927738192  ' 7
i = i + 1: Zeta(i) = 1.00407735619794  ' 8
i = i + 1: Zeta(i) = 1.00200839282608  ' 9
i = i + 1: Zeta(i) = 1.00099457512782  ' 10
i = i + 1: Zeta(i) = 1.00049418860412  ' 11
i = i + 1: Zeta(i) = 1.00024608655331  ' 12
i = i + 1: Zeta(i) = 1.00012271334758  ' 13
i = i + 1: Zeta(i) = 1.00006124813506  ' 14
i = i + 1: Zeta(i) = 1.00003058823631  ' 15
i = i + 1: Zeta(i) = 1.00001528225941  ' 16
i = i + 1: Zeta(i) = 1.00000763719764  ' 17
i = i + 1: Zeta(i) = 1.00000381729327  ' 18
i = i + 1: Zeta(i) = 1.00000190821272  ' 19
i = i + 1: Zeta(i) = 1.00000095396203  ' 20
i = i + 1: Zeta(i) = 1.00000047693299  ' 21
i = i + 1: Zeta(i) = 1.0000002384505   ' 22
i = i + 1: Zeta(i) = 1.00000011921993  ' 23
i = i + 1: Zeta(i) = 1.00000005960819  ' 24
i = i + 1: Zeta(i) = 1.0000000298035   ' 25
i = i + 1: Zeta(i) = 1.00000001490155  ' 26
i = i + 1: Zeta(i) = 1.00000000745071  ' 27
i = i + 1: Zeta(i) = 1.00000000372533  ' 28
i = i + 1: Zeta(i) = 1.00000000186266  ' 29
i = i + 1: Zeta(i) = 1.00000000093133  ' 30
i = i + 1: Zeta(i) = 1.00000000046566  ' 31
i = i + 1: Zeta(i) = 1.00000000023283  ' 32
i = i + 1: Zeta(i) = 1.00000000011642  ' 33
i = i + 1: Zeta(i) = 1.00000000005821  ' 34
i = i + 1: Zeta(i) = 1.0000000000291   ' 35
i = i + 1: Zeta(i) = 1.00000000001455  ' 36
i = i + 1: Zeta(i) = 1.00000000000728  ' 37
i = i + 1: Zeta(i) = 1.00000000000364  ' 38
i = i + 1: Zeta(i) = 1.00000000000182  ' 39
i = i + 1: Zeta(i) = 1.00000000000091  ' 40
i = i + 1: Zeta(i) = 1.00000000000045  ' 41
i = i + 1: Zeta(i) = 1.00000000000023  ' 42
i = i + 1: Zeta(i) = 1.00000000000006  ' 43
i = i + 1: Zeta(i) = 1.00000000000003  ' 44
i = i + 1: Zeta(i) = 1.00000000000001  ' 45
i = i + 1: Zeta(i) = 1.00000000000001  ' 46
i = i + 1: Zeta(i) = 1.00000000000001  ' 47

If N <= 47 Then
    S_Zeta_Entera = Zeta(N)
Else
    S_Zeta_Entera = 1
End If
End Function


' FUNCIÓN F_erf

Public Function F_erf(x As Double) As Double
' Esta función calcula la función de error de Gauss
' Llama a la función de distribución de la distribución N(0,1)
Dim R2 As Double

R2 = 1.4142135623731
F_erf = 2 * FD_Normal_01_H(R2 * x) - 1

End Function


' FUNCIÓN F_erfc

Public Function F_erfc(x As Double) As Double
' Esta función calcula la función de error complementaria
' Llama a la función de error de Gauss

F_erfc = 1 - F_erf(x)

End Function


