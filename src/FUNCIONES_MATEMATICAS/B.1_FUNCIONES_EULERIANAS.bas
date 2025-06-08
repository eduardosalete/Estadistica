
' FUNCIÓN GAMMA

Public Function F_Gamma(xx As Double) As Variant
' Calcula la función Gamma(xx) para cualquier argumento real
' Se emplea la aproximación de 26 sumandos de H.T. Davis
' "Table of mathematical functions". Principia Press, Bloomington, Ind. 1935.
' Con correcciones debidas a H.E. Salzer

Dim c(1 To 26) As Double, a As Double, Suma As Double
Dim x As Double, Pi As Double
Dim i As Integer, N As Integer
Dim Caso As String, Eps As Double
Dim xAsintSup As Integer

Pi = 3.14159265358979
Eps = 0.0000001
xAsintSup = 150

' Primero vemos si se sale del rango de números que manejamos
If xx > xAsintSup Then
  F_Gamma = "La función toma valores demasiado grandes"
  Exit Function
End If

' Calculamos la parte entera de xx (N)
N = Fix(xx)
If xx <= 0 And Abs(N - xx) < Eps Then
   ' xx es cero o un entero negativo -> Gamma = +- infinito
   F_Gamma = ChrW(177) & " " & ChrW(8734)
   Exit Function
End If

If xx > 0 And Abs(N - xx) < Eps Then
   ' xx es un entero positivo
   F_Gamma = 1
   For i = 1 To N - 1
     F_Gamma = F_Gamma * i
   Next
   Exit Function
End If

' Tratamiento general (para xx no entero)
If xx < 0 Then
   x = 1 - xx
   N = Fix(x)
   x = x - N
   Caso = "Negativo"
Else
   x = xx - N
   Caso = "Positivo"
End If

' Establecer los coeficientes del desarrollo en serie
c(1) = 1
c(2) = 0.577215664901533
c(3) = -0.655878071520254
c(4) = -4.20026350340952E-02
c(5) = 0.166538611382292
c(6) = -4.21977345555443E-02
c(7) = -0.009621971527877
c(8) = 0.007218943246663
c(9) = -1.1651675918591E-03
c(10) = -2.152416741149E-04
c(11) = 1.280502823882E-04
c(12) = -2.01348547807E-05
c(13) = -1.2504934821E-06
c(14) = 0.000001133027232
c(15) = -2.056338417E-07
c(16) = 0.000000006116095
c(17) = 5.0020075E-09
c(18) = -1.1812746E-09
c(19) = 1.043427E-10
c(20) = 7.7823E-12
c(21) = -3.6968E-12
c(22) = 0.00000000000051
c(23) = 2.06E-14
c(24) = -5.4E-15
c(25) = 1.4E-15
c(26) = 1E-16

' Calculamos Gamma(x) mediante la suma de 26 sumandos
a = 1: Suma = 0
For i = 1 To 26
    a = a * x
    Suma = Suma + c(i) * a
Next
F_Gamma = 1 / Suma

If N >= 1 Then
   For i = 1 To N
       F_Gamma = x * F_Gamma
       x = x + 1
   Next
End If

If N < -1 Then
   For i = N To 0
       F_Gamma = x * F_Gamma
       x = x - 1
   Next
End If

If Caso = "Negativo" Then
   F_Gamma = Pi / Sin((1 - xx) * Pi) / F_Gamma
End If
End Function


' FUNCIÓN GAMMA ASTERISCO

Public Function F_GammaA(a As Double, x As Double, _
Optional NSumandos As Integer = 1000, _
Optional Eps As Double = 0.0000001) As Double

' Calcula la función Gamma Asterisco para cualesquiera argumentos reales
' NSumandos es el número de términos de la serie que vamos a sumar
' Eps error permitido al comparar con cero

' Llama a la rutina Mi_es_Entero
' Llama a la función F_Gamma

Dim Suma As Double, Sumando As Double, Ga As Double
Dim i As Integer, N As Integer, nn As Integer, Tipo As Integer
Dim SS(0 To 1000) As Double

' Expresión asintótica
'If x > xAsintSup Then
'   F_GammaA = x ^ (-a)
'   Exit Function
'End If

' Analizamos si a es un número entero negativo
Call Mi_es_Entero(a, nn, Tipo)
If Tipo = -1 Then
   ' a es un número entero negativo, simplificamos el cálculo
   F_GammaA = x ^ (-nn)
   Exit Function
End If

' Calculamos la parte entera de a (N)
N = Fix(a)
If a <= 0 And Abs(N - a) < Eps Then
   ' a es cero o un entero negativo
   F_GammaA = 1
   For i = 1 To Abs(N)
      F_GammaA = F_GammaA * x
   Next
   Exit Function
End If

Sumando = 1 / F_Gamma(a + 1): Suma = Sumando
SS(0) = Sumando
For i = 1 To NSumandos
   Sumando = Sumando * x / (a + i)
   Suma = Suma + Sumando
   SS(i) = Sumando
Next
Ga = Exp(-x)
F_GammaA = Ga * Suma

End Function


' FUNCIÓN GAMMA INCOMPLETA INFERIOR

Public Function F_Gamma_Inf(a As Double, x As Double, _
Optional NSumandos As Integer = 1000, _
Optional xAsintInf As Double = 0.001, _
Optional xAsintSup As Double = 150, _
Optional Eps As Double = 0.0000001) As Variant

' Función Gamma Incompleta Inferior (gamma minúscula)
' NSumandos es el número de sumandos de la serie
' xAsintInf valor para emplear la aproximación asntótica cero
' xAsintSup valor para emplear la aproximación asntótica infinito
' Eps valor para comparar con cero

' Llama a la función F_Gamma
' Llama a la función F_GammaA

Dim N As Integer, i As Integer
Dim Ga As Double, Xa As Double
Dim FG As Double, Sumando As Double, Suma As Double

' Primero miramos si a y x son negativos
If a <= 0 And x < 0 Then
   ' a y x son negativos => la función no está definida
   F_Gamma_Inf = "(" & ChrW(945) & ", x) Fuera de rango"
   Exit Function
End If

' Ahora si x es cero
If Abs(x) < Eps Then
   ' x es nulo => la función es nula
   F_Gamma_Inf = 0
   Exit Function
End If

' Límites asintóticos
If Abs(x) < xAsintInf Then
   ' Aplicamos relación asíntótica cuando x->0
   F_Gamma_Inf = x ^ a / a
   Exit Function
ElseIf x >= xAsintSup Then
   F_Gamma_Inf = F_Gamma(a)
   Exit Function
End If

' Ahora calculamos la parte entera de a (N)
N = Fix(a)
If a <= 0 And Abs(N - a) < Eps Then
   ' a es cero o un entero negativo -> Gamma = +- infinito
   F_Gamma_Inf = ChrW(177) & " " & ChrW(8734)
   Exit Function
End If

' Miramos si a es un entero positivo
If a > 0 And Abs(N - a) < Eps Then
   ' a es un entero positivo
   a = N
   Xa = 1: Ga = 1 / N
   For i = 1 To N
     Xa = Xa * x         ' x^N
     Ga = Ga * i         ' Ga es (N-1)!
   Next
   
   Suma = 1: Sumando = 1
   For i = 1 To N - 1
     Sumando = Sumando * x / i
     Suma = Suma + Sumando
   Next
   F_Gamma_Inf = Ga * (1 - Suma * Exp(-x))
   Exit Function
End If

Ga = F_Gamma(a): Xa = x ^ a
F_Gamma_Inf = Ga * Xa * F_GammaA(a, x, NSumandos, Eps)

End Function


' FUNCIÓN GAMMA INCOMPLETA SUPERIOR

Public Function F_Gamma_Sup(a As Double, x As Double, _
Optional NSumandos As Integer = 1000, _
Optional xAsintSup As Double = 150, _
Optional Eps As Double = 0.0000001) As Variant

' Función Gamma Incompleta Superior
' NSumandos es el número de sumandos de la serie
' xAsintSup valor para emplear la aproximación asntótica infinito
' Eps valor para comparar con cero

' Llama a la función F_En y se realiza la cadena de llamadas:
' F_En -> F_E1 -> F_Ei -> F_li -> Aux_li
' Llama a la función F_Gamma
' Llama a la función F_Gamma_Inf

Dim N As Integer, nn As Integer, MM As Integer, i As Integer
Dim Ga As Double, Xa As Double, FG As Variant, Fi As Variant
Dim Fg1 As Double, Fg2 As Double

' Primero miramos si a y x son negativos
If a <= 0 And x <= 0 Then
   ' a y x son negativos => la función no está definida
   F_Gamma_Sup = "(" & ChrW(945) & ", x) Fuera de rango"
   Exit Function
End If

' Límites asintóticos
If x > xAsintSup Then
   'FG = F_Gamma(a)
   FG = 1
   Fg1 = FG / (a - 1): Fg2 = Fg1 / (a - 2)
   F_Gamma_Sup = x ^ (a - 1) * Exp(-x) * (1 + FG / Fg1 / x + FG / Fg2 / x / x)
   Exit Function
End If

' Ahora calculamos la parte entera de a (N)
N = Fix(a)
If a <= 0 And Abs(N - a) < Eps Then
   ' a es cero o un entero negativo
   MM = 1 - N
   F_Gamma_Sup = x ^ (1 - MM) * F_En(x, MM, NSumandos)
   Exit Function
End If

FG = F_Gamma(a)
If IsNumeric(FG) Then
   Fi = F_Gamma_Inf(a, x, NSumandos)
   If IsNumeric(Fi) Then
      F_Gamma_Sup = FG - Fi
      Exit Function
   End If
   F_Gamma_Sup = "(" & ChrW(945) & ", x) Fuera de rango"
Else
   F_Gamma_Sup = FG
End If

End Function


' FUNCIÓN GAMMA INCOMPLETA INFERIOR NORMALIZADA P

Public Function F_P_Gamma(a As Double, x As Double, _
Optional NSumandos As Integer = 1000, _
Optional xAsintSup As Double = 150, _
Optional Eps As Double = 0.0000001) As Variant

' Función Gamma Incompleta Inferior normalizada P(a, x)
' NSumandos es el número de términos de la serie que vamos a sumar
' xAsintSup valor límite superior
' Eps error permitido al comparar con cero

' Llama a la rutina Mi_es_Entero
' Llama a la función F_GammaA

Dim N As Integer, Tipo As Integer, i As Integer
Dim Sumando As Double, Suma As Double

' Miramos si es entero positivo el argumento
Call Mi_es_Entero(a, N, Tipo)

If Tipo = 0 Or Tipo = -1 Then
   ' La función vale 1
   F_P_Gamma = 1
   Exit Function
ElseIf Tipo = 1 Then
   ' El argumento a es entero y aplicamos
   ' otra formulación para mejorar convergencia
   Suma = 1: Sumando = 1
   For i = 1 To N - 1
     Sumando = Sumando * x / i
     Suma = Suma + Sumando
   Next
   F_P_Gamma = 1 - Suma * Exp(-x)
   Exit Function
End If

If x >= xAsintSup Then
   ' Para este valor la función gamma inferior se calcula como la gamma
   F_P_Gamma = 1
   Exit Function
End If

F_P_Gamma = x ^ a * F_GammaA(a, x, NSumandos, Eps)
End Function


' FUNCIÓN GAMMA INCOMPLETA SUPERIOR NORMALIZADA Q

Public Function F_Q_Gamma(a As Double, x As Double, _
Optional NSumandos As Integer = 1000, _
Optional xAsintSup As Double = 150, _
Optional Eps As Double = 0.0000001) As Variant

' Función Gamma Incompleta Superior normalizada Q(a, x)
' NSumandos es el número de términos de la serie que vamos a sumar
' xAsintSup valor límite superior
' Eps error permitido al comparar con cero

' Llama a la función F_P_Gamma

F_Q_Gamma = 1 - F_P_Gamma(a, x, NSumandos, xAsintSup, Eps)
End Function


' FUNCIÓN BETA

Public Function F_Beta(x As Double, y As Double) As Variant
' Calcula la función Beta (x,y) para cualesquiera argumentos reales
' Llama a la función F_Gamma

Dim Gx As Double, Gy As Double, GxMy As Double
Dim N As Integer, xMy As Double, Eps As Double

Eps = 0.0000001

' Primero vemos si x o y son números enteros negativos o cero
N = Fix(x)
If x <= 0 And Abs(N - x) < Eps Then
   ' x es cero o un entero negativo -> Beta = +- infinito
   F_Beta = ChrW(177) & " " & ChrW(8734)
   Exit Function
End If

N = Fix(y)
If y <= 0 And Abs(N - y) < Eps Then
   ' y es cero o un entero negativo -> Beta = +- infinito
   F_Beta = ChrW(177) & " " & ChrW(8734)
   Exit Function
End If

If x + y < 0 Then
   xMy = x + y - Eps
Else
   xMy = x + y
End If

If Abs(x + y) <= Eps Then
   ' Primera prueba de valor nulo
   ' x+y es cero -> Beta = +- infinito
   F_Beta = ChrW(177) & " " & ChrW(8734)
   Exit Function
End If

N = Fix(xMy)
If xMy <= 0 And Abs(N - xMy) < Eps Then
   ' x+y es cero o un entero negativo -> Beta = +- infinito
   F_Beta = ChrW(177) & " " & ChrW(8734)
   Exit Function
End If

' Calcular Beta (x,y) mediante su relación con la función Gamma

Gx = F_Gamma(x)
Gy = F_Gamma(y)
GxMy = F_Gamma(x + y)

F_Beta = Gx * Gy / GxMy
End Function


' FUNCIÓN BETA INCOMPLETA

Public Function F_BetaI(a As Double, b As Double, x As Double, Optional Nsumandos As Long = 100) As Variant
' Calcula la función Beta Incompleta para cualesquiera argumentos reales
' Llama a la función F_Beta
' Llama a la rutina Mi_es_Entero

Dim Suma As Double, Sumando As Double, Factor As Double
Dim i As Long, Na As Integer, Nb As Integer, Nx As Integer
Dim Tipoa As Integer, Tipob As Integer, Tipox As Integer
Dim SS() As Double, Eps As Double

ReDim SS(0 To Nsumandos)
Eps = 0.0000001

If Abs(x) < Eps And a < 0 Then
  ' x es cero y a es negativo -> BetaI = +- infinito
  F_BetaI = ChrW(177) & " " & ChrW(8734)
  Exit Function
End If

If Abs(x - 1) < Eps Then
  ' x es 1 -> llama a la función F_Beta
  F_BetaI = F_Beta(a, b)
  Exit Function
End If

Call Mi_es_Entero(x, Nx, Tipox)
Call Mi_es_Entero(a, Na, Tipoa)
Call Mi_es_Entero(b, Nb, Tipob)

If Tipoa <= 0 Then
  ' a es entero negativo -> BetaI = +- infinito
  F_BetaI = ChrW(177) & " " & ChrW(8734)
  Exit Function
End If

If x < 0 And Tipoa = 2 Then
  ' x es negativo y a no es entero -> BetaI = Imaginario
  F_BetaI = "Imaginaria"
  Exit Function
End If

If x > 1 And Tipob = 2 Then
  ' x es mayor que 1 y b no es entero -> BetaI = Imaginario
  F_BetaI = "Imaginaria"
  Exit Function
End If

' Calcular BetaI (a,b,x) mediante desarrollo en serie

Suma = 1 / a: Sumando = 1 / a
SS(0) = Suma

For i = 1 To Nsumandos
    Factor = (a + i - 1) * (i - b) / i / (a + i) * x
    Sumando = Sumando * Factor
    Suma = Suma + Sumando
    SS(i) = Sumando
Next

F_BetaI = x ^ a * Suma

End Function


' FUNCIÓN BETA REGULARIZADA

Public Function F_BetaI_Regularizada(xx As Double, MM As Long, NN As Long, Optional Nsumandos = 100) As Variant
' Este desarrollo vale para 0<=xx<1 y
' mm y nn enteros
' Para mejorar la convergencia aplicamos la propiedad: I(x,m,n)=1-I(1-x,n,m)
' Llama a la función N_Combinatorio

Dim i As Long
Dim Sumando As Double
Dim x As Double
Dim m As Long, N As Long
Dim kInvertir As Integer

If Abs(xx - 1) < 0.0000001 Then xx = 1

If xx < 0 Or xx > 1 Then
   F_BetaI_Regularizada = "x debe ser >=0 y <=1"
   Exit Function
End If

If MM < 0 Or NN < 0 Then
   F_BetaI_Regularizada = "m y n deben ser >=0"
   Exit Function
End If

If xx <= 0.5 Then
   x = xx
   N = NN
   m = MM
   kInvertir = 0
Else
   x = 1 - xx
   N = MM
   m = NN
   kInvertir = 1
End If

' Calcular BetaI_Regularizada (x,m,n) mediante desarrollo en serie

Sumando = N_Combinatorio(N + m - 1, m) * x ^ m
F_BetaI_Regularizada = Sumando

For i = m + 1 To m + 1 + Nsumandos
    Sumando = Sumando * (i + N - 1) / i * x
    F_BetaI_Regularizada = F_BetaI_Regularizada + Sumando
Next

F_BetaI_Regularizada = F_BetaI_Regularizada * (1 - x) ^ N

If kInvertir = 1 Then
   F_BetaI_Regularizada = 1 - F_BetaI_Regularizada
End If

End Function


' FUNCIÓN DIGAMMA

Public Function F_Digamma(xx As Double, Optional Nsumandos As Long = 100000) As Variant
'
' Obtiene el valor de la función Digamma. Derivada logarítmica de la función Gamma
' Se utiliza la relación asintótica (E) del apartado B.1.7
'
Dim x As Double, Pi As Double, Gamma As Double
Dim xf As Double, xf2 As Double, xf3 As Double, xf5 As Double, xf7 As Double
Dim y  As Double
Dim i As Long, N As Integer
Dim Caso As String, Eps As Double, aa As Integer
Dim a1 As Double, a2 As Double, a3 As Double, a4 As Double
Dim Menor_de_10 As Integer

Eps = 0.0000001
Pi = 3.14159265358979
Gamma = 0.577215664901533
a1 = 1 / 24: a2 = 37 / 5760
a3 = 3.55248291446208E-03
a4 = 3.95355744897303E-03

Caso = "Positivo"

' Primero calculamos la parte entera de xx (N)
N = Fix(xx)
If xx <= 0 And Abs(N - xx) < Eps Then
   ' x es cero o un entero negativo -> Digamma = +- infinito
   F_Digamma = ChrW(177) & " " & ChrW(8734)
   Exit Function
End If

If xx > 0 And Abs(N - xx) < Eps Then
   ' x es un entero positivo
   F_Digamma = -Gamma
   For i = 1 To N - 1
     F_Digamma = F_Digamma + 1 / i
   Next
   Exit Function
End If

If xx < 0 Then
   x = 1 - xx
   Caso = "Negativo"
Else
   x = xx
   Caso = "Positivo"
End If

If x > 10 Then
   Menor_de_10 = 0
Else
   Menor_de_10 = 1
   y = x
   x = x + 10
End If

xf = 1 / (x - 0.5): xf2 = xf * xf
xf3 = xf * xf2: xf5 = xf3 * xf2: xf7 = xf5 * xf2

F_Digamma = 1 / xf + a1 * xf - a2 * xf3 + a3 * xf5 - a4 * xf7
F_Digamma = Log(F_Digamma)

If Menor_de_10 = 1 Then
   For i = 9 To 0 Step -1
       F_Digamma = F_Digamma - 1 / (y + i)
   Next
End If

If Caso = "Negativo" Then
   ' Fórmula de reflexión
   F_Digamma = F_Digamma - Pi / Tan(Pi * (1 - x))
End If

End Function


' FUNCIÓN DIGAMMAC

Public Function F_DigammaC(xx As Double, Optional Nsumandos As Long = 100000) As Variant
'
' Obtiene el valor de la función Digamma. Derivada logarítmica de la función Gamma
' Se utiliza el desarrollo en serie (C) del apartado B.1.7
'
Dim x As Double, xf As Double, Pi As Double, Gamma As Double
Dim i As Long, N As Integer
Dim Caso As String, Eps As Double, aa As Integer

Eps = 0.0000001
Pi = 3.14159265358979
Gamma = 0.577215664901533
Caso = "Positivo"

' Primero calculamos la parte entera de xx (N)
N = Fix(xx)
If xx <= 0 And Abs(N - xx) < Eps Then
   ' x es cero o un entero negativo -> Digamma = +- infinito
   F_DigammaC = ChrW(177) & " " & ChrW(8734)
   Exit Function
End If

If xx > 0 And Abs(N - xx) < Eps Then
   ' x es un entero positivo
   F_DigammaC = -Gamma
   For i = 1 To N - 1
     F_DigammaC = F_DigammaC + 1 / i
   Next
   Exit Function
End If

If xx < 0 Then
   x = 1 - xx
   N = Fix(x)
   x = x - N
   Caso = "Negativo"
Else
   x = xx - N
   Caso = "Positivo"
End If

F_DigammaC = -Gamma
For i = 1 To Nsumandos
  F_DigammaC = F_DigammaC + (x - 1) / i / (i + x - 1)
Next

If N >= 1 Then
   For i = 1 To N
       F_DigammaC = F_DigammaC + 1 / x
       x = x + 1
   Next
End If

If Caso = "Negativo" Then
   ' Fórmula de reflexión
   F_DigammaC = F_DigammaC - Pi / Tan(Pi * (1 - x))
End If

End Function


' FUNCIÓN DIGAMMAZ

Public Function F_DigammaZ(xx As Double) As Variant
'
' Obtiene el valor de la función Digamma. Derivada logarítmica de la función Gamma
' Se utiliza el desarrollo en serie (D) del apartado B.1.7
' Llama a la función auxiliar S_Zeta_Entera

Dim x As Double, Pi As Double, Gamma As Double, Factor As Double
Dim i As Integer, N As Integer, Signo As Integer
Dim Caso As String, Eps As Double, Residuo As Double
Dim Menor1 As Integer, Mayor1 As Integer

Eps = 0.0000001
Pi = 3.14159265358979
Gamma = 0.577215664901533
Caso = "Positivo"

' Primero calculamos la parte entera de xx (N)
N = Fix(xx)
If xx <= 0 And Abs(N - xx) < Eps Then
   ' x es cero o un entero negativo -> Digamma = +- infinito
   F_DigammaZ = ChrW(177) & " " & ChrW(8734)
   Exit Function
End If

If xx > 0 And Abs(N - xx) < Eps Then
   ' x es un entero positivo
   F_DigammaZ = -Gamma
   For i = 1 To N - 1
     F_DigammaZ = F_DigammaZ + 1 / i  ' (Ecuación A)
   Next
   Exit Function
End If

x = xx
If xx < 0 Then
   ' En este caso aplicaremos la fórmula de los complementos
   x = 1 - xx
   N = Fix(x)
   ' x = x - N
   Caso = "Negativo"
End If

Menor1 = 0: Mayor1 = 0
If x > 0 And x < 1 Then
   ' Al llegar aquí x es un número positivo "no entero"
    Menor1 = 1
    x = x + 1
ElseIf x > 1 Then
   ' Caso en que x>1
    Mayor1 = 1
    x = x - N + 1
End If

F_DigammaZ = -Gamma: Signo = -1: Factor = 1
For i = 2 To 47
  ' Primero sumamos los términos 2 a 47 de Zeta(i)>1
  ' Los posteriores valen aproximadamente 1 y se agrupan
  ' en una progresión geométrica (ecuación D)
  Signo = -Signo
  Factor = Factor * (x - 1)
  F_DigammaZ = F_DigammaZ + Signo * S_Zeta_Entera(i) * Factor
Next
' A esta suma le sumamos los infinitos sumandos
' restantes, que forman una progresión geométrica
' a0=(-x)^49  r=-(x-1)
Residuo = (x - 1) ^ 47 / x  ' Suma de la progresión geométrica
F_DigammaZ = F_DigammaZ + Residuo

If Menor1 = 1 Then
   ' x era menor que 1
   x = x - 1
   F_DigammaZ = F_DigammaZ - 1 / x
End If

If N >= 1 Then
   For i = 1 To N - 1
       F_DigammaZ = F_DigammaZ + 1 / x
       x = x + 1
   Next
End If

If Caso = "Negativo" Then
   ' Fórmula de reflexión (ecuación B)
   F_DigammaZ = F_DigammaZ - Pi / Tan(Pi * xx)
End If

End Function


