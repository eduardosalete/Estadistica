
' FUNCIÓN HIPERGEOMÉTRICA GENERAL pFq

Public Function F_pFq(a() As Double, b() As Double, x As Double, _
                      Optional NSumandos As Integer = 32000, Optional Eps As Double = 0.0000001) As Variant
' Calcula la función Hipergeométrica pFq {a(1...p), b(1...q)| x)
' Llama a la rutina Mi_es_Entero

Dim p As Integer, q As Integer, Vacio As Integer
Dim ii As Integer, jj As Integer, kk As Integer
Dim Cero_a As Integer, Cero_b As Integer
Dim Nega_a As Integer, Nega_b As Integer
Dim Suma_a As Double, Suma_b As Double
Dim Suma As Double, Sumando() As Double, Factor() As Double
Dim Fact_a As Double, Fact_b As Double

ReDim Sumando(0 To NSumandos), Factor(0 To NSumandos)

'
' ---> Tratamiento de los vectores a() y b()
'
' Obtenemos las longitudes declaradas, pero habrá que acotarlos
p = UBound(a)
q = UBound(b)

' Recorremos los vectores para asegurarnos
' de que no ha introducido valores nulos al final
jj = p: kk = 0
For ii = jj To 1 Step -1
    ' Obtenemos la longitud real del vector a
    If a(ii) <> 0 Then
       kk = ii
       Exit For
    End If
Next ii
p = kk
If p = 0 Or a(1) = 0 Then
   F_pFq = 1
   Exit Function
End If

jj = q: kk = 0
For ii = jj To 1 Step -1
    ' Obtenemos la longitud real del vector b
    If b(ii) <> 0 Then
       kk = ii
       Exit For
    End If
Next ii
q = kk
If q = 0 Or b(1) = 0 Then
   F_pFq = ChrW(8734)
   Exit Function
End If
    
' Caso en que alguno de los vectores esté vacío
Vacio = 0
If p = 0 Or q = 0 Then
   Vacio = 1
   p = p + 1: a(p) = 1
   q = q + 1: b(q) = 1
End If

' Rellenamos los vectores
Suma_a = 0
For ii = 1 To p
  Suma_a = Suma_a + a(ii)
Next

Suma_b = 0
For ii = 1 To q
  Suma_b = Suma_b + b(ii)
Next

'
' Detección de ceros y enteros negativos
'
Cero_a = 0: Nega_a = 0
For ii = 1 To p
  Call Mi_es_Entero(a(ii), jj, kk)
  If kk = 0 Then
     Cero_a = ii
     Exit For
   End If
  If kk = -1 Then
     Nega_a = -jj
     Exit For
   End If
Next

Cero_b = 0: Nega_b = 0
For ii = 1 To q
  Call Mi_es_Entero(b(ii), jj, kk)
  If kk = 0 Then
     Cero_b = ii
     Exit For
   End If
  If kk = -1 Then
     Nega_b = -jj
     Exit For
   End If
Next

'
' ---> Tratamiento de casos especiales
'
If Cero_a * Cero_b <> 0 Then
   ' Hay ceros en numerador y denominador
   F_pFq = "Indeterminada (0/0)"
   Exit Function
End If

If Cero_a <> 0 And Cero_b = 0 Then
   ' Ceros sólo en numerador
   F_pFq = 0
   Exit Function
End If

If Cero_a = 0 And Cero_b <> 0 Then
   ' Ceros sólo en denominador
   F_pFq = ChrW(8734) & " b(i) nulo"
   Exit Function
End If

'If Nega_b < Nega_a Or (Nega_b > 0 And Nega_a = 0) Then
 '  F_pFq = ChrW(8734)
'   Exit Function
'End If

'
' ---> Tratamiento de casos especiales
'
If p > q + 1 Then
   ' Divergencia por mayor grado del numerador
   F_pFq = ChrW(8734) & " (p > q+1)"
   Exit Function
End If

' Redondeo de valores próximos a 1
If x > 0 And Abs(x - 1) < Eps Then x = 1
If x < 0 And Abs(x + 1) < Eps Then x = -1

If p = q + 1 And Abs(x) > 1 Then
   ' Divergencia por estar fuera del círculo de convergencia
   F_pFq = ChrW(8734) & " (fuera del c. conv.)"
   Exit Function
End If

'If p = q + 1 And Abs(x) = 1 And Suma_b - Suma_a < 0 Then
   ' Divergencia por no cumplir la cond. de convergencia en el círculo
   ' La he quitado porque es suficiente pero no necesaria
'   If x = 1 Or Suma_b - Suma_a < -1 Then
'      F_pFq = ChrW(8734) & " (cond. conv.)"
'   End If
'   Exit Function
'End If

'
' ---> Suma de la serie (se han eliminado divergencias)
'
Sumando(0) = 1: Suma = Sumando(0): Factor(0) = 1
For ii = 1 To NSumandos
   Fact_a = 1: Fact_b = 1
   For jj = 1 To p
     Fact_a = Fact_a * (a(jj) + ii - 1)
   Next
   ' Nos salimos del For si Fact_a=0 porque este caso
   ' se reduce a un polinomio
   If Abs(Fact_a) < Eps Then Exit For
   For jj = 1 To q
     Fact_b = Fact_b * (b(jj) + ii - 1)
   Next
   Factor(ii) = Fact_a / Fact_b / ii * x
   Sumando(ii) = Sumando(ii - 1) * Factor(ii)
   Suma = Suma + Sumando(ii)
Next

F_pFq = Suma

' Dejamos en orden los vectores de entrada
If Vacio = 1 Then
   a(p) = 0
   b(q) = 0
End If
End Function


' FUNCIÓN HIPERGEOMÉTRICA DE GAUSS 2F1

Public Function F_HG_Gauss(aa As Double, bb As Double, cc As Double, xx As Double, _
                          Optional NSumandos As Integer = 32000, Optional Eps As Double = 0.0000001) As Variant
' Calcula la función Hipergeométrica de Gauss 2F1(a,b;c|x)
' Para x>0.5 y a+b-c entero ajusta automáticamente el número de sumandos
' En este último caso no se aconseja para x>0.999
' Llama a la rutina Mi_es_Entero
' Llama a la función F_Gamma
' Llama a la rutina F_HG_Gauss_Especiales
' Llama a la función F_pFq

Dim va(1 To 2) As Double, vb(1 To 1) As Double
Dim ma As Integer, Tipoa As Integer
Dim mb As Integer, Tipob As Integer
Dim mc As Integer, Tipoc As Integer
Dim mp As Integer, Tipop As Integer
Dim mq As Integer, Tipoq As Integer
Dim ms As Integer, Tipos As Integer
Dim NSumandos2 As Integer
Dim Pi As Double
Dim Factor As Double, Sumando As Double, Suma As Double
Dim Sumando1 As Double, Sumando2 As Double
Dim Factor1 As Double, Factor2 As Double
Dim p As Double, q As Double, s As Double
Dim a As Double, b As Double, c As Double, x As Double
Dim ii As Integer
Dim Mi_funcion As Variant, Calculado As String

Pi = 3.14159265358979

' Casos particulares
If Abs(aa + cc) < Eps And Abs(bb) < Eps Then
   ' F(aa,0,-aa,xx)=1
   F_HG_Gauss = 1
End If

If Abs(bb + cc) < Eps And Abs(aa) < Eps Then
   ' F(0,bb,-bb,xx)=1
   F_HG_Gauss = 1
   Exit Function
End If

If Abs(bb) < Eps Then
   ' F(aa,0,cc,xx)=1
   F_HG_Gauss = 1
   Exit Function
End If

If Abs(aa) < Eps Then
   ' F(0,bb,cc,xx)=1
   F_HG_Gauss = 1
   Exit Function
End If

If Abs(bb - cc) < Eps And Abs(aa) < Eps Then
   ' F(0,bb,bb,xx)=1
   F_HG_Gauss = 1
   Exit Function
End If

If Abs(aa + bb) < Eps And Abs(bb - cc) < Eps Then
   ' F(-aa,aa,aa,xx)=(1-x)^aa
   F_HG_Gauss = (1 - xx) ^ Abs(aa)
   Exit Function
End If

If Abs(aa + bb) < Eps And Abs(aa - cc) < Eps Then
   ' F(aa,-aa,aa,xx)=(1-x)^aa
   F_HG_Gauss = (1 - xx) ^ Abs(aa)
   Exit Function
End If

' Rutina Mi_es_Entero:
' Tipo =  1 si es entero positivo.
'      =  0 si vale 0.
'      = -1 si es entero negativo.
'      =  2 si no es entero.
Call Mi_es_Entero(cc, mc, Tipoc)
If Tipoc <= 0 Then
   ' c es un entero negativo y la función tiende a infinito
   F_HG_Gauss = ChrW(8734)
   Exit Function
End If

Call Mi_es_Entero(aa, ma, Tipoa)
If Tipoa = -1 Then
   ' En este caso es un polinomio
   Sumando = 1: Suma = Sumando
   For ii = 1 To Abs(ma)
       Factor = (bb + ii - 1) * (ma + ii - 1) / (cc + ii - 1) * xx / ii
       Sumando = Sumando * Factor
       Suma = Suma + Sumando
   Next
   F_HG_Gauss = Suma
   Exit Function
End If

Call Mi_es_Entero(bb, mb, Tipob)
If Tipob = -1 Then
   ' En este caso también es un polinomio
   Sumando = 1: Suma = Sumando
   For ii = 1 To Abs(mb)
       Factor = (aa + ii - 1) * (mb + ii - 1) / (cc + ii - 1) * xx / ii
       Sumando = Sumando * Factor
       Suma = Suma + Sumando
   Next
   F_HG_Gauss = Suma
   Exit Function
End If

If Abs(aa - bb) < Eps And Abs(aa - cc) < Eps And (bb - cc) < Eps Then
   ' F(aa,aa,aa,xx)=(1-xx)^(-aa)
   If xx > 1 And aa > 0 Then
         F_HG_Gauss = "Imaginario"
   ElseIf Abs(xx) < 0 And aa > 0 Then
         F_HG_Gauss = ChrW(8734)
   Else
         F_HG_Gauss = (1 - xx) ^ (-aa)
   End If
   Exit Function
End If

' Más casos especiales
'(Abramowitz & Stegun, pp. 556, 557)
Calculado = "NO"
Call F_HG_Gauss_Especiales(aa, bb, cc, xx, Tipoa, Tipob, Tipoc, ma, mb, mc, Mi_funcion, Calculado)
a = 1
If Calculado = "SI" Then
   F_HG_Gauss = Mi_funcion
   Exit Function
End If

Factor = 1
x = xx
a = aa
b = bb
c = cc
If xx < 0 And Abs(xx) > 1 Then
   ' Valores negativos menores que -1
   x = xx / (xx - 1)
   If Abs(x) < 1 Then
      ' Expresión a.2
      Factor = (1 - xx) ^ (-aa)
      a = aa
      b = cc - bb
      c = cc
   Else
      x = xx
   End If
End If

s = bb - aa: Call Mi_es_Entero(s, ms, Tipos)
If x < 0 And Abs(x) > 1 And Abs(Tipos) > 1 Then
   ' Valores negativos menores que -1
   ' y b-a no es un número entero, aplicamos una
   ' transformación para mejorar la convergencia
   va(1) = aa: va(2) = 1 - cc + aa: vb(1) = 1 - bb + aa: x = 1 / x
   p = bb: Call Mi_es_Entero(p, mp, Tipop)
   q = cc - aa: Call Mi_es_Entero(q, mq, Tipoq)
   If Tipop = -1 Or Tipoq = -1 Then
      Sumando1 = 0
   Else
      Factor1 = (-x) ^ (-aa) * F_Gamma(cc) * F_Gamma(bb - aa) / F_Gamma(bb) / F_Gamma(cc - aa)
      Sumando1 = Factor1 * F_pFq(va, vb, x, NSumandos)
   End If
   
   va(1) = bb: va(2) = 1 - cc + bb: vb(1) = 1 - aa + bb
   p = aa: Call Mi_es_Entero(p, mp, Tipop)
   q = cc - bb: Call Mi_es_Entero(q, mq, Tipoq)
   If Tipop = -1 Or Tipoq = -1 Then
      Sumando2 = 0
   Else
      Factor2 = (-xx) ^ (-bb) * F_Gamma(cc) * F_Gamma(aa - bb) / F_Gamma(aa) / F_Gamma(cc - bb)
      Sumando2 = Factor2 * F_pFq(va, vb, x, NSumandos)
   End If
   F_HG_Gauss = Sumando1 + Sumando2
   Exit Function
End If

If Abs(x - 1) < Eps Then
   If c > a + b Then
      ' Teorema de Gauss
      F_HG_Gauss = F_Gamma(c) * F_Gamma(c - a - b) / _
                   F_Gamma(c - a) / F_Gamma(c - b)
      F_HG_Gauss = Factor * F_HG_Gauss
   Else
      F_HG_Gauss = ChrW(8734)
   End If
   Exit Function
End If

If x > 1 Then
  F_HG_Gauss = "Imaginaria"
  Exit Function
End If

s = a + b - c: Call Mi_es_Entero(s, ms, Tipos)

If x >= -1 And (x <= 0.5 Or Tipos <= 1) Then
  ' Estamos dentro del círculo de convergencia
  ' o a+b-c es un número entero
  NSumandos2 = 10000
  If NSumandos > NSumandos2 Then NSumandos2 = NSumandos
  va(1) = a: va(2) = b: vb(1) = c
  F_HG_Gauss = F_pFq(va, vb, x, NSumandos2)
  F_HG_Gauss = Factor * F_HG_Gauss
  Exit Function
End If

' Si hemos llegado aquí, estamos en el caso 0.5<x< 1
' y además a+b-c no es un número entero
' Aplicamos la transformación A.4

va(1) = a: va(2) = b: vb(1) = a + b - c + 1: x = 1 - x
p = c - a: Call Mi_es_Entero(p, mp, Tipop)
q = c - b: Call Mi_es_Entero(q, mq, Tipoq)
If Tipop = -1 Or Tipoq = -1 Then
   Sumando1 = 0
Else
   Factor1 = F_Gamma(c) * F_Gamma(c - a - b) / F_Gamma(c - a) / F_Gamma(c - b)
   Sumando1 = Factor1 * F_pFq(va, vb, x, NSumandos)
End If

va(1) = c - a: va(2) = c - b: vb(1) = c - a - b + 1
p = a: Call Mi_es_Entero(p, mp, Tipop)
q = b: Call Mi_es_Entero(q, mq, Tipoq)
If Tipop = -1 Or Tipoq = -1 Then
   Sumando2 = 0
Else
   Factor2 = x ^ (c - a - b) * F_Gamma(c) * F_Gamma(a + b - c) / F_Gamma(a) / F_Gamma(b)
   Sumando2 = Factor2 * F_pFq(va, vb, x, NSumandos)
End If
F_HG_Gauss = Factor * (Sumando1 + Sumando2)

End Function


' RUTINA HIPERGEOMÉTRICA DE GAUSS 2F1 (casos especiales)

Sub F_HG_Gauss_Especiales(a As Double, b As Double, c As Double, x As Double, _
                          Tipoa As Integer, Tipob As Integer, Tipoc As Integer, _
                          ma As Integer, mb As Integer, mc As Integer, _
                          Mi_funcion As Variant, Calculado As String)
' En esta rutina se estudian los casos especiales contemplados en Abramowitz & Stegun, pp. 556, 557.
' En todos los casos se indican las ecuaciones con sus números
' Tipo =  1 si es entero positivo.
'      =  0 si vale 0.
'      = -1 si es entero negativo.
'      =  2 si no es entero.
' Llama a la rutina Mi_es_Entero
' Llama a la función F_Gamma
' Llama a la función F_DigammaZ

Dim Eps As Double, xx As Double, Tan As Double, yy As Double, y As Double
Dim Pi As Double
Dim t As Double, Tipot As Integer, nt As Integer
Dim u As Double, Tipou As Integer, nu As Integer
Dim f1 As Integer, f2 As Integer
Dim Sumando1 As Double, Sumando2 As Double

Eps = 0.0000001
Pi = 3.14159265358979
Calculado = "NO"

' Ecuación 15.1.3
If (Tipoa = 1 And ma = 1) And (Tipob = 1 And mb = 1) And (Tipoc = 1 And mc = 2) And x < 1 And Abs(x) > Eps Then
  Calculado = "SI"
  Mi_funcion = -1 / x * Log(1 - x)
  Exit Sub
End If

' Ecuación 15.1.4
If (Abs(a - 1 / 2) < Eps) And (Tipob = 1 And mb = 1) And (Abs(c - 3 / 2) < Eps) And (x < 1 And x > Eps) Then
  Calculado = "SI"
  xx = Sqr(x)
  Mi_funcion = 1 / 2 * 1 / xx * Log((1 + xx) / (1 - xx))
  Exit Sub
End If

' Ecuación 15.1.5
If (Abs(a - 1 / 2) < Eps) And (Tipob = 1 And mb = 1) And (Abs(c - 3 / 2) < Eps) And (-x > Eps) Then
  Calculado = "SI"
  xx = Sqr(-x)
  Mi_funcion = 1 / xx * Atn(xx)
  Exit Sub
End If

' Ecuación 15.1.6a
If (Abs(a - 1 / 2) < Eps) And (Abs(b - 1 / 2) < Eps) And (Abs(c - 3 / 2) < Eps) And (x < 1 And x > Eps) Then
  Calculado = "SI"
  xx = Sqr(x)
  Tan = xx / Sqr(1 - x)
  Mi_funcion = 1 / xx * Atn(Tan)
  Exit Sub
End If

' Ecuación 15.1.6b
If (Tipoa = 1 And ma = 1) And (Tipob = 1 And mb = 1) And (Abs(c - 3 / 2) < Eps) And (x < 1 And x > Eps) Then
  Calculado = "SI"
  xx = Sqr(x)
  Tan = xx / Sqr(1 - x)
  Mi_funcion = 1 / xx * Atn(Tan) / Sqr(1 - x)
  Exit Sub
End If

' Ecuación 15.1.7a
If (Abs(a - 1 / 2) < Eps) And (Abs(b - 1 / 2) < Eps) And (Abs(c - 3 / 2) < Eps) And (-x > Eps) Then
  Calculado = "SI"
  xx = Sqr(-x)
  Mi_funcion = 1 / xx * Log(xx + Sqr(1 - x))
  Exit Sub
End If

' Ecuación 15.1.7b
If (Tipoa = 1 And ma = 1) And (Tipob = 1 And mb = 1) And (Abs(c - 3 / 2) < Eps) And (-x > Eps) Then
  Calculado = "SI"
  xx = Sqr(-x)
  Mi_funcion = 1 / xx * Log(xx + Sqr(1 - x)) / Sqr(1 - x)
  Exit Sub
End If

' Ecuación 15.1.8
If (Abs(b - c) < Eps) And (x < 1) Then
  Calculado = "SI"
  Mi_funcion = (1 - x) ^ (-a)
  Exit Sub
End If

' Ecuación 15.1.9
If (Abs(b - a - 0.5) < Eps) And c = 0.5 And x > 0 And Abs(a - 0.5) > Eps Then
  Calculado = "SI"
  xx = Sqr(x)
  yy = -2 * a
  Mi_funcion = 0.5 * ((1 + xx) ^ yy + (1 - xx) ^ yy)
  Exit Sub
End If

' Ecuación 15.1.10
If (Abs(b - a - 0.5) < Eps) And c = 1.5 And (Abs(x) > Eps) And Abs(a - 0.5) > Eps Then
  Calculado = "SI"
  xx = Sqr(x)
  yy = 1 - 2 * a
  Mi_funcion = 0.5 / (xx * yy) * ((1 + xx) ^ yy - (1 - xx) ^ yy)
  Exit Sub
End If

' Ecuación 15.1.11
If a < 0 And (Abs(b + a) < Eps) And c = 0.5 And x <= 0 Then
  Calculado = "SI"
  xx = Sqr(-x)
  yy = Sqr(1 - x)
  Mi_funcion = 0.5 * ((yy + xx) ^ (2 * a) + (yy - xx) ^ (2 * a))
  Exit Sub
End If

' Ecuación 15.1.12
If (Abs(b + a - 1) < Eps) And c = 0.5 And x <= 0 Then
  Calculado = "SI"
  xx = Sqr(-x)
  yy = Sqr(1 - x)
  Mi_funcion = 0.5 / yy * ((yy + xx) ^ (2 * a - 1) + (yy - xx) ^ (2 * a - 1))
  Exit Sub
End If

' Ecuación 15.1.13a
If (Abs(b - a - 0.5) < Eps) And (Abs(c - 1 - 2 * a) < Eps) And (x < 1) Then
  Calculado = "SI"
  Mi_funcion = 2 ^ (2 * a) * (1 + Sqr(1 - x)) ^ (-2 * a)
  Exit Sub
End If

' Ecuación 15.1.13b
If (Abs(b - a + 0.5) < Eps) And (Abs(2 * a - c - 1) < Eps) And (x < 1) Then
  Calculado = "SI"
  y = a - 1
  Mi_funcion = 2 ^ (2 * y) * (1 + Sqr(1 - x)) ^ (-2 * y)
  Mi_funcion = Mi_funcion / (Sqr(1 - x))
  Exit Sub
End If

' Ecuación 15.1.14
If (Abs(b - a - 0.5) < Eps) And (Abs(c - 2 * a) < Eps) And (x < 1) Then
  Calculado = "SI"
  Mi_funcion = 2 ^ (2 * a - 1) / Sqr(1 - x) * (1 + Sqr(1 - x)) ^ (1 - 2 * a)
  Exit Sub
End If

' Ecuaciones de seno cuadrado
If x >= 0 And x < 1 Then
   xx = Sqr(x):   Tan = xx / Sqr(1 - x): y = Atn(Tan)
   
   ' Ecuación 15.1.15
   If (Abs(a + b - 1) < Eps) And c = 1.5 And x > 0 And (Abs(2 * a - 1) > Eps) Then
      Calculado = "SI"
      Mi_funcion = Sin((2 * a - 1) * y) / ((2 * a - 1) * Sin(y))
      Exit Sub
   End If
   
   ' Ecuación 15.1.16
   If (Abs(a + b - 2) < Eps) And c = 1.5 And x > 0 And (Abs(a - 1) > Eps) Then
      Calculado = "SI"
      Mi_funcion = Sin((2 * a - 2) * y) / ((a - 1) * Sin(2 * y))
      Exit Sub
   End If
   
   ' Ecuación 15.1.17
   If (a <= 0) And (Abs(a + b) < Eps) And c = 0.5 Then
      Calculado = "SI"
      Mi_funcion = Cos(2 * a * y)
      Exit Sub
   End If
   
   ' Ecuación 15.1.18
   If (Abs(a + b - 1) < Eps) And c = 0.5 Then
      Calculado = "SI"
      Mi_funcion = Cos((2 * a - 1) * y) / Cos(y)
      Exit Sub
   End If
   
End If

' Ecuación en tangente cuadrado: 15.1.19
If (Abs(b - a - 0.5) < Eps) And (Abs(c - 0.5) < Eps) And (x < 0) Then
  Calculado = "SI"
  xx = Sqr(-x): y = Atn(xx)
  Mi_funcion = (Cos(y)) ^ (2 * a) * Cos(2 * a * y)
  Exit Sub
End If

' Valores especiales del argumento

' Ecuación 15.1.20a
If (Tipoc <= 0 Or (c <= a + b)) And Abs(x - 1) < Eps Then
     Mi_funcion = ChrW(8734)
     Exit Sub
End If

' Ecuación 15.1.20b
If Tipoc > 0 And (c > a + b) And Abs(x - 1) < Eps Then
  Calculado = "SI"
  t = c - a: Call Mi_es_Entero(t, nt, Tipot)
    If Tipot <= 0 Then
     Mi_funcion = 0
     Exit Sub
  End If
  t = c - b: Call Mi_es_Entero(t, nt, Tipot)
  If Tipot <= 0 Then
     Mi_funcion = 0
     Exit Sub
  End If
  Mi_funcion = F_Gamma(c) * F_Gamma(c - a - b) / F_Gamma(c - a) / F_Gamma(c - b)
  Exit Sub
End If

' Ecuación 15.1.21
t = 1 + a - b
Call Mi_es_Entero(t, nt, Tipot)
If Tipot <= 0 And Abs(x + 1) < Eps Then
     Mi_funcion = ChrW(8734)
     Exit Sub
End If

' Ecuación 15.1.21 b
If Tipot > 0 And Abs(x + 1) < Eps And Abs(c - (a - b + 1)) < Eps Then
  Calculado = "SI"
  t = 1 + 0.5 * a - b: Call Mi_es_Entero(t, nt, Tipot)
  If Tipot <= 0 Then
     Mi_funcion = 0
     Exit Sub
  End If
  t = 0.5 + 0.5 * a: Call Mi_es_Entero(t, nt, Tipot)
  If Tipot <= 0 Then
     Mi_funcion = 0
     Exit Sub
  End If
  Mi_funcion = 2 ^ (-a) * Sqr(Pi) * F_Gamma(1 + a - b)
  Mi_funcion = Mi_funcion / F_Gamma(1 + 0.5 * a - b) / F_Gamma(0.5 + 0.5 * a)
  Exit Sub
End If

' Ecuación 15.1.22a
t = 2 + a - b
Call Mi_es_Entero(t, nt, Tipot)

If Tipot <= 0 And (Abs(c - a + b - 2) < Eps) And (Abs(x + 1)) < Eps Then
  Calculado = "SI"
  Mi_funcion = ChrW(8734)
  Exit Sub
End If

' Ecuación 15.1.22b

If Tipot > 0 And Abs(x + 1) < Eps And Abs(c - (a - b + 2)) < Eps And (Abs(b - 1) > Eps) Then
  Calculado = "SI"
  f1 = 1: f2 = 1
  t = 0.5 * a: Call Mi_es_Entero(t, nt, Tipot)
  If Tipot <= 0 Then f1 = 0
  u = 1.5 + 0.5 * a - b: Call Mi_es_Entero(u, nu, Tipou)
  If Tipou <= 0 Then f2 = 0
  If Abs(f1 * f2) < Eps Then
     Sumando1 = 0
  Else
     Sumando1 = 1 / F_Gamma(0.5 * a) / F_Gamma(1.5 + 0.5 * a - b)
  End If
  
  f1 = 1: f2 = 1
  t = 0.5 + 0.5 * a: Call Mi_es_Entero(t, nt, Tipot)
  If Tipot <= 0 Then f1 = 0
  u = 1 + 0.5 * a - b: Call Mi_es_Entero(u, nu, Tipou)
  If Tipou <= 0 Then f2 = 0
  If Abs(f1 * f2) < Eps Then
     Sumando2 = 0
  Else
     Sumando2 = 1 / F_Gamma(0.5 + 0.5 * a) / F_Gamma(1 + 0.5 * a - b)
  End If
  
  Mi_funcion = 2 ^ (-a) * Sqr(Pi) / (b - 1) * F_Gamma(2 + a - b)
  Mi_funcion = Mi_funcion * (Sumando1 - Sumando2)
  Exit Sub
End If

' Ecuación 15.1.23
If (Abs(c - b - 1) < Eps) And (Abs(a - 1) < Eps) And (Abs(x + 1) < Eps) Then
  Calculado = "SI"
  Mi_funcion = 0.5 * b * (F_DigammaZ(0.5 + 0.5 * b) - F_DigammaZ(0.5 * b))
  Exit Sub
End If

' Ecuación 15.1.24a
t = 0.5 * (1 + a + b)
Call Mi_es_Entero(t, nt, Tipot)

If (Tipot <= 0) And (Abs(c - (0.5 * (a + b + 1))) < Eps) And (Abs(x - 0.5) < Eps) Then
  Calculado = "SI"
  Mi_funcion = ChrW(8734)
  Exit Sub
End If

' Ecuación 15.1.24b
If (Tipot > 0) And (Abs(c - (0.5 * (a + b + 1))) < Eps) And (Abs(x - 0.5) < Eps) Then
  Calculado = "SI"
  t = 0.5 + 0.5 * a: Call Mi_es_Entero(t, nt, Tipot)
  If Tipot <= 0 Then
     Mi_funcion = 0
     Exit Sub
  End If
  t = 0.5 + 0.5 * b: Call Mi_es_Entero(t, nt, Tipot)
  If Tipot <= 0 Then
     Mi_funcion = 0
     Exit Sub
  End If
  Mi_funcion = Sqr(Pi) * F_Gamma(0.5 * (1 + a + b))
  Mi_funcion = Mi_funcion / F_Gamma(0.5 + 0.5 * a) / F_Gamma(0.5 + 0.5 * b)
  Exit Sub
End If

' La ecuación 15.1.25 resuelve el mismo caso de la anterior con otra formulación

' Ecuación 15.1.26a
t = c
Call Mi_es_Entero(t, nt, Tipot)

If (Tipot <= 0) And (Abs(a + b) < Eps) And (Abs(x - 0.5) < Eps) Then
  Calculado = "SI"
  Mi_funcion = ChrW(8734)
  Exit Sub
End If

' Ecuación 15.1.26b
If (Tipot > 0) And (Abs(a + b - 1) < Eps) And (Abs(x - 0.5) < Eps) Then
  Calculado = "SI"
  t = 0.5 * a + 0.5 * c: Call Mi_es_Entero(t, nt, Tipot)
  If Tipot <= 0 Then
     Mi_funcion = 0
     Exit Sub
  End If
  t = 0.5 + 0.5 * c - 0.5 * a: Call Mi_es_Entero(t, nt, Tipot)
  If Tipot <= 0 Then
     Mi_funcion = 0
     Exit Sub
  End If
  Mi_funcion = 2 ^ (1 - c) * Sqr(Pi) * F_Gamma(c)
  Mi_funcion = Mi_funcion / F_Gamma(0.5 * a + 0.5 * c) / F_Gamma(0.5 + 0.5 * c - 0.5 * a)
  Exit Sub
End If

' Ecuación 15.1.27
If (Abs(a - 1) < Eps) And (Abs(b - 1) < Eps) And (Abs(x - 0.5) < Eps) Then
  Calculado = "SI"
  t = c - 1
  Mi_funcion = t * (F_DigammaZ(0.5 + 0.5 * t) - F_DigammaZ(0.5 * t))
  Exit Sub
End If

' Ecuación 15.1.28
If (Abs(a - b) < Eps) And (Abs(c - a - 1) < Eps) And (Abs(x - 0.5) < Eps) Then
  Calculado = "SI"
  Mi_funcion = 2 ^ (a - 1) * a * (F_DigammaZ(0.5 + 0.5 * a) - F_DigammaZ(0.5 * a))
  Exit Sub
End If

' Ecuación 15.1.29a
t = 1.5 - 2 * a
Call Mi_es_Entero(t, nt, Tipot)

If (Tipot <= 0) And (Abs(b - a - 0.5) < Eps) And (Abs(c + 2 * a - 1.5) < Eps) And (Abs(x + 1 / 3) < Eps) Then
  Calculado = "SI"
  Mi_funcion = ChrW(8734)
  Exit Sub
End If

' Ecuación 15.1.29b
If (Tipot > 0) And (Abs(b - a - 0.5) < Eps) And (Abs(c + 2 * a - 1.5) < Eps) And (Abs(x + 1 / 3) < Eps) Then
  Calculado = "SI"
  t = 4 / 3 - 2 * a: Call Mi_es_Entero(t, nt, Tipot)
  If Tipot <= 0 Then
     Mi_funcion = 0
     Exit Sub
  End If
  Mi_funcion = (8 / 9) ^ (-2 * a) * F_Gamma(4 / 3) * F_Gamma(1.5 - 2 * a)
  Mi_funcion = Mi_funcion / F_Gamma(1.5) / F_Gamma(4 / 3 - 2 * a)
  Exit Sub
End If

End Sub


' FUNCIÓN HIPERGEOMÉTRICA DE GAUSS 3F2

Public Function F_HG_3F2(aa As Double, bb As Double, cc As Double, _
                         dd As Double, ee As Double, _
                         xx As Double, Optional NSumandos As Integer = 32000, _
                         Optional Eps As Double = 0.0000001) As Variant
' Calcula la función Hipergeométrica  3F2(a,b,c;d,e|x)
' Llama a F_HG_Gauss
' Llama a la función F_pFq

Dim va(1 To 3) As Double, vb(1 To 3) As Double

' Valores nulos
If Abs(aa * bb * cc) < Eps Then
   F_HG_3F2 = 1
   Exit Function
End If

If Abs(dd * ee) < Eps Then
   F_HG_3F2 = ChrW(8734) & " valor nulo en denominador"
   Exit Function
End If

' Reducimos un grado cuando dos factores son iguales pasando a la
' función Hipergeométrica de Gauss
If Abs(aa - dd) < Eps Then
   F_HG_3F2 = F_HG_Gauss(bb, cc, ee, xx, NSumandos)
   Exit Function
ElseIf Abs(aa - ee) < Eps Then
   F_HG_3F2 = F_HG_Gauss(bb, cc, dd, xx, NSumandos)
   Exit Function
ElseIf Abs(aa - dd) < Eps Then
   F_HG_3F2 = F_HG_Gauss(bb, cc, ee, xx, NSumandos)
   Exit Function
ElseIf Abs(bb - ee) < Eps Then
   F_HG_3F2 = F_HG_Gauss(aa, cc, dd, xx, NSumandos)
   Exit Function
ElseIf Abs(bb - dd) < Eps Then
   F_HG_3F2 = F_HG_Gauss(aa, cc, ee, xx, NSumandos)
   Exit Function
ElseIf Abs(cc - ee) < Eps Then
   F_HG_3F2 = F_HG_Gauss(aa, bb, dd, xx, NSumandos)
   Exit Function
ElseIf Abs(cc - dd) < Eps Then
   F_HG_3F2 = F_HG_Gauss(aa, bb, ee, xx, NSumandos)
   Exit Function
End If

va(1) = aa: va(2) = bb: va(3) = cc
vb(1) = dd: vb(2) = ee
F_HG_3F2 = F_pFq(va, vb, xx, NSumandos)

End Function


