
' FUNCIÓN C_Pi

Public Function C_Pi(N As Long) As Double
' Obtiene el valor del número Pi mediante el método de Montecarlo
' N es el número de puntos a "tirar"
Dim x As Double, y As Double, r As Double
Dim i As Long
Dim Cuenta As Long

Randomize
 
For i = 1 To N
    x = Rnd
    y = Rnd
    r = x * x + y * y
    If r <= 1 Then
       Cuenta = Cuenta + 1
    End If
Next

C_Pi = 4 * Cuenta / N

End Function


' RUTINA S_Pi

Public Sub S_Pi()
' Obtiene el valor del número Pi mediante el método de Montecarlo y pinta lo que hace
' N es el número de puntos a "tirar"

Dim x As Double, y As Double, r As Double
Dim i As Long, j1 As Long, j2 As Long
Dim Pi As Double
Dim N As Long
Dim Mi_Chart As Chart    ' Objeto Chart donde se representan los gráficos

N = Range("N_Pi").Value

' Borramos los datos previos si existen
Columns("H:K").ClearContents
Columns("N:P").ClearContents
Columns("R:S").ClearContents

' Cabeceras de los datos
Cells(1, 8).Value = "x-Si"
Cells(1, 9).Value = "y-Si"
Cells(1, 10).Value = "x-No"
Cells(1, 11).Value = "y-No"

Cells(1, 14).Value = "Phi"
Cells(1, 15).Value = "x"
Cells(1, 16).Value = "y"

Cells(1, 18).Value = "x"
Cells(1, 19).Value = "y"

j1 = 0
j2 = 0

' Coordenadas de los puntos aleatorios
Randomize
For i = 1 To N
    x = Rnd
    y = Rnd
    r = x * x + y * y
    If r <= 1 Then
       j1 = j1 + 1
       Cells(j1 + 1, 8).Value = x
       Cells(j1 + 1, 9).Value = y
    Else
       j2 = j2 + 1
       Cells(j2 + 1, 10).Value = x
       Cells(j2 + 1, 11).Value = y
    End If
Next i

Pi = 4 * j1 / N
Cells(5, 6).Value = Pi

' Coordenadas de la circunferencia
Cells(2, 14).Value = 0
Cells(2, 15).Value = Cos(Cells(2, 14).Value)
Cells(2, 16).Value = Sin(Cells(2, 14).Value)
For i = 1 To 15
  Cells(i + 2, 14).Value = Cells(i + 1, 14).Value + WorksheetFunction.Pi() / 30
  Cells(i + 2, 15).Value = Cos(Cells(i + 2, 14).Value)
  Cells(i + 2, 16).Value = Sin(Cells(i + 2, 14).Value)
Next i

' Coordenadas del cuadrado
Cells(2, 18).Value = 0
Cells(2, 19).Value = 1
Cells(3, 18).Value = 1
Cells(3, 19).Value = 1
Cells(4, 18).Value = 1
Cells(4, 19).Value = 0

' Borramos el gráfico existente previamente
For i = 1 To Worksheets("Uniforme E1").ChartObjects.Count
  If Worksheets("Uniforme E1").ChartObjects(i).Name = "Grafico_1" Then
    Worksheets("Uniforme E1").ChartObjects(i).Delete
  End If
Next i

' Creamos el gráfico y le asignamos nombre y dimensiones
ActiveSheet.Shapes.AddChart2(240, xlXYScatter, 700, 280, _
                             350, 350).Name = "Grafico_1"
Set Mi_Chart = ActiveSheet.Shapes("Grafico_1").Chart

' Primera serie de datos del gráfico: puntos dentro del círculo
' Asignamos los datos y configuramos (nombre, marcadores, tipo de gráfico)
Mi_Chart.SetSourceData Source:=Range("'Uniforme E1'!$H$2:$I$" & j1 + 1)

With Mi_Chart.FullSeriesCollection(1)
    .Name = "Dentro Círculo"
    .MarkerStyle = xlMarkerStyleCircle
    .MarkerSize = 6
    .MarkerBackgroundColorIndex = xlColorIndexNone
    .MarkerForegroundColor = RGB(91, 155, 213)
    .Format.Line.Visible = msoFalse
    .Format.Line.Weight = 0.5
    .ChartType = xlXYScatter
    .AxisGroup = 1
End With

' Segunda serie de datos: puntos fuera del círculo
Mi_Chart.SeriesCollection.NewSeries
' Asignamos los datos y configuramos (nombre, marcadores, tipo de gráfico)
With Mi_Chart.FullSeriesCollection(2)
    .Name = "Fuera Círculo"
    .XValues = Range("'Uniforme E1'!$J$2:$J$" & j2 + 1)
    .Values = Range("'Uniforme E1'!$K$2:$K$" & j2 + 1)
    .MarkerStyle = xlMarkerStyleCircle
    .MarkerSize = 6
    .MarkerBackgroundColorIndex = xlColorIndexNone
    .MarkerForegroundColor = RGB(254, 131, 92)
    .Format.Line.Visible = msoFalse
    .Format.Line.Weight = 0.5
    .ChartType = xlXYScatter
    .AxisGroup = 1
End With

' Tercera serie de datos: circunferencia
Mi_Chart.SeriesCollection.NewSeries
' Asignamos los datos y configuramos (nombre, marcadores, tipo de gráfico)
With Mi_Chart.FullSeriesCollection(3)
    .Name = "Circunferencia"
    .XValues = Range("'Uniforme E1'!$O$2:$O$17")
    .Values = Range("'Uniforme E1'!$P2:$P$17")
    .ChartType = xlXYScatterSmoothNoMarkers
    .AxisGroup = 1
    .Format.Line.ForeColor.RGB = RGB(84, 130, 53)
    .Format.Line.Weight = 2
End With

' Cuarta serie de datos: cuadrado
Mi_Chart.SeriesCollection.NewSeries
' Asignamos los datos y configuramos (nombre, marcadores, tipo de gráfico)
With Mi_Chart.FullSeriesCollection(4)
    .Name = "Cuadrado"
    .XValues = Range("'Uniforme E1'!$R$2:$R$4")
    .Values = Range("'Uniforme E1'!$S2:$S$4")
    .ChartType = xlXYScatterLinesNoMarkers
    .AxisGroup = 1
    .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
    .Format.Line.Weight = 1.5
End With

' Valores extremos y escala del eje x
With Mi_Chart.Axes(xlCategory)
    .MinimumScale = 0
    .MaximumScale = 1
    .MajorUnit = 0.2
End With

' Valores extremos y escala del eje y
With Mi_Chart.Axes(xlValue)
    .MinimumScale = 0
    .MaximumScale = 1
    .MajorUnit = 0.2
End With

' Título del gráfico
Mi_Chart.HasTitle = True
Mi_Chart.ChartTitle.Text = "Montecarlo: número " & ChrW(960)
Mi_Chart.ChartTitle.Font.Color = RGB(0, 0, 0)

' Leyenda no visible
Mi_Chart.SetElement (msoElementLegendNone)

' Configuramos el eje x (color, ancho, etiquetas y marcas de graduación)
With Mi_Chart.Axes(xlCategory).Format.Line
    .Visible = msoTrue
    .ForeColor.RGB = RGB(0, 0, 0)
    .Weight = 1.25
End With
Mi_Chart.Axes(xlCategory).MajorTickMark = xlOutside

Mi_Chart.Axes(xlCategory).TickLabels.Font.Size = 11
Mi_Chart.Axes(xlCategory).TickLabels.Font.Color = RGB(0, 0, 0)

' Configuramos el eje y (color, ancho, etiquetas y marcas de graduación)
With Mi_Chart.Axes(xlValue).Format.Line
    .Visible = msoTrue
    .ForeColor.RGB = RGB(0, 0, 0)
    .Weight = 1.25
End With
Mi_Chart.Axes(xlValue).MajorTickMark = xlOutside

Mi_Chart.Axes(xlValue).TickLabels.Font.Size = 11
Mi_Chart.Axes(xlValue).TickLabels.Font.Color = RGB(0, 0, 0)

' Configuramos el borde del área del gráfico
With Mi_Chart.ChartArea.Border
    .LineStyle = xlContinuous
    .Color = RGB(100, 100, 100)
    .Weight = xlThin
End With

Set Mi_Chart = Nothing

Range("A1").Select

End Sub


' FUNCIÓN Integral

Public Function Integral(Fun As String, a As Double, b As Double, N As Long) As Double

' Montecarlo (Sólo para funciones NO NEGATIVAS)
' N  = Número de puntos en el método de Montecarlo (<=2.147.483.648)
' Llama a la función Evalua

Dim Puntos As Long
Dim P As Double, q As Double, h As Double
Dim Psi As Double, MaxFun As Double
Dim xx() As Double, yy() As Double, zz() As Double
Dim i As Long
Dim Area As Double

ReDim xx(1 To N), yy(1 To N), zz(1 To N)

' Pasamos a minúsculas el nombre de la función
Fun = LCase(Fun)


' Comienza el algoritmo

P = (b + a) / 2    ' Punto medio del intervalo
q = (b - a) / 2    ' Semilongitud del intervalo

Puntos = N
Randomize          ' Creamos la semills

MaxFun = -999999999
For i = 1 To Puntos
    Psi = a + 2 * q * Rnd ' Generamos valores aleatorios
    xx(i) = Psi
    zz(i) = Evalua(Fun, Psi)
    If zz(i) > MaxFun Then
       MaxFun = zz(i)
    End If
Next

Randomize
Integral = 0
Area = 2 * q * MaxFun
For i = 1 To Puntos
    Psi = 0 + MaxFun * Rnd
    yy(i) = Psi
    If yy(i) <= zz(i) Then
       Integral = Integral + 1
    End If
Next
Integral = Integral / Puntos * Area
End Function


' FUNCIÓN Evalua

Public Function Evalua(Funcion As String, a As Double) As Double

' Pasamos a minúsculas y arreglamos algún problema con nombres de funciones
Funcion = LCase(Funcion)
Funcion = Replace(Funcion, " ", "")
Funcion = Replace(Funcion, "sqrt", "sqr")
Funcion = Replace(Funcion, "sqr", "sqrt")
Funcion = Replace(Funcion, "atn", "atan")
Funcion = Replace(Funcion, "seno", "sin")
Funcion = Replace(Funcion, "sen", "sin")
Funcion = Replace(Funcion, "sgn", "sign")
Funcion = Replace(Funcion, "rnd", "rand")
Funcion = Replace(Funcion, "exp", "EXP")
Funcion = Replace(Funcion, "pi", Pi)

' primero substituimos la variable por el valor
' luego cambiamos coma decimal por punto decimal
Evalua = Evaluate(Replace(Replace(Funcion, "x", a), ",", "."))

End Function


' FUNCIÓN U_e

Public Function U_e(N As Long) As Double
' Realiza una estimación del número e
Dim i As Long, j() As Long, Min As Long
Dim Suma() As Double, eN As Double

ReDim Suma(1 To N), j(1 To N)

eN = 0
Min = 9999999
For i = 1 To N
  Suma(i) = 0: j(i) = 0
  Randomize
  Do While Suma(i) <= 1
     Suma(i) = Suma(i) + Rnd
     j(i) = j(i) + 1
  Loop
  If j(i) < Min Then Min = j(i)
  eN = eN + j(i)
Next

U_e = eN / N

End Function


' FUNCIÓN Esperanza_Moneda

Public Function Esperanza_Moneda(N As Long, p As Double, d As Double, Optional Metodo As Integer = 1) As Variant
' Esta función calcula el beneficio esperable tirando una moneda "cargada" n veces.
' La probabilidad de que salga, por ejemplo, cara vale 0<=p<=1 y apostamos la cantidad de d dólares
' siempre a que salga precisamente cara.
' Si Metodo=0 se emplea la función de probabilidad de la distribución Binomial y
' Si Metodo=1 se emplea la aproximación a la Normal
' Llama a la función p_Binomial
' Llama a la función F_Normal_01_H

Dim q As Double, Mu As Double, Sigma As Double
Dim k As Long, Pr() As Double
Dim Suma As Double, Sumando As Double
Dim x1 As Double, x2 As Double

ReDim Pr(0 To N)

' Primero miramos si p tiene un valor razonable
If p < 0 Or p > 1 Then
   ' Valor incorrecto de p
   Esperanza_Moneda = "p Fuera de rango"
   Exit Function
End If

' Ahora comprobamos que N sea mayor que 0
If N <= 0 Then
   ' Valor incorrecto de N
   Esperanza_Moneda = "N debe ser >0"
   Exit Function
End If

q = 1 - p
Mu = N * p
Sigma = Sqr(N * p * q)

Suma = 0: Sumando = 0
For k = 0 To N
  If Metodo = 0 Then
     ' Empleamos la f.p. de la distribución Binomial
     Pr(k) = p_Binomial(k, N, p)
  Else
     ' Empleamos la aproximacióna a la Normal
     ' Utilizamos la formulación aproximada de Hastings
     x1 = (k - Mu - 0.5) / Sigma
     x2 = (k - Mu + 0.5) / Sigma
     Pr(k) = F_Normal_01_H(x2) - F_Normal_01_H(x1)
  End If
  Sumando = Pr(k) * (2 * k - N)
  Suma = Suma + Sumando
Next k
Esperanza_Moneda = d * Suma

End Function


' FUNCIÓN Saca_Alfa_Pareto

Public Function Saca_Alfa_Pareto(p As Double, Fracc As Double, x0 As Double, xM As Double) As Variant
                
' Esta función resuelve por bisección la ecuación Func(Alfa)=f-f0[1-p*Beta^(-1/Alfa)]=0
' Siendo Beta = 1-f0^(-1/Alfa)

' p      Es la probabilidad que queremos "resolver" (por ejemplo 0,8=80%)
' Fracc  Es el porcentaje de población que queremos que acapare la probabilidad p (por ejemplo 0,20=20%)
' x0, xM Parámetros de la distribución de Pareto truncada

Dim f0 As Double, f As Double
Dim Alfa1 As Double, Alfa2 As Double
Dim Func1 As Double, Func2 As Double, Func As Double
Dim Alfa As Double
Dim Eps As Double

Eps = 0.00001
f0 = x0 / xM
f = Fracc

Alfa1 = 0.001
Alfa2 = 1000

Func1 = f - F_ParetoT_Inv(p, Alfa1, x0, xM) / xM
Func2 = f - F_ParetoT_Inv(p, Alfa2, x0, xM) / xM

If Func1 * Func2 > 0 Then
   Saca_Alfa_Pareto = "Rango Mal"
   Exit Function
End If

Func = 1
Do While Abs(Func) > Eps
   Alfa = (Alfa1 + Alfa2) / 2
   Func = f - F_ParetoT_Inv(p, Alfa, x0, xM) / xM
   If Func * Func1 > 0 Then
      Alfa1 = Alfa
      Func1 = Func
   Else
      Alfa2 = Alfa
   End If
Loop

Saca_Alfa_Pareto = Alfa
End Function


' RUTINA Cuadro

Sub Cuadro()
' Llama a la función lnMV
Dim Col1 As Integer, Fil1 As Integer, Col As Integer, Fil As Integer
Dim i As Integer, j As Integer
Dim Alfa As Double, Beta As Double
Dim dAlfa As Double, dBeta As Double
Dim Suma As Double, LSuma As Double, n As Long
Dim AlfaV(1 To 11) As Double, BetaV(1 To 11) As Double, LogMV(1 To 11, 1 To 11) As Double

Col1 = 16: Fil1 = 6

Alfa = Range("E12")
Beta = Range("E13")
Suma = Range("E8")
LSuma = Range("E9")
n = Range("E10")
 
dAlfa = Alfa / 10
dBeta = Beta / 10

' Rellenamos los vectores de Alfa y Beta
For i = 1 To 11
    AlfaV(i) = Alfa + dAlfa * (i - 6)
    BetaV(i) = Beta + dBeta * (i - 6)
    Cells(Fil1, Col1 + i) = BetaV(i)
    Cells(Fil1 + i, Col1) = AlfaV(i)
Next

' Calculamos la función
For i = 1 To 11
For j = 1 To 11
    Col = Col1 + i
    Fil = Fil1 + j
    LogMV(i, j) = lnMV(AlfaV(i), BetaV(j), Suma, LSuma, n)
    Cells(Fil, Col) = LogMV(i, j)
Next
Next

End Sub


' FUNCIÓN lnMV

Public Function lnMV(Alfa As Double, Beta As Double, Suma As Double, LSuma As Double, n As Long) As Double
' Llama a la función F_Gamma
Dim Ga As Double

Ga = Log(F_Gamma(Alfa))
lnMV = -n * Ga - n * Alfa * Log(Beta) + (Alfa - 1) * LSuma - 1 / Beta * Suma

End Function


' RUTINA Opt_Gradiente

Sub Opt_Gradiente()
' Llama a la función Fu
' Llama a la función Fx
' Llama a la función Fy
Dim Eps As Double                ' Error permitido
Dim x As Double, y As Double     ' Coordenadas del punto
Dim x0 As Double, y0 As Double   ' Coordenadas del punto en el paso anterior
Dim dx As Double, dy As Double   ' Incrementos en las variables
Dim dgx As Double, dgy As Double ' Incrementos en el gradiente
Dim NGrad As Double              ' Norma de la diferencia del gradiente
Dim MGrad As Double              ' Módulo del gradiente
Dim gx As Double, gy As Double   ' Componentes del gradiente (derivadas parciales)
Dim gx0 As Double, gy0 As Double ' Componentes del gradiente en el paso anterior
Dim F0 As Double, F As Double    ' Valores de la función
Dim Alfa As Double               ' "Porción" del gradiente que avanzamos
Dim NMax As Integer              ' Número máximo de iteraciones permitidas
Dim i As Integer, j As Integer   ' Contadores
Dim k As Integer
Dim Texto As String              ' Texto cuando para convergencia

' Limpiamos el campo
Columns("F:AB").Select
Selection.ClearContents
Selection.Font.ColorIndex = xlAutomatic
Selection.HorizontalAlignment = xlCenter

' Cabeceras en negrita
Range("F1:AB1").Select
Selection.Font.Bold = True
j = 2: k = 7

Cells(1, k + 1) = "Paso"
Cells(1, k + 2) = "x"
Cells(1, k + 3) = "y"
Cells(1, k + 4) = "F(x,y)"
Cells(1, k + 5) = "F'x"
Cells(1, k + 6) = "F'y"
Cells(1, k + 7) = "Alfa"
Cells(1, k + 8) = "|Avance|"
Cells(1, k + 9) = "D(F)"
Cells(1, k + 10) = "|g|"
Cells(1, k + 11) = "dx.dg"
Cells(1, k + 12) = "|dg|^2"

' Paso 0: Leemos los valores iniciales y parámetros de convergencia

x0 = Range("x0")
y0 = Range("y0")
Eps = Range("Eps")
NMax = Range("NMax")
hx = Range("D_x")
hy = Range("D_y")


' Paso inicial
Alfa = 0.001
F0 = Fu(x0, y0)
gx0 = Fx(F0, x0, y0)
gy0 = Fy(F0, x0, y0)
gx = gx0
gy = gy0

Cells(j, k + 1) = 0
Cells(j, k + 2) = x0
Cells(j, k + 3) = y0
Cells(j, k + 4) = F0
Cells(j, k + 5) = gx0
Cells(j, k + 6) = gy0
Cells(j, k + 7) = Alfa
Cells(j, k + 9) = ""

Texto = "No"
For i = 1 To NMax
    j = i + 2
    x = x0 - Alfa * gx
    y = y0 - Alfa * gy
    F = Fu(x, y)
    Cells(j, k + 1) = i
    Cells(j, k + 2) = x
    Cells(j, k + 3) = y
    Cells(j, k + 4) = F
    Cells(j, k + 9) = F - F0
    
    If Abs(F - F0) <= Eps And MGrad < Eps Then
       ' Hemos alcanzado el mínimo y salimos
       Cells(j, k + 8) = Sqr((x - x0) ^ 2 + (y - y0) ^ 2)
       Cells(j, k + 5) = "CONVERGENCIA"
       Cells(j, k + 5).Select
       Selection.Font.Color = -16776961
       Selection.HorizontalAlignment = xlLeft
       Texto = "Si"
       Exit For
    End If
    ' Hay que seguir
    
    ' 1. Calculamos valores nuevos
    ' Gradiente
    gx = Fx(F, x, y)
    gy = Fy(F, x, y)
    MGrad = Sqr(gx * gx + gy * gy)
    
    ' Incremento en movimiento y gradiente
    dx = x - x0: dy = y - y0
    dgx = gx - gx0: dgy = gy - gy0: NGrad = dgx * dgx + dgy * dgy

    ' Alfa
    Alfa = (dx * dgx + dy * dgy) / NGrad
    
    ' 2. Guardamos valores de la iteración
    x0 = x: y0 = y
    gx0 = gx: gy0 = gy
    F0 = F
    Cells(j, k + 5) = gx0
    Cells(j, k + 6) = gy0
    Cells(j, k + 7) = Alfa
    Cells(j, k + 8) = Alfa * Sqr(NGrad)
    Cells(j, k + 10) = MGrad
    Cells(j, k + 11) = dx * dgx + dy * dgy
    Cells(j, k + 12) = NGrad
Next

' Preguntamos si ha convergido
If Texto = "No" Then
   Texto = "Después de " & NMax & " Iteraciones"
   i = MsgBox(Texto, vbOK, "NO HAY CONVERGENCIA")
End If

End Sub


' FUNCIÓN Fu

Public Function Fu(x As Double, y As Double) As Double
    ' Función a optimizar
    ' Llama a la función lnMV2

    Fu = lnMV2(x, y)
End Function


' FUNCIÓN Fx

Public Function Fx(F As Double, x As Double, y As Double) As Double
    ' Derivada parcial x
    ' Se introduce el valor F para no calcularlo de nuevo
    ' Llama a la función Fu
    Dim F1 As Double, dF As Double

    F1 = Fu(x + hx, y)
    dF = F1 - F
    Fx = dF / hx
End Function


' FUNCIÓN Fy

Public Function Fy(F As Double, x As Double, y As Double) As Double
    ' Derivada parcial x
    ' Se introduce el valor F para no calcularlo de nuevo
    ' Llama a la función Fu
    Dim F1 As Double, dF As Double

    F1 = Fu(x, y + hy)
    dF = F1 - F
    Fy = dF / hy
End Function


' FUNCIÓN lnMV2

Public Function lnMV2(Alfa As Double, Beta As Double) As Double
' Llama a la función F_Gamma
Dim Ga As Double

Ga = Log(F_Gamma(Alfa))
lnMV2 = -100 * Ga - 100 * Alfa * Log(Beta) - (Alfa - 1) * 93.7553 - 1 / Beta * 62.251

End Function


' RUTINA Ajuste_Lineal

Sub Ajuste_Lineal()
Dim x(1 To 100000) As Double, y(1 To 100000) As Double         ' Datos de entrada
Dim yE(1 To 10000) As Double                                   ' Ordenada estimada con el ajuste
Dim n As Long                                                  ' Número de puntos
Dim xM As Double, yM As Double                                 ' Valores medios
Dim Cov As Double, S2x As Double, S2y As Double                ' Cuasivarianzas
Dim Sx As Double, Sy As Double                                 ' Desviaciones estándar
Dim a As Double, b As Double                                   ' Coeficientes de la recta y = bx + a
Dim aMin As Double, aMax As Double
Dim bMin As Double, bMax As Double
Dim RSS As Double, TSS As Double                               ' Errores cuadráticos
Dim sR As Double, Sa As Double, Sb As Double, SyN As Double    ' Desviaciones estándar para intervalos de confianza
Dim r As Double, R2 As Double                                  ' Coeficientes de correlación y determinación
Dim i As Long                                                  ' Índice para bucles
Dim alfa As Double                                             ' Nivel de confianza
Dim t As Double                                                ' F^(-1)(alfa, n-2)
Dim SemiAmplitud As Double                                     ' Semiamplitud de intervalo de confianza
Dim xmin As Double, xmax As Double                             ' Extremos de abscisas

' Leemos los datos
xmin = 1E+99: xmax = -1E+99
For i = 3 To 100000
    If Cells(i, 1) = "" Then Exit For
    x(i - 2) = Cells(i, 1)
    y(i - 2) = Cells(i, 2)
    n = i - 2
    xM = xM + x(i - 2)
    yM = yM + y(i - 2)
    If x(i - 2) < xmin Then xmin = x(i - 2)
    If x(i - 2) > xmax Then xmax = x(i - 2)
Next

' Leemos el nivel de confianza Alfa
alfa = Cells(1, 2)

If alfa <= 0 Or alfa > 1 Then
   MsgBox "El nivel de confianza debe ser >0 y <1", vbOKOnly, "Nivel de confianza incorrecto"      Exit Sub
End If

' Cálculo de estadísticos
xM = xM / n: yM = yM / n                                      ' Valores medios
Cov = 0: S2x = 0: S2y = 0: RSS = 0                            ' Cuasivarianzas
For i = 1 To n
    S2x = S2x + (x(i) - xM) ^ 2
    S2y = S2y + (y(i) - yM) ^ 2
    Cov = Cov + (x(i) - xM) * (y(i) - yM)
Next
TSS = S2y
S2x = S2x / (n - 1): S2y = S2y / (n - 1): Cov = Cov / (n - 1)
Sx = Sqr(S2x): Sy = Sqr(S2y)                                  ' Desviaciones estándar

' Escribimos valores
Cells(3, 8) = xM: Cells(4, 8) = yM
Cells(5, 8) = Sx: Cells(6, 8) = Sy: Cells(7, 8) = Cov

' Coeficientes de la recta
b = Cov / S2x
a = yM - b * xM
Cells(8, 8) = a: Cells(9, 8) = b

' Estimaciones de los valores introducidos
RSS = 0
For i = 1 To n
    yE(i) = b * x(i) + a
    RSS = RSS + (y(i) - yE(i)) ^ 2
Next
sR = Sqr(RSS / (n - 2))
Sb = sR / Sqr(n - 1) / Sx
Sa = sR * Sqr(1 / n + xM ^ 2 / ((n - 1) * S2x))

' Coeficientes de correlación y determinación
r = Sx / Sy * b
R2 = 1 - RSS / TSS
Cells(10, 8) = TSS: Cells(11, 8) = RSS
Cells(20, 8) = r: Cells(21, 8) = R2

' t de Student
t = F_t_Student_Inv(1 - alfa / 2, n - 2)
Cells(1, 5) = t

' Escribimos valores estimados

' Limpiamos primero los campos
Range("C3:F10000").ClearContents
    
' Intervalos de confianza
For i = 1 To n
    yE(i) = b * x(i) + a
    Cells(i + 2, 3) = yE(i)
    SyN = sR * Sqr(1 / n + (x(i) - xM) ^ 2 / ((n - 1) * S2x))
    SemiAmplitud = t * SyN
    Cells(i + 2, 4) = yE(i) - SemiAmplitud
    Cells(i + 2, 5) = yE(i) + SemiAmplitud
Next

' Pendiente
Cells(13, 8) = sR
Cells(14, 8) = Sb: Cells(15, 8) = Sa

SemiAmplitud = t * Sb
bMin = b - SemiAmplitud
bMax = b + SemiAmplitud
Cells(16, 8) = bMin: Cells(17, 8) = bMax

' Ordenada en el origen
SemiAmplitud = t * Sa
aMin = a - SemiAmplitud
aMax = a + SemiAmplitud
Cells(18, 8) = aMin: Cells(19, 8) = aMax

' Puntos extremos de las rectas
Cells(25, 8) = xmin           ' Abscisas
Cells(26, 8) = xmax

Cells(25, 9) = b * xmin + a   ' Recta de regresión
Cells(26, 9) = b * xmax + a

Cells(25, 10) = bMin * xmin + aMax  ' Recta de regresión baja
Cells(26, 10) = bMin * xmax + aMax

Cells(25, 11) = bMax * xmin + aMin  ' Recta de regresión alta
Cells(26, 11) = bMax * xmax + aMin

End Sub


' FUNCIÓN Ajuste_Lineal_y

Public Function Ajuste_Lineal_y(XRango As Range, YRango As Range, xx As Double) As Variant
Dim x() As Double, y() As Double                               ' Datos de entrada
Dim yE As Double                                               ' Ordenada estimada con el ajuste
Dim YEMin As Double, YEMax As Double
Dim n As Long                                                  ' Número de puntos
Dim xM As Double, yM As Double                                 ' Valores medios
Dim Cov As Double, S2x As Double, S2y As Double                ' Cuasivarianzas
Dim Sx As Double, Sy As Double                                 ' Desviaciones estándar
Dim a As Double, b As Double                                   ' Coeficientes de la recta y = bx + a
Dim sR As Double, SyN As Double                                ' Desviaciones estándar para intervalos de confianza
Dim i As Long                                                  ' Índice para bucles
Dim xmin As Double, xmax As Double                             ' Extremos de abscisas

'
' ---> Lectura de los vectores x() e y(), introducidos como rangos.
'
n = XRango.Rows.Count
i = YRango.Rows.Count
If n <> i Then
   Ajuste_Lineal_y = "Los rangos deben tener la misma longitud"
   Exit Function
End If

ReDim x(1 To n)
ReDim y(1 To n)

xmin = 1E+99: xmax = -1E+99
For i = 1 To n
    x(i) = XRango.Cells(i, 1)
    y(i) = YRango.Cells(i, 1)
    xM = xM + x(i)
    yM = yM + y(i)
    If x(i) < xmin Then xmin = x(i)
    If x(i) > xmax Then xmax = x(i)
Next

' Cálculo de estadísticos
xM = xM / n: yM = yM / n                                      ' Valores medios
Cov = 0: S2x = 0: S2y = 0                                     ' Cuasivarianzas
For i = 1 To n
    S2x = S2x + (x(i) - xM) ^ 2
    S2y = S2y + (y(i) - yM) ^ 2
    Cov = Cov + (x(i) - xM) * (y(i) - yM)
Next
S2x = S2x / (n - 1): S2y = S2y / (n - 1): Cov = Cov / (n - 1)
Sx = Sqr(S2x): Sy = Sqr(S2y)                                  ' Desviaciones estándar

' Coeficientes de la recta
b = Cov / S2x
a = yM - b * xM
    
' Ajuste
yE = b * xx + a
Ajuste_Lineal_y = yE

End Function


' FUNCIÓN Ajuste_Lineal_yMin

Public Function Ajuste_Lineal_yMin(XRango As Range, YRango As Range, xx As Double, alfa As Double) As Variant
Dim x() As Double, y() As Double                               ' Datos de entrada
Dim yE As Double                                               ' Ordenada estimada con el ajuste
Dim YEMin As Double, YEMax As Double
Dim n As Long                                                  ' Número de puntos
Dim xM As Double, yM As Double                                 ' Valores medios
Dim Cov As Double, S2x As Double, S2y As Double                ' Cuasivarianzas
Dim Sx As Double, Sy As Double                                 ' Desviaciones estándar
Dim a As Double, b As Double                                   ' Coeficientes de la recta y = bx + a
Dim RSS As Double, TSS As Double                               ' Errores cuadráticos
Dim sR As Double, SyN As Double                                ' Desviaciones estándar para intervalos de confianza
Dim r As Double, R2 As Double                                  ' Coeficientes de correlación  y determinación
Dim i As Long                                                  ' Índice para bucles
Dim t As Double                                                ' F^(-1)(alfa, n-2)
Dim SemiAmplitud                                               ' Semiamplitud de intervalo de confianza
Dim xmin As Double, xmax As Double                             ' Extremos de abscisas

'
' ---> Lectura de los vectores x() e y(), introducidos como rangos.
'
n = XRango.Rows.Count
i = YRango.Rows.Count
If n <> i Then
   Ajuste_Lineal_yMin = "Los rangos deben tener la misma longitud"
   Exit Function
End If

ReDim x(1 To n)
ReDim y(1 To n)

xmin = 1E+99: xmax = -1E+99
For i = 1 To n
    x(i) = XRango.Cells(i, 1)
    y(i) = YRango.Cells(i, 1)
    xM = xM + x(i)
    yM = yM + y(i)
    If x(i) < xmin Then xmin = x(i)
    If x(i) > xmax Then xmax = x(i)
Next

If alfa <= 0 Or alfa > 1 Then
   Ajuste_Lineal_yMin = "Alfa Debe ser >0 y <1"
   Exit Function
End If

' Cálculo de estadísticos
xM = xM / n: yM = yM / n                                      ' Valores medios
Cov = 0: S2x = 0: S2y = 0: RSS = 0                            ' Cuasivarianzas
For i = 1 To n
    S2x = S2x + (x(i) - xM) ^ 2
    S2y = S2y + (y(i) - yM) ^ 2
    Cov = Cov + (x(i) - xM) * (y(i) - yM)
Next
TSS = S2y
S2x = S2x / (n - 1): S2y = S2y / (n - 1): Cov = Cov / (n - 1)
Sx = Sqr(S2x): Sy = Sqr(S2y)                                  ' Desviaciones estándar

' Coeficientes de la recta
b = Cov / S2x
a = yM - b * xM

' Estimaciones de los valores introducidos
RSS = 0
For i = 1 To n
    yE = b * x(i) + a
    RSS = RSS + (y(i) - yE) ^ 2
Next
sR = Sqr(RSS / (n - 2))

' Coeficientes de correlación y determinación
r = Sx / Sy * b
R2 = 1 - RSS / TSS

' t de Student
t = F_t_Student_Inv(1 - alfa / 2, n - 2)
    
' Intervalos de confianza
yE = b * xx + a
SyN = sR * Sqr(1 / n + (xx - xM) ^ 2 / ((n - 1) * S2x))
SemiAmplitud = t * SyN
YEMin = yE - SemiAmplitud
YEMax = yE + SemiAmplitud

Ajuste_Lineal_yMin = YEMin

End Function


' FUNCIÓN Ajuste_Lineal_yMax

Public Function Ajuste_Lineal_yMax(XRango As Range, YRango As Range, xx As Double, alfa As Double) As Variant
Dim x() As Double, y() As Double                               ' Datos de entrada
Dim yE As Double                                               ' Ordenada estimada con el ajuste
Dim YEMin As Double, YEMax As Double
Dim n As Long                                                  ' Número de puntos
Dim xM As Double, yM As Double                                 ' Valores medios
Dim Cov As Double, S2x As Double, S2y As Double                ' Cuasivarianzas
Dim Sx As Double, Sy As Double                                 ' Desviaciones estándar
Dim a As Double, b As Double                                   ' Coeficientes de la recta y = bx + a
Dim RSS As Double, TSS As Double                               ' Errores cuadráticos
Dim sR As Double, SyN As Double                                ' Desviaciones estándar para intervalos de confianza
Dim r As Double, R2 As Double                                  ' Coeficientes de correlación y determinación
Dim i As Long                                                  ' Índice para bucles
Dim t As Double                                                ' F^(-1)(alfa, n-2)
Dim SemiAmplitud                                               ' Semiamplitud de intervalo de confianza
Dim xmin As Double, xmax As Double                             ' Extremos de abscisas

'
' ---> Lectura de los vectores x() e y(), introducidos como rangos.
'
n = XRango.Rows.Count
i = YRango.Rows.Count
If n <> i Then
   Ajuste_Lineal_yMax = "Los rangos deben tener la misma longitud"
   Exit Function
End If

ReDim x(1 To n)
ReDim y(1 To n)

xmin = 1E+99: xmax = -1E+99
For i = 1 To n
    x(i) = XRango.Cells(i, 1)
    y(i) = YRango.Cells(i, 1)
    xM = xM + x(i)
    yM = yM + y(i)
    If x(i) < xmin Then xmin = x(i)
    If x(i) > xmax Then xmax = x(i)
Next

If alfa <= 0 Or alfa > 1 Then
   Ajuste_Lineal_yMax = "Alfa Debe ser >0 y <1"
   Exit Function
End If

' Cálculo de estadísticos
xM = xM / n: yM = yM / n                                      ' Valores medios
Cov = 0: S2x = 0: S2y = 0: RSS = 0                            ' Cuasivarianzas
For i = 1 To n
    S2x = S2x + (x(i) - xM) ^ 2
    S2y = S2y + (y(i) - yM) ^ 2
    Cov = Cov + (x(i) - xM) * (y(i) - yM)
Next
TSS = S2y
S2x = S2x / (n - 1): S2y = S2y / (n - 1): Cov = Cov / (n - 1)
Sx = Sqr(S2x): Sy = Sqr(S2y)                                  ' Desviaciones estándar

' Coeficientes de la recta
b = Cov / S2x
a = yM - b * xM

' Estimaciones de los valores introducidos
RSS = 0
For i = 1 To n
    yE = b * x(i) + a
    RSS = RSS + (y(i) - yE) ^ 2
Next
sR = Sqr(RSS / (n - 2))

' Coeficientes de correlación y determinación
r = Sx / Sy * b
R2 = 1 - RSS / TSS

' t de Student
t = F_t_Student_Inv(1 - alfa / 2, n - 2)
    
' Intervalos de confianza
yE = b * xx + a
SyN = sR * Sqr(1 / n + (xx - xM) ^ 2 / ((n - 1) * S2x))
SemiAmplitud = t * SyN
YEMin = yE - SemiAmplitud
YEMax = yE + SemiAmplitud

Ajuste_Lineal_yMax = YEMax

End Function


' FUNCIÓN Ajuste_Lineal_r

Public Function Ajuste_Lineal_r(XRango As Range, YRango As Range) As Variant
Dim x() As Double, y() As Double                               ' Datos de entrada
Dim yE As Double                                               ' y estimada
Dim n As Long                                                  ' Número de puntos
Dim xM As Double, yM As Double                                 ' Valores medios
Dim Cov As Double, S2x As Double, S2y As Double                ' Cuasivarianzas
Dim Sx As Double, Sy As Double                                 ' Desviaciones estándar
Dim a As Double, b As Double                                   ' Coeficientes de la recta y = bx + a
Dim RSS As Double, TSS As Double                               ' Errores cuadráticos
Dim sR As Double, SyN As Double                                ' Desviaciones estándar para intervalos de confianza
Dim r As Double, R2 As Double                                  ' Coeficientes de correlación y determinación
Dim i As Long                                                  ' Índice para bucles
Dim xmin As Double, xmax As Double                             ' Extremos de abscisas

'
' ---> Lectura de los vectores x() e y(), introducidos como rangos.
'
n = XRango.Rows.Count
i = YRango.Rows.Count
If n <> i Then
   Ajuste_Lineal_r = "Los rangos deben tener la misma longitud"
   Exit Function
End If

ReDim x(1 To n)
ReDim y(1 To n)

xmin = 1E+99: xmax = -1E+99
For i = 1 To n
    x(i) = XRango.Cells(i, 1)
    y(i) = YRango.Cells(i, 1)
    xM = xM + x(i)
    yM = yM + y(i)
    If x(i) < xmin Then xmin = x(i)
    If x(i) > xmax Then xmax = x(i)
Next

' Cálculo de estadísticos
xM = xM / n: yM = yM / n                                      ' Valores medios
Cov = 0: S2x = 0: S2y = 0: RSS = 0                            ' Cuasivarianzas
For i = 1 To n
    S2x = S2x + (x(i) - xM) ^ 2
    S2y = S2y + (y(i) - yM) ^ 2
    Cov = Cov + (x(i) - xM) * (y(i) - yM)
Next
TSS = S2y
S2x = S2x / (n - 1): S2y = S2y / (n - 1): Cov = Cov / (n - 1)
Sx = Sqr(S2x): Sy = Sqr(S2y)                                  ' Desviaciones estándar

' Coeficientes de la recta
b = Cov / S2x
a = yM - b * xM

' Estimaciones de los valores introducidos
RSS = 0
For i = 1 To n
    yE = b * x(i) + a
    RSS = RSS + (y(i) - yE) ^ 2
Next
sR = Sqr(RSS / (n - 2))

' Coeficientes de correlación y determinación
r = Sx / Sy * b
R2 = 1 - RSS / TSS

Ajuste_Lineal_r = r

End Function


' RUTINA ANOVA

Sub ANOVA()
'
' Realización tabla ANOVA
'
Dim FilaIn As Integer, FilaFin As Integer    ' Filas y columnas
Dim ColIn As Integer, ColFin As Integer      ' que contienen los datos
Dim NFilas As Integer, NColumnas As Integer  ' Dimensiones matriz
Dim NTotal As Integer                        ' Nº total de datos
Dim x() As Double                            ' Matriz con los datos
Dim Mui() As Double                          ' Matriz con las medias por método
Dim S2i() As Double                          ' Matriz con las cuasivarianzas
Dim s2 As Double                             ' Cuasivarianza muestral
Dim SI() As Double                           ' Sumas en cads método
Dim Mu As Double                             ' Media global
Dim i As Integer, j As Integer               ' Variables enteras auxiliares
Dim iIni As Integer, jIni As Integer         ' Posiciones de referencia para escribir
Dim SCT As Double                            ' Suma de cuadrados total
Dim SCI As Double                            ' Suma de cuadrados intragrupos
Dim SCE As Double                            ' Suma de cuadrados entre grupos
Dim s As Double                              ' Suma de todas las mediciones
Dim F As Double                              ' Estadístico F
Dim FSnedecor As Double                      ' Valor de la F de Snedecor directa
                                             ' NColumnas - 1, NTotal - NColumnas
Dim FSnedecorI As Double                     ' Valor de la F de Snedecor inversa
                                             ' 1 - Alfa, NColumnas - 1, NTotal - NColumnas
Dim alfa As Double                           ' Nivel de confidencialidad
Dim Resultado As String                      ' Resultado del test
Dim Color As Long                            ' Variable con el color del resultado
Dim ColFin_2 As Integer                      ' Referencia para la posición de la tabla ANOVA
Dim rng_tabla As Range                       ' Variable Rango de Excel auxiliar

' Leemos la posición de los datos
FilaIn = Cells(2, 1): FilaFin = Cells(2, 2)
ColIn = Cells(2, 3): ColFin = Cells(2, 4)
alfa = Cells(2, 5)

If (FilaIn <= 3 And ColFin <= 4) Then
   MsgBox "La primera fila de datos debe ser posterior a la 3ª.", vbExclamation + vbOKOnly, "Atención"
   Exit Sub
End If
If (ColIn <= 1) Then
   MsgBox "La primera columna de datos debe ser la 2ª o posterior.", vbExclamation + vbOKOnly, "Atención"
   Exit Sub
End If

' Primeros cálculos
NFilas = FilaFin - FilaIn + 1
NColumnas = ColFin - ColIn + 1
NTotal = NFilas * NColumnas

' Posiciones de referencia para escribir resultados
iIni = FilaFin + 1
jIni = ColIn

If (ColFin < 4) Then
  ColFin_2 = 5
Else
  ColFin_2 = ColFin + 1
End If

' Limpiamos la zona de escritura de resultados
'Range(Cells(FilaIn - 1, ColFin + 1), Cells(Rows.Count, Columns.Count)).Select
Range(Cells(FilaIn - 1, ColFin + 1), Cells(Rows.Count, Columns.Count)).Clear
'Range(Cells(iIni, 1), Cells(Rows.Count, ColFin)).Select
Range(Cells(iIni, 1), Cells(Rows.Count, ColFin)).Clear

' Damos el formato a la zona de escritura de resultados
'Range(Cells(iIni, 1), Cells(iIni + 6, ColFin)).Select
Range(Cells(iIni, 1), Cells(iIni + 6, ColFin)).NumberFormat = "0.000"
Range(Cells(iIni, 1), Cells(iIni + 6, ColFin)).HorizontalAlignment = xlCenter
Range(Cells(iIni, 1), Cells(iIni + 6, ColFin)).Font.Bold = False

'Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 3, ColFin_2 + 7)).Select
Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 3, ColFin_2 + 7)).NumberFormat = "0.000"
Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 3, ColFin_2 + 7)).HorizontalAlignment = xlCenter
Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 3, ColFin_2 + 7)).Font.Bold = False

' Damos el formato al bloque fijo de datos
Range(Cells(1, 1), Cells(1, 5)).HorizontalAlignment = xlCenter
Range(Cells(1, 1), Cells(1, 5)).Font.Bold = True

Range(Cells(2, 1), Cells(2, 4)).NumberFormat = "0"
Cells(2, 5).NumberFormat = "0.000"
Range(Cells(2, 1), Cells(2, 5)).HorizontalAlignment = xlCenter
Range(Cells(2, 1), Cells(2, 5)).Font.Bold = False
Cells(2, 5).Font.Color = -6279056

' Ahora centrados en azul los datos introducidos
Range(Cells(FilaIn, ColIn), Cells(FilaFin, ColFin)).ClearFormats
Range(Cells(FilaIn, ColIn), Cells(FilaFin, ColFin)).HorizontalAlignment = xlCenter
Range(Cells(FilaIn, ColIn), Cells(FilaFin, ColFin)).Font.TintAndShade = 0
Range(Cells(FilaIn, ColIn), Cells(FilaFin, ColFin)).Font.Bold = False
Range(Cells(FilaIn, ColIn), Cells(FilaFin, ColFin)).Font.Color = -1003520

' Ajustamos las dimensiones de las matrices
ReDim x(1 To NFilas, 1 To NColumnas)
ReDim Mui(1 To NColumnas)
ReDim SI(1 To NColumnas)
ReDim S2i(1 To NColumnas)

' Leemos la matriz de datos
For i = 1 To NFilas
  For j = 1 To NColumnas
    x(i, j) = Cells(i + FilaIn - 1, j + ColIn - 1)
  Next j
Next i

' Media global, varianza
' y suma de cuadrados total
s = 0: SCT = 0
For i = 1 To NColumnas
  Mui(i) = 0: S2i(i) = 0
  For j = 1 To NFilas
    SI(i) = SI(i) + x(j, i)
    s = s + x(j, i)
    SCT = SCT + x(j, i) ^ 2
    S2i(i) = S2i(i) + x(j, i) ^ 2
  Next j
  Mui(i) = SI(i) / NFilas
Next i
Mu = s / NTotal
SCT = SCT - s * s / NFilas / NColumnas
s2 = SCT / (NTotal - 1)

For i = 1 To NColumnas
  S2i(i) = (S2i(i) - NFilas * Mui(i) * Mui(i)) / (NFilas - 1)
Next

' Suma de cuadrados total
SCT = 0
For i = 1 To NColumnas
  For j = 1 To NFilas
    SCT = SCT + (x(j, i) - Mu) ^ 2
  Next j
Next i

' Suma de cuadrados entre grupos
SCE = 0
For i = 1 To NColumnas
    SCE = SCE + SI(i) * SI(i)
Next i
SCE = SCE / NFilas - s * s / NFilas / NColumnas

' Suma de cuadrados intragrupos
SCI = SCT - SCE

' Estimador F
F = SCE * (NTotal - NColumnas) / SCI / (NColumnas - 1)

' F de Snedecor
FSnedecor = 1 - FD_F_Snedecor(F, NColumnas - 1, NTotal - NColumnas)
FSnedecorI = F_F_Snedecor_Inv(1 - alfa, NColumnas - 1, NTotal - NColumnas)

' Resultado
If FSnedecor > alfa Then
   Resultado = "ACEPTAMOS H0 PARA " & ChrW(945) & " = " & alfa
   Color = -1003520
Else
   Resultado = "NO ACEPTAMOS H0 PARA " & ChrW(945) & " = " & alfa
   Color = -16776961
End If

' Escribimos valores

Cells(iIni, jIni - 1) = "x" & ChrW(773) & ChrW(7522)      ' (x media sub-i)
For i = 1 To NColumnas
    Cells(iIni, jIni - 1 + i).Value = Mui(i)
Next i

Cells(iIni + 1, jIni - 1) = "s" & ChrW(7522) & ChrW(178)  ' Cuasivarianza sub-i
For i = 1 To NColumnas
    Cells(iIni + 1, jIni - 1 + i).Value = S2i(i)
Next i

Cells(iIni + 2, jIni - 1) = "x" & ChrW(773)               ' (x media)
Cells(iIni + 2, jIni) = s / NTotal

Cells(iIni + 3, jIni - 1) = "s" & ChrW(178)               ' Cuasivarianza
Cells(iIni + 3, jIni) = s2

Cells(iIni + 4, jIni - 1) = "N"
Cells(iIni + 4, jIni) = NTotal
Cells(iIni + 4, jIni).NumberFormat = "0"
Cells(iIni + 5, jIni - 1) = "m"
Cells(iIni + 5, jIni) = NColumnas
Cells(iIni + 5, jIni).NumberFormat = "0"
Cells(iIni + 6, jIni - 1) = "n"
Cells(iIni + 6, jIni) = NFilas
Cells(iIni + 6, jIni).NumberFormat = "0"

Cells(FilaIn - 1, ColFin_2 + 1) = "TABLA ANOVA"
Cells(FilaIn - 1, ColFin_2 + 1).HorizontalAlignment = xlLeft
Cells(FilaIn, ColFin_2 + 1) = "VARIACIÓN"
Cells(FilaIn, ColFin_2 + 2) = "G. de L."
Cells(FilaIn, ColFin_2 + 3) = "SUMA CUAD."
Cells(FilaIn, ColFin_2 + 4) = "S.C./g.d.l."
Cells(FilaIn, ColFin_2 + 5) = "F"
Cells(FilaIn, ColFin_2 + 6) = "1-F(F)"
Cells(FilaIn, ColFin_2 + 7) = "f(1-" & ChrW(945) & ")"

Cells(FilaIn + 1, ColFin_2 + 1) = "SCE"
Cells(FilaIn + 1, ColFin_2 + 2) = NColumnas - 1
Cells(FilaIn + 1, ColFin_2 + 2).NumberFormat = "0"
Cells(FilaIn + 1, ColFin_2 + 3) = SCE
Cells(FilaIn + 1, ColFin_2 + 4) = SCE / (NColumnas - 1)

Cells(FilaIn + 2, ColFin_2 + 1) = "SCI"
Cells(FilaIn + 2, ColFin_2 + 2) = NTotal - NColumnas
Cells(FilaIn + 2, ColFin_2 + 2).NumberFormat = "0"
Cells(FilaIn + 2, ColFin_2 + 3) = SCI
Cells(FilaIn + 2, ColFin_2 + 4) = SCI / (NTotal - NColumnas)

Cells(FilaIn + 3, ColFin_2 + 1) = "SCT"
Cells(FilaIn + 3, ColFin_2 + 2).NumberFormat = "0"
Cells(FilaIn + 3, ColFin_2 + 3) = SCT

Cells(FilaIn + 1, ColFin_2 + 5) = F
Cells(FilaIn + 1, ColFin_2 + 5).Font.Color = -11489280

Cells(FilaIn + 1, ColFin_2 + 6) = FSnedecor
Cells(FilaIn + 1, ColFin_2 + 6).Font.Color = -6279056

Cells(FilaIn + 1, ColFin_2 + 7) = FSnedecorI
Cells(FilaIn + 1, ColFin_2 + 7).Font.Color = -11489280

Cells(FilaIn + 3, ColFin_2 + 6).Characters().Font.Superscript = False
Cells(FilaIn + 3, ColFin_2 + 6).Characters().Font.Subscript = False
Cells(FilaIn + 3, ColFin_2 + 6) = Resultado
Cells(FilaIn + 3, ColFin_2 + 6).Characters(Start:=InStr(Resultado, "H0") + 1, Length:=1).Font.Subscript = True
Cells(FilaIn + 3, ColFin_2 + 6).Font.Color = Color

'Range(Cells(FilaIn + 1, ColFin + 2), Cells(FilaIn + 3, ColFin + 2)).NumberFormat = "0"

' Ponemos en negrita los títulos
Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn, ColFin_2 + 7)).Font.Bold = True

' Ponemos bordes en el cuadro ANOVA

Set rng_tabla = Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 3, ColFin_2 + 7))

rng_tabla.Borders(xlDiagonalDown).LineStyle = xlNone
rng_tabla.Borders(xlDiagonalUp).LineStyle = xlNone
With rng_tabla.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With

Set rng_tabla = Nothing

Range("A1").Select

End Sub


' RUTINA ANOVA21

Sub ANOVA21()
'
' Realización tabla ANOVA con 2 factores y un ensayo (sin replicación)
'
Dim FilaIn As Integer, FilaFin As Integer    ' Filas y columnas
Dim ColIn As Integer, ColFin As Integer      ' que contiene los datos
Dim NFilas As Integer, NColumnas As Integer  ' Dimensiones matriz
Dim NTotal As Integer                        ' Nº total de datos
Dim x() As Double                            ' Matriz con los datos
Dim MuF() As Double, MuC() As Double         ' Matriz con las medias por fila/columna
Dim S2F() As Double, S2C() As Double         ' Matriz con las cuasivarianzas por fila/columna
Dim s2 As Double                             ' Cuasivarianza muestral
Dim SiF() As Double, SiC() As Double         ' Sumas en cada método por fila/columna
Dim Mu As Double                             ' Media global
Dim i As Integer, j As Integer               ' Variables enteras auxiliares
Dim iIni As Integer, jIni As Integer         ' Posiciones de referencia para escribir
Dim SCT As Double                            ' Suma de cuadrados total
Dim SCE As Double                            ' Suma de cuadrados intragrupos
Dim SCE_F As Double, SCE_C As Double         ' Suma de cuadrados entre grupos
Dim s As Double                              ' Suma de todas las mediciones
Dim FF As Double, FC As Double               ' Estimadores F
Dim FSnedecorF As Double                     ' Valor de la F de Snedecor directa
Dim FSnedecorC As Double
Dim FSnedecorFI As Double                    ' Valor de la F de Snedecor inversa
Dim FSnedecorCI As Double
Dim alfa As Double                           ' Nivel de confidencialidad
Dim ResultadoF As String                     ' Resultados del test
Dim ResultadoC As String
Dim ColorF As Long, ColorC As Long           ' Variables con el color del resultado
Dim ColFin_2 As Integer                      ' Referencia para la posición de la tabla ANOVA
Dim rng_tabla As Range                       ' Variable Rango de Excel auxiliar

' Leemos la posición de los datos
FilaIn = Cells(2, 1): FilaFin = Cells(2, 2)
ColIn = Cells(2, 3): ColFin = Cells(2, 4)
alfa = Cells(2, 5)

If (FilaIn <= 3 And ColFin <= 4) Then
   MsgBox "La primera fila de datos debe ser posterior a la 3ª.", vbExclamation + vbOKOnly, "Atención"
   Exit Sub
End If
If (ColIn <= 1) Then
   MsgBox "La primera columna de datos debe ser la 2ª o posterior.", vbExclamation + vbOKOnly, "Atención"
   Exit Sub
End If

' Primeros cálculos
NFilas = FilaFin - FilaIn + 1
NColumnas = ColFin - ColIn + 1
NTotal = NFilas * NColumnas

' Posiciones de referencia para escribir resultados
iIni = FilaFin + 1
jIni = ColIn

ColFin_2 = ColFin + 3

' Limpiamos la zona de escritura de resultados
'Range(Cells(FilaIn - 1, ColFin + 1), Cells(Rows.Count, Columns.Count)).Select
Range(Cells(FilaIn - 1, ColFin + 1), Cells(Rows.Count, Columns.Count)).Clear
'Range(Cells(iIni, 1), Cells(Rows.Count, ColFin)).Select
Range(Cells(iIni, 1), Cells(Rows.Count, ColFin)).Clear

' Damos el formato a la zona de escritura de resultados
'Range(Cells(iIni, 1), Cells(iIni + 6, ColFin)).Select
Range(Cells(iIni, 1), Cells(iIni + 6, ColFin)).NumberFormat = "0.000"
Range(Cells(iIni, 1), Cells(iIni + 6, ColFin)).HorizontalAlignment = xlCenter
Range(Cells(iIni, 1), Cells(iIni + 6, ColFin)).Font.Bold = False

'Range(Cells(FilaIn - 1, ColFin + 1), Cells(FilaFin, ColFin + 2)).Select
Range(Cells(FilaIn - 1, ColFin + 1), Cells(FilaFin, ColFin + 2)).NumberFormat = "0.000"
Range(Cells(FilaIn - 1, ColFin + 1), Cells(FilaFin, ColFin + 2)).HorizontalAlignment = xlCenter
Range(Cells(FilaIn - 1, ColFin + 1), Cells(FilaFin, ColFin + 2)).Font.Bold = False

'Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 4, ColFin_2 + 7)).Select
Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 4, ColFin_2 + 7)).NumberFormat = "0.000"
Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 4, ColFin_2 + 7)).HorizontalAlignment = xlCenter
Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 4, ColFin_2 + 7)).Font.Bold = False

' Damos el formato al bloque fijo de datos
Range(Cells(1, 1), Cells(1, 5)).HorizontalAlignment = xlCenter
Range(Cells(1, 1), Cells(1, 5)).Font.Bold = True

Range(Cells(2, 1), Cells(2, 4)).NumberFormat = "0"
Cells(2, 5).NumberFormat = "0.000"
Range(Cells(2, 1), Cells(2, 5)).HorizontalAlignment = xlCenter
Range(Cells(2, 1), Cells(2, 5)).Font.Bold = False
Cells(2, 5).Font.Color = -6279056

' Ahora centrados en azul los datos introducidos
Range(Cells(FilaIn, ColIn), Cells(FilaFin, ColFin)).ClearFormats
Range(Cells(FilaIn, ColIn), Cells(FilaFin, ColFin)).HorizontalAlignment = xlCenter
Range(Cells(FilaIn, ColIn), Cells(FilaFin, ColFin)).Font.TintAndShade = 0
Range(Cells(FilaIn, ColIn), Cells(FilaFin, ColFin)).Font.Bold = False
Range(Cells(FilaIn, ColIn), Cells(FilaFin, ColFin)).Font.Color = -1003520

' Ajustamos las dimensiones de las matrices
ReDim x(1 To NFilas, 1 To NColumnas)
ReDim MuC(1 To NColumnas)
ReDim MuF(1 To NFilas)
ReDim SiC(1 To NColumnas)
ReDim SiF(1 To NFilas)
ReDim S2C(1 To NColumnas)
ReDim S2F(1 To NFilas)

' Leemos la matriz de datos
For i = 1 To NFilas
  For j = 1 To NColumnas
    x(i, j) = Cells(i + FilaIn - 1, j + ColIn - 1)
  Next j
Next i

' Media global, varianza
' y suma de cuadrados total
s = 0: SCT = 0
For i = 1 To NColumnas
  MuC(i) = 0: S2C(i) = 0: SiC(i) = 0
  For j = 1 To NFilas
    SiC(i) = SiC(i) + x(j, i)
    s = s + x(j, i)
    SCT = SCT + x(j, i) ^ 2
    S2C(i) = S2C(i) + x(j, i) ^ 2
  Next j
  MuC(i) = SiC(i) / NFilas
Next i

For i = 1 To NFilas
  MuF(i) = 0: S2F(i) = 0: SiF(i) = 0
  For j = 1 To NColumnas
    SiF(i) = SiF(i) + x(i, j)
    S2F(i) = S2F(i) + x(i, j) ^ 2
  Next
  MuF(i) = SiF(i) / NColumnas
Next

Mu = s / NTotal
SCT = SCT - s * s / NFilas / NColumnas
s2 = SCT / (NTotal - 1)

For i = 1 To NFilas
  S2F(i) = (S2F(i) - NColumnas * MuF(i) * MuF(i)) / (NColumnas - 1)
Next

For i = 1 To NColumnas
  S2C(i) = (S2C(i) - NFilas * MuC(i) * MuC(i)) / (NFilas - 1)
Next

' Suma de cuadrados total
SCT = 0
For i = 1 To NColumnas
  For j = 1 To NFilas
    SCT = SCT + (x(j, i) - Mu) ^ 2
  Next j
Next i

' Suma de cuadrados entre grupos
SCE_C = 0
For i = 1 To NColumnas
    SCE_C = SCE_C + SiC(i) * SiC(i)
Next i
SCE_C = SCE_C / NFilas - s * s / NFilas / NColumnas

SCE_F = 0
For i = 1 To NFilas
    SCE_F = SCE_F + SiF(i) * SiF(i)
Next i
SCE_F = SCE_F / NColumnas - s * s / NFilas / NColumnas

' Suma de cuadrados intragrupos
SCE = SCT - SCE_F - SCE_C

' Estimadores F
FF = SCE_F * (NFilas - 1) * (NColumnas - 1) / SCE / (NFilas - 1)
FC = SCE_C * (NFilas - 1) * (NColumnas - 1) / SCE / (NColumnas - 1)

' F de Snedecor
FSnedecorF = 1 - FD_F_Snedecor(FF, NFilas - 1, (NFilas - 1) * (NColumnas - 1))
FSnedecorFI = F_F_Snedecor_Inv(1 - alfa, NFilas - 1, (NFilas - 1) * (NColumnas - 1))

FSnedecorC = 1 - FD_F_Snedecor(FC, NColumnas - 1, (NFilas - 1) * (NColumnas - 1))
FSnedecorCI = F_F_Snedecor_Inv(1 - alfa, NColumnas - 1, (NFilas - 1) * (NColumnas - 1))

' Resultados
If FSnedecorF > alfa Then
   ResultadoF = "ACEPTAMOS H0F PARA " & ChrW(945) & " = " & alfa
   ColorF = -1003520
Else
   ResultadoF = "NO ACEPTAMOS H0F PARA " & ChrW(945) & " = " & alfa
   ColorF = -16776961
End If

If FSnedecorC > alfa Then
   ResultadoC = "ACEPTAMOS H0C PARA " & ChrW(945) & " = " & alfa
   ColorC = -1003520
Else
   ResultadoC = "NO ACEPTAMOS H0C PARA " & ChrW(945) & " = " & alfa
   ColorC = -16776961
End If

' Escribimos valores

Cells(iIni, jIni - 1) = "x" & ChrW(773) & ("." & ChrW(7522))     ' (x. media sub-i)
For i = 1 To NColumnas
    Cells(iIni, jIni - 1 + i).Value = MuC(i)
Next i

Cells(iIni + 1, jIni - 1) = "s." & ChrW(7522) & ChrW(178)  ' Cuasivarianza sub-i
For i = 1 To NColumnas
    Cells(iIni + 1, jIni - 1 + i).Value = S2C(i)
Next i

Cells(iIni + 2, jIni - 1) = "x" & ChrW(773)               ' (x media)
Cells(iIni + 2, jIni) = s / NTotal

Cells(iIni + 3, jIni - 1) = "s" & ChrW(178)               ' Cuasivarianza
Cells(iIni + 3, jIni) = s2

Cells(iIni + 4, jIni - 1) = "N"
Cells(iIni + 4, jIni) = NTotal
Cells(iIni + 4, jIni).NumberFormat = "0"
Cells(iIni + 5, jIni - 1) = "m"
Cells(iIni + 5, jIni) = NColumnas
Cells(iIni + 5, jIni).NumberFormat = "0"
Cells(iIni + 6, jIni - 1) = "n"
Cells(iIni + 6, jIni) = NFilas
Cells(iIni + 6, jIni).NumberFormat = "0"

Cells(FilaIn - 1, ColFin + 1) = "x" & ChrW(773) & (ChrW(7522) & ".")   ' (x media sub-i punto)
Cells(FilaIn - 1, ColFin + 2) = "s" & (ChrW(7522) & (".")) & ChrW(178) ' Cuasivarianza sub-i
For i = 1 To NFilas
   Cells(FilaIn - 1 + i, ColFin + 1) = MuF(i)
   Cells(FilaIn - 1 + i, ColFin + 2) = S2F(i)
   Cells(FilaIn - 1 + i, ColFin + 1).NumberFormat = "0.000"
   Cells(FilaIn - 1 + i, ColFin + 2).NumberFormat = "0.000"
Next

Cells(FilaIn - 1, ColFin_2 + 1) = "TABLA ANOVA"
Cells(FilaIn - 1, ColFin_2 + 1).HorizontalAlignment = xlLeft
Cells(FilaIn, ColFin_2 + 1) = "VARIACIÓN"
Cells(FilaIn, ColFin_2 + 2) = "G. de L."
Cells(FilaIn, ColFin_2 + 3) = "SUMA CUAD."
Cells(FilaIn, ColFin_2 + 4) = "S.C./g.d.l."
Cells(FilaIn, ColFin_2 + 5) = "F"
Cells(FilaIn, ColFin_2 + 6) = "1-F(F)"
Cells(FilaIn, ColFin_2 + 7) = "f(1-" & ChrW(945) & ")"

Cells(FilaIn + 1, ColFin_2 + 1).Characters().Font.Superscript = False
Cells(FilaIn + 1, ColFin_2 + 1).Characters().Font.Subscript = False
Cells(FilaIn + 1, ColFin_2 + 1) = "SCEF"
Cells(FilaIn + 1, ColFin_2 + 1).Characters(Start:=4, Length:=1).Font.Subscript = True
Cells(FilaIn + 1, ColFin_2 + 2) = NFilas - 1
Cells(FilaIn + 1, ColFin_2 + 2).NumberFormat = "0"
Cells(FilaIn + 1, ColFin_2 + 3) = SCE_F
Cells(FilaIn + 1, ColFin_2 + 4) = SCE_F / (NFilas - 1)
Cells(FilaIn + 1, ColFin_2 + 5) = FF

Cells(FilaIn + 2, ColFin_2 + 1).Characters().Font.Superscript = False
Cells(FilaIn + 2, ColFin_2 + 1).Characters().Font.Subscript = False
Cells(FilaIn + 2, ColFin_2 + 1) = "SCEC"
Cells(FilaIn + 2, ColFin_2 + 1).Characters(Start:=4, Length:=1).Font.Subscript = True
Cells(FilaIn + 2, ColFin_2 + 2) = NColumnas - 1
Cells(FilaIn + 2, ColFin_2 + 2).NumberFormat = "0"
Cells(FilaIn + 2, ColFin_2 + 3) = SCE_C
Cells(FilaIn + 2, ColFin_2 + 4) = SCE_C / (NColumnas - 1)
Cells(FilaIn + 2, ColFin_2 + 5) = FC

Cells(FilaIn + 3, ColFin_2 + 1) = "SCE"
Cells(FilaIn + 3, ColFin_2 + 2) = (NFilas - 1) * (NColumnas - 1)
Cells(FilaIn + 3, ColFin_2 + 2).NumberFormat = "0"
Cells(FilaIn + 3, ColFin_2 + 3) = SCE
Cells(FilaIn + 3, ColFin_2 + 4) = SCE / (NFilas - 1) / (NColumnas - 1)

Cells(FilaIn + 4, ColFin_2 + 1) = "SCT"
Cells(FilaIn + 4, ColFin_2 + 2) = NFilas * NColumnas - 1
Cells(FilaIn + 4, ColFin_2 + 2).NumberFormat = "0"
Cells(FilaIn + 4, ColFin_2 + 3) = SCT

Cells(FilaIn + 1, ColFin_2 + 5).Font.Color = -11489280

Cells(FilaIn + 1, ColFin_2 + 6) = FSnedecorF
Cells(FilaIn + 1, ColFin_2 + 6).Font.Color = -6279056

Cells(FilaIn + 2, ColFin_2 + 5).Font.Color = -11489280

Cells(FilaIn + 2, ColFin_2 + 6) = FSnedecorC
Cells(FilaIn + 2, ColFin_2 + 6).Font.Color = -6279056

Cells(FilaIn + 1, ColFin_2 + 7) = FSnedecorFI
Cells(FilaIn + 1, ColFin_2 + 7).Font.Color = -11489280

Cells(FilaIn + 2, ColFin_2 + 7) = FSnedecorCI
Cells(FilaIn + 2, ColFin_2 + 7).Font.Color = -11489280

Cells(FilaIn + 3, ColFin_2 + 6).Characters().Font.Superscript = False
Cells(FilaIn + 3, ColFin_2 + 6).Characters().Font.Subscript = False
Cells(FilaIn + 3, ColFin_2 + 6) = ResultadoF
Cells(FilaIn + 3, ColFin_2 + 6).Characters(Start:=InStr(ResultadoF, "H0F") + 1, Length:=2).Font.Subscript = True
Cells(FilaIn + 3, ColFin_2 + 6).Font.Color = ColorF

Cells(FilaIn + 4, ColFin_2 + 6).Characters().Font.Superscript = False
Cells(FilaIn + 4, ColFin_2 + 6).Characters().Font.Subscript = False
Cells(FilaIn + 4, ColFin_2 + 6) = ResultadoC
Cells(FilaIn + 4, ColFin_2 + 6).Characters(Start:=InStr(ResultadoC, "H0C") + 1, Length:=2).Font.Subscript = True
Cells(FilaIn + 4, ColFin_2 + 6).Font.Color = ColorC

' Ponemos en negrita los títulos
Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn, ColFin_2 + 7)).Font.Bold = True

' Ponemos bordes en el cuadro ANOVA

Set rng_tabla = Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 4, ColFin_2 + 7))

rng_tabla.Borders(xlDiagonalDown).LineStyle = xlNone
rng_tabla.Borders(xlDiagonalUp).LineStyle = xlNone
With rng_tabla.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With

Set rng_tabla = Nothing

Range("A1").Select

End Sub


' RUTINA LIMPIA

Sub Limpia()
Dim FilaIn As Integer, FilaFin As Integer   ' Filas y columnas
Dim ColIn As Integer, ColFin As Integer     ' que contiene los datos
Dim iIni As Integer, jIni As Integer        ' Posiciones de referencia para escribir
Dim respuesta As Integer                    ' Respuesta del usuario para confirmar el borrado

respuesta = MsgBox("¿Seguro que desea eliminar los datos existentes? ", vbQuestion + vbOKCancel, "Atención")

If respuesta <> 1 Then Exit Sub

' Leemos la posición de los datos
FilaIn = Cells(2, 1): FilaFin = Cells(2, 2)
ColIn = Cells(2, 3): ColFin = Cells(2, 4)

iIni = FilaFin + 1

If (FilaIn <= 3 And ColFin <= 4) Then
  Range(Cells(FilaIn, ColFin + 1), Cells(Rows.Count, Columns.Count)).Clear
Else
  Range(Cells(FilaIn - 1, ColFin + 1), Cells(Rows.Count, Columns.Count)).Clear
End If

Range(Cells(iIni, 1), Cells(Rows.Count, ColFin)).Clear

' Limpiamos los datos de entrada anteriores
Range(Cells(FilaIn, ColIn), Cells(FilaFin, ColFin)).Clear

Range("A1").Select

End Sub


' FUNCIÓN ANOVA2R

Sub ANOVA2R()
'
' Realización tabla ANOVA con 2 factores y replicación
'
Dim FilaIn As Integer, FilaFin As Integer     ' Filas y columnas
Dim ColIn As Integer, ColFin As Integer       ' que contienen los datos
Dim NFilas As Integer, NColumnas As Integer   ' Dimensiones matriz
Dim NReplicas As Integer
Dim NTotal As Integer                         ' Nº total de datos
Dim x() As Double                             ' Matriz con los datos
Dim xioo() As Double, xojo() As Double        ' Medias parciales
Dim xook() As Double
Dim xijo() As Double, xiok() As Double
Dim xojk() As Double
Dim xiooCh As String, xojoCh As String        ' Medias parciales (Texto)
Dim xookCh As String
Dim xijoCh As String, xiokCh As String
Dim xojkCh As String
Dim xRaya As String, PuntoC As String         ' Caracteres especiales
Dim Mu As Double                              ' Media global
Dim i As Integer, j As Integer, k As Integer  ' Variables enteras auxiliares
Dim iIni As Integer, jIni As Integer          ' Posiciones de referencia para escribir
Dim SCT As Double                             ' Suma de cuadrados total
Dim SCE As Double                             ' Suma de cuadrados residual
Dim SCE_F As Double, SCE_C As Double          ' Suma de cuadrados entre filas y columnas
Dim SCE_I As Double                           ' Suma de cuadrados entre réplicas
Dim GDL_F As Integer, GDL_C As Integer        ' Grados de libertad
Dim GDL_I As Integer, GDL_E As Integer
Dim SF As Double, SC As Double                ' Variaciones/g.d.l.
Dim SI As Double, SE As Double
Dim ff As Double, FC As Double, Fi As Double  ' Estimadores F
Dim FSnedecorF As Double                      ' Valor de la F de Snedecor directa
Dim FSnedecorC As Double
Dim FSnedecorI As Double
Dim FSnedecorFI As Double                     ' Valor de la F de Snedecor inversa
Dim FSnedecorCI As Double
Dim FSnedecorII As Double
Dim Alfa As Double                            ' Nivel de confidencialidad
Dim ResultadoF As String                      ' Resultados del test
Dim ResultadoC As String
Dim ResultadoI As String
Dim ColorF As Long, ColorC As Long            ' Variables con el color del resultado
Dim ColorI As Long
Dim NCuadrosHSal As Integer
Dim ColFin_2 As Integer                       ' Referencia para la posición de la tabla ANOVA
Dim rng_tabla As Range                        ' Variable Rango de Excel auxiliar

' Nombres de las medias parciales
xRaya = "x" & ChrW(773)
PuntoC = "."
xiooCh = xRaya & "i" & PuntoC & PuntoC
xojoCh = xRaya & PuntoC & "j" & PuntoC
xookCh = xRaya & PuntoC & PuntoC & "k"
xijoCh = xRaya & "i" & "j" & PuntoC
xiokCh = xRaya & "i" & PuntoC & "k"
xojkCh = xRaya & PuntoC & "j" & "k"

' Leemos la posición de los datos
FilaIn = Cells(2, 1): FilaFin = Cells(2, 2)
ColIn = Cells(2, 3): ColFin = Cells(2, 4)
NReplicas = Cells(2, 5)
Alfa = Cells(2, 6)

If (FilaIn <= 3 And ColFin <= 5) Then
   MsgBox "La primera fila de datos debe ser posterior a la 3ª.", vbExclamation + vbOKOnly, "Atención"
   Exit Sub
End If

' Primeros cálculos
NFilas = FilaFin - FilaIn + 1
NColumnas = ColFin - ColIn + 1
NTotal = NFilas * NColumnas * NReplicas

' Posiciones de referencia para escribir resultados
iIni = FilaFin + (NReplicas - 1) * (1 + NFilas) + 1

If (ColFin < 5) Then
  ColFin_2 = 6
Else
  ColFin_2 = ColFin + 1
End If

' Limpiamos la zona de escritura de resultados
'Range(Cells(FilaIn - 1, ColFin + 1), Cells(Rows.Count, Columns.Count)).Select
Range(Cells(FilaIn - 1, ColFin + 1), Cells(Rows.Count, Columns.Count)).Clear
'Range(Cells(iIni, 1), Cells(Rows.Count, ColFin)).Select
Range(Cells(iIni, 1), Cells(Rows.Count, ColFin)).Clear

' Damos el formato a la zona de escritura de resultados
'Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 6, ColFin_2 + 7)).Select
Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 6, ColFin_2 + 7)).NumberFormat = "0.000"
Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 6, ColFin_2 + 7)).HorizontalAlignment = xlCenter
Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 6, ColFin_2 + 7)).Font.Bold = False

' Damos el formato al bloque fijo de datos
Range(Cells(1, 1), Cells(1, 6)).HorizontalAlignment = xlCenter
Range(Cells(1, 1), Cells(1, 6)).Font.Bold = True

Range(Cells(2, 1), Cells(2, 5)).NumberFormat = "0"
Cells(2, 6).NumberFormat = "0.000"
Range(Cells(2, 1), Cells(2, 6)).HorizontalAlignment = xlCenter
Range(Cells(2, 1), Cells(2, 6)).Font.Bold = False
Cells(2, 6).Font.Color = -6279056

' Ahora centrados en azul los datos introducidos
Range(Cells(FilaIn, ColIn), Cells(FilaFin + (NReplicas - 1) * (1 + NFilas), ColFin)).ClearFormats
Range(Cells(FilaIn, ColIn), Cells(FilaFin + (NReplicas - 1) * (1 + NFilas), ColFin)).HorizontalAlignment = xlCenter
Range(Cells(FilaIn, ColIn), Cells(FilaFin + (NReplicas - 1) * (1 + NFilas), ColFin)).Font.TintAndShade = 0
Range(Cells(FilaIn, ColIn), Cells(FilaFin + (NReplicas - 1) * (1 + NFilas), ColFin)).Font.Bold = False
Range(Cells(FilaIn, ColIn), Cells(FilaFin + (NReplicas - 1) * (1 + NFilas), ColFin)).Font.Color = -1003520

' Ajustamos las dimensiones de las matrices
ReDim x(1 To NFilas, 1 To NColumnas, 1 To NReplicas)
ReDim xioo(1 To NFilas)
ReDim xojo(1 To NColumnas)
ReDim xook(1 To NReplicas)

ReDim xijo(1 To NFilas, 1 To NColumnas)
ReDim xiok(1 To NFilas, 1 To NReplicas)
ReDim xojk(1 To NColumnas, 1 To NReplicas)

' Leemos la matriz de datos
For i = 1 To NFilas
For j = 1 To NColumnas
For k = 1 To NReplicas
    x(i, j, k) = Cells(i + FilaIn - 1 + (k - 1) * (NFilas + 1), j + ColIn - 1)
Next k
Next j
Next i

' Sumas parciales
For i = 1 To NFilas
For j = 1 To NColumnas
  xijo(i, j) = 0
  For k = 1 To NReplicas
     xijo(i, j) = xijo(i, j) + x(i, j, k)
  Next k
  xijo(i, j) = xijo(i, j) / NReplicas
Next j
Next i

For i = 1 To NFilas
For k = 1 To NReplicas
  xiok(i, k) = 0
  For j = 1 To NColumnas
     xiok(i, k) = xiok(i, k) + x(i, j, k)
  Next j
  xiok(i, k) = xiok(i, k) / NColumnas
Next k
Next i

For j = 1 To NColumnas
For k = 1 To NReplicas
  xojk(j, k) = 0
  For i = 1 To NFilas
     xojk(j, k) = xojk(j, k) + x(i, j, k)
  Next i
  xojk(j, k) = xojk(j, k) / NFilas
Next k
Next j

For i = 1 To NFilas
  xioo(i) = 0
  For j = 1 To NColumnas
     xioo(i) = xioo(i) + xijo(i, j)
  Next j
  xioo(i) = xioo(i) / NColumnas
Next i

For j = 1 To NColumnas
  xojo(j) = 0
  For i = 1 To NFilas
     xojo(j) = xojo(j) + xijo(i, j)
  Next i
  xojo(j) = xojo(j) / NFilas
Next j

For k = 1 To NReplicas
  xook(k) = 0
  For i = 1 To NFilas
     xook(k) = xook(k) + xiok(i, k)
  Next i
  xook(k) = xook(k) / NFilas
Next k

' Media total
Mu = 0
For i = 1 To NFilas
  Mu = Mu + xioo(i)
Next
Mu = Mu / NFilas

' Variaciones
SCT = 0
For i = 1 To NFilas
For j = 1 To NColumnas
For k = 1 To NReplicas
    SCT = SCT + (x(i, j, k) - Mu) ^ 2
Next k
Next j
Next i

SCE_F = 0
For i = 1 To NFilas
  SCE_F = SCE_F + (xioo(i) - Mu) ^ 2
Next i
SCE_F = NColumnas * NReplicas * SCE_F

SCE_C = 0
For j = 1 To NColumnas
  SCE_C = SCE_C + (xojo(j) - Mu) ^ 2
Next j
SCE_C = NFilas * NReplicas * SCE_C

SCE_I = 0
For i = 1 To NFilas
For j = 1 To NColumnas
  SCE_I = SCE_I + (xijo(i, j) - xioo(i) - xojo(j) + Mu) ^ 2
Next j
Next i
SCE_I = SCE_I * NReplicas

SCE = SCT - SCE_F - SCE_C - SCE_I

' Estimadores F
GDL_F = NFilas - 1
GDL_C = NColumnas - 1
GDL_I = (NFilas - 1) * (NColumnas - 1)
GDL_E = NFilas * NColumnas * (NReplicas - 1)

SF = SCE_F / GDL_F
SC = SCE_C / GDL_C
SI = SCE_I / GDL_I
SE = SCE / GDL_E

ff = SF / SE
FC = SC / SE
Fi = SI / SE

' F de Snedecor
FSnedecorF = 1 - FD_F_Snedecor(ff, GDL_F, GDL_E)
FSnedecorFI = F_F_Snedecor_Inv(1 - Alfa, GDL_F, GDL_E)

FSnedecorC = 1 - FD_F_Snedecor(FC, GDL_C, GDL_E)
FSnedecorCI = F_F_Snedecor_Inv(1 - Alfa, GDL_C, GDL_E)

FSnedecorI = 1 - FD_F_Snedecor(Fi, GDL_I, GDL_E)
FSnedecorII = F_F_Snedecor_Inv(1 - Alfa, GDL_I, GDL_E)

' Resultados
If FSnedecorF > Alfa Then
   ResultadoF = "ACEPTAMOS H0F PARA " & ChrW(945) & " = " & Alfa
   ColorF = -1003520
Else
   ResultadoF = "NO ACEPTAMOS H0F PARA " & ChrW(945) & " = " & Alfa
   ColorF = -16776961
End If

If FSnedecorC > Alfa Then
   ResultadoC = "ACEPTAMOS H0C PARA " & ChrW(945) & " = " & Alfa
   ColorC = -1003520
Else
   ResultadoC = "NO ACEPTAMOS H0C PARA " & ChrW(945) & " = " & Alfa
   ColorC = -16776961
End If

If FSnedecorI > Alfa Then
   ResultadoI = "ACEPTAMOS H0I PARA " & ChrW(945) & " = " & Alfa
   ColorI = -1003520
Else
   ResultadoI = "NO ACEPTAMOS H0I PARA " & ChrW(945) & " = " & Alfa
   ColorI = -16776961
End If

' Vamos a la página auxiliar ANOVA E3-Sumas para escribir los cuadros de
' medias parciales y primero la limpiamos

Sheets("ANOVA E3-Sumas").Select

With Cells
  .Clear
  .HorizontalAlignment = xlCenter
  .Font.Size = 11
End With

iIni = 1: jIni = 1

Cells(iIni, jIni).Characters().Font.Superscript = False
Cells(iIni, jIni).Characters().Font.Subscript = False

Cells(iIni, jIni + NColumnas + 1).Characters().Font.Superscript = False
Cells(iIni, jIni + NColumnas + 1).Characters().Font.Subscript = False

Cells(iIni + NFilas + 1, jIni).Characters().Font.Superscript = False
Cells(iIni + NFilas + 1, jIni).Characters().Font.Subscript = False

Cells(iIni, jIni) = xijoCh
Cells(iIni, jIni + NColumnas + 1) = xiooCh
Cells(iIni + NFilas + 1, jIni) = xojoCh

Cells(iIni, jIni).Characters(Start:=3, Length:=3).Font.Subscript = True
Cells(iIni, jIni).Font.Size = 14
Cells(iIni, jIni + NColumnas + 1).Characters(Start:=3, Length:=3).Font.Subscript = True
Cells(iIni, jIni + NColumnas + 1).Font.Size = 14
Cells(iIni + NFilas + 1, jIni).Characters(Start:=3, Length:=3).Font.Subscript = True
Cells(iIni + NFilas + 1, jIni).Font.Size = 14

For i = 1 To NFilas
    Cells(iIni + i, jIni) = i
    Cells(iIni + i, jIni + NColumnas + 1) = xioo(i)
    For j = 1 To NColumnas
       Cells(iIni, jIni + j) = j
       Cells(iIni + i, jIni + j) = xijo(i, j)
       Cells(iIni + NFilas + 1, jIni + j) = xojo(j)
    Next j
Next i

iIni = 1: jIni = NColumnas + 5

Cells(iIni, jIni).Characters().Font.Superscript = False
Cells(iIni, jIni).Characters().Font.Subscript = False

Cells(iIni, jIni + NReplicas + 1).Characters().Font.Superscript = False
Cells(iIni, jIni + NReplicas + 1).Characters().Font.Subscript = False

Cells(iIni + NFilas + 1, jIni).Characters().Font.Superscript = False
Cells(iIni + NFilas + 1, jIni).Characters().Font.Subscript = False

Cells(iIni, jIni) = xiokCh
Cells(iIni, jIni + NReplicas + 1) = xiooCh
Cells(iIni + NFilas + 1, jIni) = xookCh

Cells(iIni, jIni).Characters(Start:=3, Length:=3).Font.Subscript = True
Cells(iIni, jIni).Font.Size = 14
Cells(iIni, jIni + NReplicas + 1).Characters(Start:=3, Length:=3).Font.Subscript = True
Cells(iIni, jIni + NReplicas + 1).Font.Size = 14
Cells(iIni + NFilas + 1, jIni).Characters(Start:=3, Length:=3).Font.Subscript = True
Cells(iIni + NFilas + 1, jIni).Font.Size = 14

For i = 1 To NFilas
    Cells(iIni + i, jIni) = i
    Cells(iIni + i, jIni + NReplicas + 1) = xioo(i)
    For k = 1 To NReplicas
       Cells(iIni, jIni + k) = k
       Cells(iIni + i, jIni + k) = xiok(i, k)
       Cells(iIni + NFilas + 1, jIni + k) = xook(k)
    Next k
Next i

iIni = 1 + NFilas + 4: jIni = 1

Cells(iIni, jIni).Characters().Font.Superscript = False
Cells(iIni, jIni).Characters().Font.Subscript = False

Cells(iIni, jIni + NReplicas + 1).Characters().Font.Superscript = False
Cells(iIni, jIni + NReplicas + 1).Characters().Font.Subscript = False

Cells(iIni + NColumnas + 1, jIni).Characters().Font.Superscript = False
Cells(iIni + NColumnas + 1, jIni).Characters().Font.Subscript = False

Cells(iIni, jIni) = xojkCh
Cells(iIni, jIni + NReplicas + 1) = xojoCh
Cells(iIni + NColumnas + 1, jIni) = xookCh

Cells(iIni, jIni).Characters(Start:=3, Length:=3).Font.Subscript = True
Cells(iIni, jIni).Font.Size = 14
Cells(iIni, jIni + NReplicas + 1).Characters(Start:=3, Length:=3).Font.Subscript = True
Cells(iIni, jIni + NReplicas + 1).Font.Size = 14
Cells(iIni + NColumnas + 1, jIni).Characters(Start:=3, Length:=3).Font.Subscript = True
Cells(iIni + NColumnas + 1, jIni).Font.Size = 14

For j = 1 To NColumnas
    Cells(iIni + j, jIni) = j
    Cells(iIni + j, jIni + NReplicas + 1) = xojo(j)
    For k = 1 To NReplicas
       Cells(iIni, jIni + k) = k
       Cells(iIni + j, jIni + k) = xojk(j, k)
       Cells(iIni + NColumnas + 1, jIni + k) = xook(k)
    Next k
Next j

Range("A1").Select

' Volvemos a la página "ANOVA E3"

Sheets("ANOVA E3").Select

Cells(FilaIn - 1, ColFin_2 + 1) = "TABLA ANOVA"
Cells(FilaIn - 1, ColFin_2 + 1).HorizontalAlignment = xlLeft
Cells(FilaIn, ColFin_2 + 1) = "VARIACIÓN"
Cells(FilaIn, ColFin_2 + 2) = "G. de L."
Cells(FilaIn, ColFin_2 + 3) = "SUMA CUAD."
Cells(FilaIn, ColFin_2 + 4) = "S.C./g.d.l."
Cells(FilaIn, ColFin_2 + 5) = "F"
Cells(FilaIn, ColFin_2 + 6) = "1-F(F)"
Cells(FilaIn, ColFin_2 + 7) = "f(1-" & ChrW(945) & ")"

Cells(FilaIn + 1, ColFin_2 + 1).Characters().Font.Superscript = False
Cells(FilaIn + 1, ColFin_2 + 1).Characters().Font.Subscript = False
Cells(FilaIn + 1, ColFin_2 + 1) = "SCEF"
Cells(FilaIn + 1, ColFin_2 + 1).Characters(Start:=4, Length:=1).Font.Subscript = True
Cells(FilaIn + 1, ColFin_2 + 2) = GDL_F
Cells(FilaIn + 1, ColFin_2 + 2).NumberFormat = "0"
Cells(FilaIn + 1, ColFin_2 + 3) = SCE_F
Cells(FilaIn + 1, ColFin_2 + 4) = SF
Cells(FilaIn + 1, ColFin_2 + 5) = ff

Cells(FilaIn + 2, ColFin_2 + 1).Characters().Font.Superscript = False
Cells(FilaIn + 2, ColFin_2 + 1).Characters().Font.Subscript = False
Cells(FilaIn + 2, ColFin_2 + 1) = "SCEC"
Cells(FilaIn + 2, ColFin_2 + 1).Characters(Start:=4, Length:=1).Font.Subscript = True
Cells(FilaIn + 2, ColFin_2 + 2) = GDL_C
Cells(FilaIn + 2, ColFin_2 + 2).NumberFormat = "0"
Cells(FilaIn + 2, ColFin_2 + 3) = SCE_C
Cells(FilaIn + 2, ColFin_2 + 4) = SC
Cells(FilaIn + 2, ColFin_2 + 5) = FC

Cells(FilaIn + 3, ColFin_2 + 1).Characters().Font.Superscript = False
Cells(FilaIn + 3, ColFin_2 + 1).Characters().Font.Subscript = False
Cells(FilaIn + 3, ColFin_2 + 1) = "SCEI"
Cells(FilaIn + 3, ColFin_2 + 1).Characters(Start:=4, Length:=1).Font.Subscript = True
Cells(FilaIn + 3, ColFin_2 + 2) = GDL_I
Cells(FilaIn + 3, ColFin_2 + 2).NumberFormat = "0"
Cells(FilaIn + 3, ColFin_2 + 3) = SCE_I
Cells(FilaIn + 3, ColFin_2 + 4) = SI
Cells(FilaIn + 3, ColFin_2 + 5) = Fi

Cells(FilaIn + 4, ColFin_2 + 1) = "SCE"
Cells(FilaIn + 4, ColFin_2 + 2) = GDL_E
Cells(FilaIn + 4, ColFin_2 + 2).NumberFormat = "0"
Cells(FilaIn + 4, ColFin_2 + 3) = SCE
Cells(FilaIn + 4, ColFin_2 + 4) = SE

Cells(FilaIn + 5, ColFin_2 + 1) = "SCT"
Cells(FilaIn + 5, ColFin_2 + 3) = SCT
Cells(FilaIn + 5, ColFin_2 + 2) = NFilas * NColumnas * NReplicas - 1
Cells(FilaIn + 5, ColFin_2 + 2).NumberFormat = "0"

Cells(FilaIn + 7, ColFin_2 + 1) = "Media total"
Cells(FilaIn + 7, ColFin_2 + 2) = xRaya
Cells(FilaIn + 7, ColFin_2 + 3) = Mu
Range(Cells(FilaIn + 7, ColFin_2 + 1), Cells(FilaIn + 7, ColFin_2 + 3)).NumberFormat = "0.000"
Range(Cells(FilaIn + 7, ColFin_2 + 1), Cells(FilaIn + 7, ColFin_2 + 3)).HorizontalAlignment = xlCenter
Range(Cells(FilaIn + 7, ColFin_2 + 1), Cells(FilaIn + 7, ColFin_2 + 2)).Font.Bold = True

Cells(FilaIn + 1, ColFin_2 + 6) = FSnedecorF
Range(Cells(FilaIn + 1, ColFin_2 + 6), Cells(FilaIn + 1, ColFin_2 + 6)).Font.Color = -6279056

Cells(FilaIn + 2, ColFin_2 + 6) = FSnedecorC
Range(Cells(FilaIn + 2, ColFin_2 + 6), Cells(FilaIn + 2, ColFin_2 + 6)).Font.Color = -6279056

Cells(FilaIn + 3, ColFin_2 + 6) = FSnedecorI
Range(Cells(FilaIn + 3, ColFin_2 + 6), Cells(FilaIn + 2, ColFin_2 + 6)).Font.Color = -6279056

Cells(FilaIn + 1, ColFin_2 + 7) = FSnedecorFI
Range(Cells(FilaIn + 3, ColFin_2 + 7), Cells(FilaIn + 1, ColFin_2 + 7)).Font.Color = -11489280

Cells(FilaIn + 2, ColFin_2 + 7) = FSnedecorCI
Range(Cells(FilaIn + 3, ColFin_2 + 7), Cells(FilaIn + 2, ColFin_2 + 7)).Font.Color = -11489280

Cells(FilaIn + 3, ColFin_2 + 7) = FSnedecorII
Range(Cells(FilaIn + 3, ColFin_2 + 7), Cells(FilaIn + 2, ColFin_2 + 7)).Font.Color = -11489280

Cells(FilaIn + 4, ColFin_2 + 6).Characters().Font.Superscript = False
Cells(FilaIn + 4, ColFin_2 + 6).Characters().Font.Subscript = False
Cells(FilaIn + 4, ColFin_2 + 6) = ResultadoF
Cells(FilaIn + 4, ColFin_2 + 6).Characters(Start:=InStr(ResultadoF, "H0F") + 1, Length:=2).Font.Subscript = True
Range(Cells(FilaIn + 4, ColFin_2 + 6), Cells(FilaIn + 4, ColFin_2 + 6)).Font.Color = ColorF

Cells(FilaIn + 5, ColFin_2 + 6).Characters().Font.Superscript = False
Cells(FilaIn + 5, ColFin_2 + 6).Characters().Font.Subscript = False
Cells(FilaIn + 5, ColFin_2 + 6) = ResultadoC
Cells(FilaIn + 5, ColFin_2 + 6).Characters(Start:=InStr(ResultadoC, "H0C") + 1, Length:=2).Font.Subscript = True
Range(Cells(FilaIn + 5, ColFin_2 + 6), Cells(FilaIn + 5, ColFin_2 + 6)).Font.Color = ColorC

Cells(FilaIn + 6, ColFin_2 + 6).Characters().Font.Superscript = False
Cells(FilaIn + 6, ColFin_2 + 6).Characters().Font.Subscript = False
Cells(FilaIn + 6, ColFin_2 + 6) = ResultadoI
Cells(FilaIn + 6, ColFin_2 + 6).Characters(Start:=InStr(ResultadoI, "H0I") + 1, Length:=2).Font.Subscript = True
Range(Cells(FilaIn + 6, ColFin_2 + 6), Cells(FilaIn + 6, ColFin_2 + 6)).Font.Color = ColorI

' Agrupamos celdas
With Range(Cells(FilaIn + 5, ColFin_2 + 1), Cells(FilaIn + 6, ColFin_2 + 1))
  .MergeCells = True
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlCenter
End With
With Range(Cells(FilaIn + 5, ColFin_2 + 2), Cells(FilaIn + 6, ColFin_2 + 2))
  .MergeCells = True
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlCenter
End With
With Range(Cells(FilaIn + 5, ColFin_2 + 3), Cells(FilaIn + 6, ColFin_2 + 3))
  .MergeCells = True
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlCenter
End With
With Range(Cells(FilaIn + 5, ColFin_2 + 4), Cells(FilaIn + 6, ColFin_2 + 4))
  .MergeCells = True
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlCenter
End With

' Ponemos en negrita los títulos
Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn, ColFin_2 + 7)).Font.Bold = True

' Ponemos bordes en el cuadro ANOVA

Set rng_tabla = Range(Cells(FilaIn - 1, ColFin_2 + 1), Cells(FilaIn + 6, ColFin_2 + 7))

rng_tabla.Borders(xlDiagonalDown).LineStyle = xlNone
rng_tabla.Borders(xlDiagonalUp).LineStyle = xlNone
With rng_tabla.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
With rng_tabla.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With

Set rng_tabla = Nothing

Range("A1").Select

End Sub


' RUTINA LIMPIA2R

Sub Limpia2R()
Dim FilaIn As Integer, FilaFin As Integer    ' Filas y columnas
Dim ColIn As Integer, ColFin As Integer      ' que contiene los datos
Dim NFilas As Integer, NColumnas As Integer  ' Dimensiones matriz
Dim NReplicas As Integer
Dim iIni As Integer, jIni As Integer         ' Posiciones de referencia para escribir
Dim respuesta As Integer                     ' Respuesta del usuario para confirmar el borrado

respuesta = MsgBox("¿Seguro que desea eliminar los datos existentes? ", vbQuestion + vbOKCancel, "Atención")

If respuesta <> 1 Then Exit Sub

' Limpiamos la página auxiliar ANOVA E3-Sumas

Sheets("ANOVA E3-Sumas").Select

Cells.Clear

Range("A1").Select

Sheets("ANOVA E3").Select

' Leemos la posición de los datos
FilaIn = Cells(2, 1): FilaFin = Cells(2, 2)
ColIn = Cells(2, 3): ColFin = Cells(2, 4)
NReplicas = Cells(2, 5)

NFilas = FilaFin - FilaIn + 1
iIni = FilaFin + (NReplicas - 1) * (1 + NFilas) + 1

If (FilaIn <= 3 And ColFin <= 5) Then
  Range(Cells(FilaIn, ColFin + 1), Cells(Rows.Count, Columns.Count)).Clear
Else
  Range(Cells(FilaIn - 1, ColFin + 1), Cells(Rows.Count, Columns.Count)).Clear
End If

Range(Cells(iIni, 1), Cells(Rows.Count, ColFin)).Clear

' Limpiamos los datos de entrada anteriores
Range(Cells(FilaIn, ColIn), Cells(FilaFin + (NReplicas - 1) * (1 + NFilas), ColFin)).Clear

Range("A1").Select

End Sub


' RUTINA AjusteLinealN

Sub AjusteLinealN()
' Ajuste lineal a una nube de M puntos, en un espacio n-dimensional
' Llama a las rutinas:
'   Producto(), que realiza una operación matricial
'           (mutiplicación de una matriz por su transpuesta)
'   WorksheetFunction.MInverse que es una rutina nativa de EXCEL
'   TestNormalidad(), que realiza el test de normalidad
'   Test_Breusch_Pagan(), que realiza el test de homocedasticidad
' Y a las funciones estadísticas:
'   F_t_Student_Inv
'   FD_F_Snedecor
'   F_F_Snedecor_Inv

Dim x(1 To MMx, 1 To nMx) As Double
Dim y() As Double, yy() As Double
Dim Res() As Double, Res2() As Double
Dim ResMin As Double, ResMax As Double, dRes As Double
Dim a() As Double, SE()
Dim XtX() As Double
Dim XtXi() As Variant
Dim XtXiXt() As Double
Dim iIni As Integer, jIni As Integer, iIni0 As Integer
Dim i As Integer, j As Integer, k As Integer, Id As Integer
Dim N As Integer, m As Integer
Dim yM As Double, SCE As Double, SCR As Double, SCT As Double
Dim F0 As Double, R2 As Double
Dim Alfa As Double, FSnedecor As Double
Dim VarRes As Double, tStudent As Double
Dim nRango As Integer, MediaRes As Double, SigmaRes As Double
Dim Rango() As Double, Frec() As Double, FrecNormal() As Double
Dim ChiQ0 As Double, ChiQAlfa As Double
Dim R2_BP As Double, ChQ0_BP As Double, ChQAlfa_BP As Double, ChAlfaBP As Double
Dim aBP() As Double, ResBP() As Double, Res2BP() As Double

' Leemos los datos en la hoja de entrada
N = Range("B1").Value
m = Range("B2").Value
Alfa = Range("E1").Value
' iIni = Range("B3"): jIni = Range("C3")
nRango = Range("E2").Value

iIni = 5: jIni = 1
iIni0 = iIni - 1

ReDim XtX(1 To N + 1, 1 To N + 1)
ReDim XtXi(1 To N + 1, 1 To N + 1)
ReDim XtXiXt(1 To N + 1, 1 To m)
ReDim a(1 To N + 1)
ReDim SE(1 To N + 1)
ReDim y(1 To m)
ReDim yy(1 To m)
ReDim Res(1 To m)
ReDim Res2(1 To m)

ReDim Rango(1 To nRango + 1)
ReDim Frec(1 To nRango + 1)
ReDim FrecNormal(1 To nRango + 1)
ReDim aBP(1 To N + 1)
ReDim ResBP(1 To m)
ReDim Res2BP(1 To m)

For j = 1 To N
  Cells(iIni - 1, jIni + j - 1).Value = "x" & j
  Cells(iIni - 1, jIni + j - 1).Characters(Start:=2, Length:=1).Font.Subscript = True
Next j

Cells(iIni - 1, jIni + N).Value = "y"
Cells(iIni - 1, jIni + N + 1).Value = ChrW(375)   ' carácter "y" con acento circunflejo
Cells(iIni - 1, jIni + N + 2).Value = ChrW(949)   ' carácter épsilon
Cells(iIni - 1, jIni + N + 3).Value = ChrW(949) & "2"
Cells(iIni - 1, jIni + N + 3).Characters(Start:=2, Length:=1).Font.Superscript = True   ' épsilon elevado al cuadrado

Range(Cells(iIni - 1, jIni), Cells(iIni - 1, jIni + N + 3)).HorizontalAlignment = xlCenter

' Matriz de coeficientes
yM = 0
For i = 1 To m
   x(i, 1) = 1
   For j = 2 To N + 1
      x(i, j) = Cells(iIni - 1 + i, jIni - 2 + j)
   Next
   y(i) = Cells(iIni - 1 + i, jIni - 1 + N + 1)
   yM = yM + y(i)
Next
yM = yM / m   ' Media

' Sub Producto(n As Integer, M As Integer, X As Double, XtX As Double)
' Primero redimensionamos las matrices para evitar que la función
' WorksheetFunction.MInverse dé errores

ReDim XtX(1 To N + 1, 1 To N + 1)
ReDim XtXi(1 To N + 1, 1 To N + 1)
ReDim XtXiXt(1 To N + 1, 1 To m)

' Calculamos el producto de la matriz Xt por X
' y  llamamos al resultado XtX
Call Producto(N + 1, m, x, XtX)

' Invertimos XtX llamando a una función nativa de EXCEL
' La matriz resultante XtXi debe de ser tipo Variant
XtXi = WorksheetFunction.MInverse(XtX)

' Multiplicamos por la  matriz X transpuesta y lo metemos en XtXiXt
' XtXi entra como Variant, X es Double y XtXiXt es Double
For i = 1 To N + 1
  For j = 1 To m
    XtXiXt(i, j) = 0
    For k = 1 To N + 1
      XtXiXt(i, j) = XtXiXt(i, j) + XtXi(i, k) * x(j, k)
    Next k
  Next j
Next i

' Resolvemos el sistema multiplicando XtXiXt(n+1 x M) por y(M x 1)
For i = 1 To N + 1
   a(i) = 0
   For j = 1 To m
      a(i) = a(i) + XtXiXt(i, j) * y(j)
   Next j
Next i

' Escribimos los coeficientes del ajuste
For i = 1 To N + 1
   Cells(iIni + m + 1, i) = "a" & (i - 1)
   Cells(iIni + m + 1, i).Characters(Start:=2, Length:=1).Font.Subscript = True
   Cells(iIni + m + 2, i) = a(i)
Next

' Escribimos estimaciones
SCE = 0: SCR = 0: ResMin = 99999: ResMax = -99999
For i = 1 To m
   yy(i) = 0
   For j = 1 To N + 1
      yy(i) = yy(i) + x(i, j) * a(j)
   Next j
   Res(i) = y(i) - yy(i)
   If Res(i) < ResMin Then ResMin = Res(i)
   If Res(i) > ResMax Then ResMax = Res(i)
   Res2(i) = Res(i) * Res(i)
   Cells(iIni - 1 + i, N + 2) = yy(i)
   Cells(iIni - 1 + i, N + 3) = Res(i)
   Cells(iIni - 1 + i, N + 4) = Res2(i)
   SCR = SCR + Res2(i)
   SCE = SCE + (yy(i) - yM) ^ 2
Next i
VarRes = SCR / (m - N - 1)  ' Varianza residual

' Calculamos los errores estándar de las estimaciones
' como diagonal de la matriz VarRes x XtXi
For i = 1 To N + 1
   SE(i) = Sqr(VarRes * XtXi(i, i))
Next

' Escribimos los errores del ajuste y el intervalo de confianza
For i = 1 To N + 1
   Cells(iIni + m + 3, i) = "e" & (i - 1)
   Cells(iIni + m + 3, i).Characters(Start:=2, Length:=1).Font.Subscript = True
   Cells(iIni + m + 4, i) = SE(i)
Next

tStudent = F_t_Student_Inv(1 - Alfa / 2, m - N + 1)
Cells(iIni + m + 3, N + 2) = "t1-" & ChrW(945) & "/2"
Cells(iIni + m + 3, i).Characters(Start:=2, Length:=5).Font.Subscript = True
Cells(iIni + m + 4, i) = tStudent

For i = 1 To N + 1
   Cells(iIni + m + 5, i) = "a" & (i - 1) & "Mín"
   Cells(iIni + m + 5, i).Characters(Start:=2, Length:=6).Font.Subscript = True
   Cells(iIni + m + 6, i) = a(i) - tStudent * SE(i)
Next
For i = 1 To N + 1
   Cells(iIni + m + 7, i) = "a" & (i - 1) & "Máx"
   Cells(iIni + m + 7, i).Characters(Start:=2, Length:=6).Font.Subscript = True
   Cells(iIni + m + 8, i) = a(i) + tStudent * SE(i)
Next

'  Variaciones
SCT = SCE + SCR
R2 = SCE / SCT
F0 = (SCE / N) / (SCR / (m - N - 1))
FSnedecor = F_F_Snedecor_Inv(1 - Alfa, N, m - N - 1)

Id = 6 + m ' Salto con el grupo anterior

Cells(iIni + Id + 3, 1).Value = ChrW(563)  ' (y media) carácter "y" con una raya encima (macrón o acento largo)
Cells(iIni + Id + 3, 2) = yM

Cells(iIni + Id + 4, 1).Value = "SR2"
Cells(iIni + Id + 4, 1).Characters(Start:=2, Length:=1).Font.Subscript = True
Cells(iIni + Id + 4, 1).Characters(Start:=3, Length:=1).Font.Superscript = True
Cells(iIni + Id + 4, 2) = VarRes
Cells(iIni + Id + 4, 3).Value = "SR"
Cells(iIni + Id + 4, 3).Characters(Start:=2, Length:=1).Font.Subscript = True
Cells(iIni + Id + 4, 4) = Sqr(VarRes)

Cells(iIni + Id + 5, 1) = "R2"
Cells(iIni + Id + 5, 1).Characters(Start:=2, Length:=1).Font.Superscript = True
Cells(iIni + Id + 5, 2) = R2

Cells(iIni + Id + 6, 1) = "SCE": Cells(iIni + Id + 6, 2) = SCE
Cells(iIni + Id + 7, 1) = "SCR": Cells(iIni + Id + 7, 2) = SCR
Cells(iIni + Id + 8, 1) = "SCT": Cells(iIni + Id + 8, 2) = SCT

Cells(iIni + Id + 9, 1) = "F0"
Cells(iIni + Id + 9, 1).Characters(Start:=2, Length:=1).Font.Subscript = True
Cells(iIni + Id + 9, 2) = F0
Cells(iIni + Id + 9, 3) = "P[F>F0]"
Cells(iIni + Id + 9, 3).Characters(Start:=6, Length:=1).Font.Subscript = True
Cells(iIni + Id + 9, 4) = 1 - FD_F_Snedecor(F0, N, m - N - 1)
Cells(iIni + Id + 10, 1) = "F1-" & ChrW(945) & "(n,M-n-1)"
Cells(iIni + Id + 10, 1).Characters(Start:=2, Length:=3).Font.Subscript = True
Cells(iIni + Id + 10, 2) = FSnedecor

If F0 < FSnedecor Then
   Cells(iIni + Id + 11, 2) = "Se acepta H0"
   Cells(iIni + Id + 11, 2).Characters(Start:=12, Length:=1).Font.Subscript = True
   With Cells(iIni + Id + 11, 2).Font
      .Color = -1003520
      .TintAndShade = 0
      .Bold = True
   End With
Else
   Cells(iIni + Id + 11, 2) = "NO se acepta H0"
   Cells(iIni + Id + 11, 2).Characters(Start:=15, Length:=1).Font.Subscript = True
   With Cells(iIni + Id + 11, 2).Font
      .Color = -16776961
      .TintAndShade = 0
      .Bold = True
   End With
End If

' Escribimos la tabla ANOVA
iIni = iIni + Id + 3
jIni = 7

' Líneas del cuadro
Range(Cells(iIni, jIni), Cells(iIni + 4, jIni + 6)).Borders(xlDiagonalDown).LineStyle = xlNone
Range(Cells(iIni, jIni), Cells(iIni + 4, jIni + 6)).Borders(xlDiagonalUp).LineStyle = xlNone
With Range(Cells(iIni, jIni), Cells(iIni + 4, jIni + 6)).Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range(Cells(iIni, jIni), Cells(iIni + 4, jIni + 6)).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range(Cells(iIni, jIni), Cells(iIni + 4, jIni + 6)).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range(Cells(iIni, jIni), Cells(iIni + 4, jIni + 6)).Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range(Cells(iIni, jIni), Cells(iIni + 4, jIni + 6)).Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Range(Cells(iIni, jIni), Cells(iIni + 4, jIni + 6)).Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

' Relleno de la tabla
Cells(iIni, jIni) = "TABLA ANOVA: H0 (Todos los ai son nulos) vs algún aj no lo es"
Cells(iIni, jIni).Characters(Start:=29, Length:=1).Font.Subscript = True
Cells(iIni, jIni).Characters(Start:=52, Length:=1).Font.Subscript = True
Cells(iIni, jIni).HorizontalAlignment = xlGeneral
Cells(iIni + 1, jIni) = "VARIACIÓN"
Cells(iIni + 2, jIni) = "SCE"
Cells(iIni + 3, jIni) = "SCR"
Cells(iIni + 4, jIni) = "SCT"

jIni = jIni + 1
Cells(iIni + 1, jIni) = "G. de L."
Cells(iIni + 2, jIni) = N
Cells(iIni + 3, jIni) = m - N - 1
Cells(iIni + 4, jIni) = m - 1

jIni = jIni + 1
Cells(iIni + 1, jIni) = "SUMA CUAD."
Cells(iIni + 2, jIni) = SCE
Cells(iIni + 3, jIni) = SCR
Cells(iIni + 4, jIni) = SCT

jIni = jIni + 1
Cells(iIni + 1, jIni) = "S.C./g.d.l."
Cells(iIni + 2, jIni) = SCE / N
Cells(iIni + 3, jIni) = SCR / (m - N - 1)

jIni = jIni + 1
Cells(iIni + 1, jIni) = "F0"
Cells(iIni + Id + 11, 2).Characters(Start:=2, Length:=1).Font.Subscript = True
Cells(iIni + 2, jIni) = F0
Cells(iIni + 2, jIni).Font.Color = -11489280  ' Color verde

jIni = jIni + 1
Cells(iIni + 1, jIni) = "P[F>F0]"
Cells(iIni + 1, jIni).Characters(Start:=6, Length:=1).Font.Subscript = True
Cells(iIni + 2, jIni) = 1 - FD_F_Snedecor(F0, N, m - N - 1)
Cells(iIni + 2, jIni).Font.Color = -6279056  ' Color violeta


jIni = jIni + 1
Cells(iIni + 1, jIni) = "F1-" & ChrW(945) & "(n,M-n-1)"
Cells(iIni + 1, jIni).Characters(Start:=2, Length:=3).Font.Subscript = True
Cells(iIni + 2, jIni) = FSnedecor
Cells(iIni + 2, jIni).Font.Color = -11489280  ' Color verde


If F0 < FSnedecor Then
   Cells(iIni + 3, jIni - 2) = "ACEPTAMOS H0 PARA " & ChrW(945) & " = " & Alfa
   Cells(iIni + 3, jIni - 2).Characters(Start:=12, Length:=1).Font.Subscript = True
   With Cells(iIni + 3, jIni - 2).Font
      .Color = -1003520
      .TintAndShade = 0
      .Bold = True
   End With
Else
   Cells(iIni + 3, jIni - 2) = "NO ACEPTAMOS H0 PARA " & ChrW(945) & " = " & Alfa
   Cells(iIni + 3, jIni - 2).Characters(Start:=15, Length:=1).Font.Subscript = True
   With Cells(iIni + 3, jIni - 2).Font
      .Color = -16776961
      .TintAndShade = 0
      .Bold = True
   End With
End If
Cells(iIni + 3, jIni - 2).HorizontalAlignment = xlLeft

' Hacemos test de normalidad de los residuos
dRes = (ResMax - ResMin) / nRango
Rango(1) = ResMin

Cells(iIni0, 9) = "Rangos"
Cells(iIni0 + 1, 9) = Rango(1)
For i = 2 To nRango + 1
    Rango(i) = Rango(i - 1) + dRes
    Cells(iIni0 + i, 9) = Rango(i)
Next

' Llamamos a la rutina que efectúa el test de normalidad
Call TestNormalidad(m, nRango, Alfa, Res(), MediaRes, SigmaRes, _
                    ChiQ0, ChiQAlfa, Rango(), Frec(), FrecNormal())

' Escribimos frecuencias acumuladas y las de la distribución normal
Cells(iIni0, 10) = "Frec. Ac."
Cells(iIni0, 11) = "F.Dist. N(0," & ChrW(963) & ChrW(949) & ")"
Cells(iIni0, 11).Characters(Start:=14, Length:=1).Font.Subscript = True  ' sigma sub-épsilon
For i = 1 To nRango + 1
    Cells(iIni0 + i, 10) = Frec(i)
    Cells(iIni0 + i, 11) = FrecNormal(i)
Next

' Escribimos título del apartado
Cells(iIni0 - 1, 9) = "TEST DE NORMALIDAD"
' Escribimos media y desviación típca errores
Cells(iIni0 - 2, 9) = ChrW(949) & ChrW(772)   ' carácter épsilon + carácter macrón combinable (épsilon media)
Cells(iIni0 - 2, 10) = MediaRes

Cells(iIni0 - 1, 9) = ChrW(963) & ChrW(949)   ' carácter sigma + épsilon
Cells(iIni0 - 1, 9).Characters(Start:=2, Length:=1).Font.Subscript = True  ' sigma sub-épsilon
Cells(iIni0 - 1, 10) = SigmaRes

' Escribimos estimador chi-cuadrado y valor chi-cuadrado
Cells(iIni0 + nRango + 1 + 2, 9).Value = ChrW(967) & "2" & "0"  ' carácter chi (+ "2" + "0")
Cells(iIni0 + nRango + 1 + 2, 9).Characters(Start:=2, Length:=1).Font.Superscript = True
Cells(iIni0 + nRango + 1 + 2, 9).Characters(Start:=3, Length:=1).Font.Subscript = True   ' chi-cuadrado sub-0
Cells(iIni0 + nRango + 1 + 2, 10) = ChiQ0

Cells(iIni0 + nRango + 1 + 3, 9).Value = ChrW(967) & "2" & "1-" & ChrW(945)   ' carácter chi (+ "2" + "1-" + alfa)
Cells(iIni0 + nRango + 1 + 3, 9).Characters(Start:=2, Length:=1).Font.Superscript = True
Cells(iIni0 + nRango + 1 + 3, 9).Characters(Start:=3, Length:=3).Font.Subscript = True   ' chi-cuadrado sub-(1-alfa)
Cells(iIni0 + nRango + 1 + 3, 10) = ChiQAlfa

If ChiQ0 < ChiQAlfa Then
   Cells(iIni0 + nRango + 1 + 4, 9) = "ACEPTAMOS NORMALIDAD PARA " & ChrW(945) & " = " & Alfa
   With Cells(iIni0 + nRango + 1 + 4, 9).Font
      .Color = -1003520
      .TintAndShade = 0
      .Bold = True
   End With
Else
   Cells(iIni0 + nRango + 1 + 4, 9) = "NO ACEPTAMOS NORMALIDAD PARA " & ChrW(945) & " = " & Alfa
   With Cells(iIni0 + nRango + 1 + 4, 9).Font
      .Color = -16776961
      .TintAndShade = 0
      .Bold = True
   End With
End If

' Realizamos el test de homocedasticidad de Breusch-Pagan
Call Test_Breusch_Pagan(m, N, x(), XtXiXt(), Res2(), R2_BP, _
                        ChQ0_BP, ChQAlfa_BP, ChAlfaBP, Alfa, _
                        aBP(), ResBP(), Res2BP())

' Escribimos título del apartado
iIni = iIni0 + nRango + 1 + 6
Cells(iIni, 9) = "TEST DE HOMOCEDASTICIDAD (BREUSCH-PAGAN)"
Cells(iIni + 1, 9) = "R2 (B_P)": Cells(iIni + 1, 10) = R2_BP
Cells(iIni + 1, 9).Characters(Start:=2, Length:=1).Font.Superscript = True

Cells(iIni + 2, 9).Value = ChrW(967) & "2" & "0" & " (B_P)"  ' carácter chi (+ "2" + "0" + " (B_P)")
Cells(iIni + 2, 9).Characters(Start:=2, Length:=1).Font.Superscript = True
Cells(iIni + 2, 9).Characters(Start:=3, Length:=1).Font.Subscript = True   ' chi-cuadrado sub-0 (B_P)
Cells(iIni + 2, 10) = ChQ0_BP

Cells(iIni + 3, 9).Value = ChrW(967) & "2" & "1-" & ChrW(945) & " (n)"  ' carácter chi (+ "2" + "1-" + alfa + " (n)")
Cells(iIni + 3, 9).Characters(Start:=2, Length:=1).Font.Superscript = True
Cells(iIni + 3, 9).Characters(Start:=3, Length:=3).Font.Subscript = True   ' chi-cuadrado sub-(1-alfa) (n)
Cells(iIni + 3, 10) = ChAlfaBP

If ChQ0_BP < ChAlfaBP Then
   Cells(iIni + 4, 9) = "ACEPTAMOS HOMOCEDASTICIDAD PARA " & ChrW(945) & " = " & Alfa
   With Cells(iIni + 4, 9).Font
      .Color = -1003520
      .TintAndShade = 0
      .Bold = True
   End With
Else
   Cells(iIni + 4, 9) = "NO ACEPTAMOS HOMOCEDASTICIDAD PARA " & ChrW(945) & " = " & Alfa
   With Cells(iIni + 4, 9).Font
      .Color = -16776961
      .TintAndShade = 0
      .Bold = True
   End With
End If
Cells(iIni + 4, 9).HorizontalAlignment = xlLeft

End Sub


' RUTINA Producto

Sub Producto(N As Integer, m As Integer, x() As Double, XXT() As Double)
' Calcula el producto XtX
Dim i As Integer
Dim j As Integer
Dim k As Integer

For i = 1 To N
For j = 1 To N
   XXT(i, j) = 0
   For k = 1 To m
     XXT(i, j) = XXT(i, j) + x(k, i) * x(k, j)
   Next
Next
Next
End Sub


' RUTINA TestNormalidad

Sub TestNormalidad(m As Integer, nRango As Integer, Alfa As Double, _
Res() As Double, MediaRes As Double, SigmaRes As Double, _
ChiQ0 As Double, ChiQAlfa As Double, _
Rango() As Double, Frec() As Double, FrecNormal() As Double)

' Esta rutina realiza el test de normalidad
' Llama a la función FD_Normal_MS
' Llama a la función F_Chi_Cuadrado_Inv

Dim i As Integer, j As Integer

' Calculamos las frecuencias relativas acumuladas
For i = 1 To m
   For j = 1 To nRango
    If Res(i) <= Rango(j) And Res(i) < Rango(j + 1) Then Frec(j) = Frec(j) + 1
   Next
Next
Frec(nRango + 1) = m

For i = 1 To nRango + 1
   Frec(i) = Frec(i) / m
Next

' Calculamos la media
MediaRes = 0
For i = 1 To m
   MediaRes = MediaRes + Res(i)
Next
MediaRes = MediaRes / m

' Calculamos la desviación estándar
SigmaRes = 0
For i = 1 To m
   SigmaRes = SigmaRes + (Res(i) - MediaRes) ^ 2
Next
SigmaRes = Sqr(SigmaRes / (m - 1))

For i = 1 To nRango + 1
   FrecNormal(i) = FD_Normal_MS(Rango(i), MediaRes, SigmaRes, 2)
Next

ChiQ0 = 0
For i = 1 To nRango + 1
    ChiQ0 = ChiQ0 + (Frec(i) - FrecNormal(i)) ^ 2 / FrecNormal(i)
Next
ChiQ0 = ChiQ0 * m
ChiQAlfa = F_Chi_Cuadrado_Inv(1 - Alfa, nRango + 1 - 2 - 1)

End Sub


' RUTINA Test_Breusch_Pagan

Sub Test_Breusch_Pagan(m As Integer, N As Integer, _
     x() As Double, XtXiXt() As Double, _
     Res2() As Double, R2_BP As Double, _
     ChQ0_BP As Double, ChQAlfa_BP As Double, ChAlfaBP As Double, Alfa As Double, _
     aBP() As Double, ResBP() As Double, Res2BP() As Double)

' Esta rutina realiza el test de homocedasticidad Breusch-Pagan
' Llama a la función FD_Chi_Cuadrado
' Llama a la función F_Chi_Cuadrado_Inv

Dim yy() As Double, Res2M As Double
Dim i As Integer, j As Integer
Dim SCR As Double, SCE As Double, SCT As Double
ReDim yy(1 To m)

' Resolvemos el sistema multiplicando XtXiXt(n+1 x M) por Res2(M x 1)
For i = 1 To N + 1
   aBP(i) = 0
   For j = 1 To m
      aBP(i) = aBP(i) + XtXiXt(i, j) * Res2(j)
   Next j
Next i

' Calculamos la media de los nuevos residuos al cuadrado
Res2M = 0
For i = 1 To m
   Res2M = Res2M + Res2(i)
Next
Res2M = Res2M / m

'  Variaciones
For i = 1 To m
   yy(i) = 0
   For j = 1 To N + 1
      yy(i) = yy(i) + x(i, j) * aBP(j)
   Next j
   ResBP(i) = Res2(i) - yy(i)
   Res2BP(i) = ResBP(i) * ResBP(i)
   SCR = SCR + Res2BP(i)
   SCE = SCE + (yy(i) - Res2M) ^ 2
Next i

SCT = SCE + SCR
R2_BP = SCE / SCT
ChQ0_BP = m * R2_BP
ChQAlfa_BP = FD_Chi_Cuadrado(ChQ0_BP, N)
ChAlfaBP = F_Chi_Cuadrado_Inv(1 - Alfa, N)

End Sub


' FUNCIÓN fDens_SUniforme

Public Function fDens_SUniforme(x As Double, a As Double, b As Double) As Variant
' Función de densidad de la suma de dos distribuciones uniformes

Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double

If b < a Then
    fDens_SUniforme = "b debe ser mayor que a"
    Exit Function
End If

x1 = 0
x2 = a
x3 = b
x4 = a + b

If x <= x1 Or x >= x4 Then
   fDens_SUniforme = 0
   Exit Function
End If

If x > x1 And x <= x2 Then
   fDens_SUniforme = x / a / b
   Exit Function
End If

If x > x2 And x <= x3 Then
   fDens_SUniforme = 1 / b
   Exit Function
End If

If x > x3 And x <= x4 Then
   fDens_SUniforme = (a + b - x) / a / b
   Exit Function
End If

End Function


' FUNCIÓN FDist_SUniforme

Public Function FDist_SUniforme(x As Double, a As Double, b As Double) As Variant
' Función de distribución de la suma de dos distribuciones uniformes

Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double, Eps As Double

Eps = 0.0000001

If b < a Then
    FDist_SUniforme = "b debe ser mayor que a"
    Exit Function
End If

x1 = 0
x2 = a
x3 = b
x4 = a + b

If Abs(x - x4) < Eps Then
   FDist_SUniforme = 1
   Exit Function
End If

If x <= x1 Or x > x4 Then
   FDist_SUniforme = 0
   Exit Function
End If

If x > x1 And x <= x2 Then
   FDist_SUniforme = x * x / 2 / a / b
   Exit Function
End If

If x > x2 And x <= x3 Then
   FDist_SUniforme = (2 * x - a) / 2 / b
   Exit Function
End If

If x > x3 And x <= x4 Then
   FDist_SUniforme = a * a - 2 * a * x + (b - x) ^ 2
   FDist_SUniforme = -FDist_SUniforme / 2 / a / b
   Exit Function
End If

End Function


' FUNCIÓN fDens_PUniforme

Public Function fDens_PUniforme(x As Double, a As Double, b As Double) As Variant
' Función de densidad del producto de dos distribuciones uniformes

Dim x1 As Double, x4 As Double, Eps As Double

Eps = 0.0000001

x1 = 0
x4 = a * b

If x >= x4 Then
   fDens_PUniforme = 0
   Exit Function
End If

If x < Eps Then
   fDens_PUniforme = ChrW(8734)
   Exit Function
End If

fDens_PUniforme = Log(a * b / x) / a / b

End Function


' FUNCIÓN FDist_PUniforme

Public Function FDist_PUniforme(x As Double, a As Double, b As Double) As Variant
' Función de distribución del producto de dos distribuciones uniformes

Dim x1 As Double, x4 As Double, Eps As Double

Eps = 0.0000001

x1 = 0
x4 = a * b

If Abs(x - x4) < Eps Then
   FDist_PUniforme = 1
   Exit Function
End If

If x < Eps Then
   FDist_PUniforme = 0
   Exit Function
End If

FDist_PUniforme = x * (Log(a * b / x) + 1) / a / b

End Function


' FUNCIÓN fDens_CUniforme

Public Function fDens_CUniforme(x As Double, a As Double, b As Double) As Variant
' Función de densidad del cociente de dos distribuciones uniformes

Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double

If b < a Then
    fDens_CUniforme = "b debe ser mayor que a"
    Exit Function
End If

x1 = 1 / (a + 1)
x2 = 1
x3 = (1 + b) / (1 + a)
x4 = 1 + b

If x <= x1 Or x >= x4 Then
   fDens_CUniforme = 0
   Exit Function
End If

If x > x1 And x <= x2 Then
   fDens_CUniforme = (a + 1) ^ 2 - 1 / x ^ 2
   fDens_CUniforme = fDens_CUniforme / 2 / a / b
   Exit Function
End If

If x > x2 And x <= x3 Then
   fDens_CUniforme = (a + 2) / 2 / b
   Exit Function
End If

If x > x3 And x <= x4 Then
   fDens_CUniforme = ((b + 1) / x) ^ 2 - 1
   fDens_CUniforme = fDens_CUniforme / 2 / a / b
   Exit Function
End If

End Function


' FUNCIÓN FDist_CUniforme

Public Function FDist_CUniforme(x As Double, a As Double, b As Double, Optional N = 100) As Variant
' Función de distribución del cociente de dos distribuciones uniformes
' Llama a la función fDens_CUniforme

Dim x1 As Double, xx As Double, Dx As Double, Eps As Double

Dim i As Integer

Eps = 0.0000001

x1 = 1 / (a + 1)
Dx = (x - x1) / N

If Abs(x - x1) <= Eps Then
   FDist_CUniforme = 0
   Exit Function
End If

xx = x1
FDist_CUniforme = 0
For i = 2 To N + 1
   xx = xx + Dx
   FDist_CUniforme = FDist_CUniforme + fDens_CUniforme(xx, a, b)
Next

FDist_CUniforme = FDist_CUniforme * Dx

End Function




