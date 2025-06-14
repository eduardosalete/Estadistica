
' DISCRETO

Sub Mi_ConvV()
' Calcula la convolución de dos vectores
' Ambas deben estar definidas en el mismo rango de valores
' j1 y j2 son las columnas de entrada y j3 la de salida (siempre empiezan en la fila 2)
' NMax1 longitud del primer vector
' NMax2 longitud del segundo vector
' NMax3 longitud del vector producto

Dim i As Long, j As Long, k As Long
Dim j1 As Integer, j2 As Integer, j3 As Integer
Dim Vect1(1 To nn) As Double, Vect2(1 To nn) As Double
Dim vect3(1 To nn) As Double, F(1 To nn) As Double
Dim NMax1 As Long, NMax2 As Long, NMax As Long, NNMax As Long

j1 = 2
j2 = 3
j3 = 4
NNMax = 32000

' Leemos el primer vector
NMax1 = 0
For i = 2 To NNMax
  If Cells(i, j1) = "" Or Cells(i, j1) = " " Then Exit For
  NMax1 = NMax1 + 1
  Vect1(NMax1) = Cells(i, j1)
Next

' Leemos el segundo vector
NMax2 = 0
For i = 2 To NNMax
  If Cells(i, j2) = "" Or Cells(i, j2) = " " Then Exit For
  NMax2 = NMax2 + 1
  Vect2(NMax2) = Cells(i, j2)
Next

' Longitd máxima
NMax = NMax1 + NMax2 - 1

' Realizamos la convolución
' Última parte
For i = NMax2 To NMax
    vect3(i) = 0
    j = i
    For k = 1 To NMax2
        vect3(i) = vect3(i) + Vect1(j) * Vect2(k)
        j = j - 1
    Next
Next

' Primera parte
For i = 1 To NMax2
    vect3(i) = 0
    j = i
    k = 1
    Do While j > 0
        vect3(i) = vect3(i) + Vect1(j) * Vect2(k)
        j = j - 1
        k = k + 1
    Loop
    If i = 1 Then
       F(1) = vect3(i)
    Else
       F(i) = F(i - 1) + vect3(i)
    End If
Next

For i = NMax2 To NMax
    F(i) = F(i - 1) + vect3(i)
Next

' Escribimos los resultados
For i = 2 To NMax + 1
   Cells(i, j3) = vect3(i - 1)
   Cells(i, j3 + 1) = F(i - 1)
Next

End Sub


' CONTINUO

Sub Mi_ConvF()
' Calcula la convolución de dos funciones
' Ambas deben estar definidas en el mismo rango de valores
' j1 y j2 son las columnas de entrada y j3 la de salida (siempre empiezan en la fila 2
' NMax1 longitud del primer vector
' NMax2 longitud del segundo vector
' NMax3 longitud del vector producto

Dim i As Long, j As Long, k As Long
Dim j1 As Integer, j2 As Integer, j3 As Integer
Dim Vect1(1 To nn) As Double, Vect2(1 To nn) As Double
Dim Vect3(1 To nn) As Double, Vect3Acum(1 To nn) As Double, F(1 To nn)

Dim NMax1 As Long, NMax2 As Long, NMax As Long, NNMax As Long
Dim xI As Double, dx As Double
Dim Area As Double, Suma As Double

j1 = 2
j2 = 3
j3 = 4
NNMax = 32000

' x inicial y dx
xI = Range("H2")
dx = Range("K2")

' Leemos el primer vector
NMax1 = 0
For i = 2 To NNMax
  If Cells(i, j1) = "" Or Cells(i, j1) = " " Then Exit For
  NMax1 = NMax1 + 1
  Vect1(NMax1) = Cells(i, j1) * dx
Next

' Leemos el segundo vector
NMax2 = 0
For i = 2 To NNMax
  If Cells(i, j2) = "" Or Cells(i, j2) = " " Then Exit For
  NMax2 = NMax2 + 1
  Vect2(NMax2) = Cells(i, j2) * dx
Next


' Reajustamos los extremos
' Dándoles la mita de masa probabilística
Vect1(1) = Vect1(1) / 2
Vect2(1) = Vect2(1) / 2
Vect1(NMax1) = Vect1(NMax1) / 2
Vect2(NMax1) = Vect1(NMax2) / 2

' Longitud máxima
NMax = NMax1 + NMax2 - 1

' Realizamos la convolución
' Última parte
Area = 0
For i = NMax2 To NMax
    Vect3(i) = 0
    j = i
    For k = 1 To NMax2
        Vect3(i) = Vect3(i) + Vect1(j) * Vect2(k)
        j = j - 1
    Next
    Vect3Acum(i) = Vect3Acum(i) + Vect3(i)
    Area = Area + Vect3Acum(i)
Next

' Primera parte
For i = 1 To NMax2
    Vect3(i) = 0
    j = i
    k = 1
    Do While j > 0
        Vect3(i) = Vect3(i) + Vect1(j) * Vect2(k)
        j = j - 1
        k = k + 1
    Loop
    Vect3Acum(i) = Vect3Acum(i) + Vect3(i)
    Area = Area + Vect3Acum(i)
Next

F(1) = Vect3Acum(1)
For i = 2 To NMax
   F(i) = F(i - 1) + Vect3Acum(i - 1)
Next

' Ajustamos para que la suma total sea la unidad
For i = 1 To NMax
    F(i) = F(i) / F(NMax)
    Vect3Acum(i) = Vect3Acum(i) / F(NMax)
Next

' Escribimos los resultados
Suma = 0
For i = 2 To NMax + 1
   Cells(i, j3) = xI + dx * (i - 2)
   Cells(i, j3 + 1) = Vect3(i - 1) / dx
   Cells(i, j3 + 2) = Vect3Acum(i - 1)
   Cells(i, j3 + 3) = F(i - 1)
   Suma = Suma + Vect3Acum(i - 1)
Next

End Sub

APÉNDICE D




