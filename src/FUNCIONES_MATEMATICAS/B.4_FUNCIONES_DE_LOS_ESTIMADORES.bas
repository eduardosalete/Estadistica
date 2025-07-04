
' FUNCIÓN Estimador_Media

Public Function Estimador_Media(VRango As Range) As Variant
Dim datos() As Double          ' Datos de entrada
Dim valorM As Double           ' Valor medio calculado de los datos de entrada
Dim n As Long                  ' Número de datos de entrada
Dim i As Long                  ' Índice para bucles

n = VRango.Rows.Count
If (n < 1) Then
   Estimador_Media = "No hay datos"
   Exit Function
End If

ReDim datos(1 To n)

' Calcular la Media de los datos

valorM = 0
For i = 1 To n
    datos(i) = VRango.Cells(i, 1).Value   ' Lectura del vector datos() introducido como rango
    valorM = valorM + datos(i)
Next i

valorM = valorM / n

Estimador_Media = valorM

End Function


' FUNCIÓN Estimador_Cuasivarianza

Public Function Estimador_Cuasivarianza(VRango As Range) As Variant
' Llama a la función Estimador_Media
Dim datos() As Double          ' Datos de entrada
Dim valorM As Double           ' Valor medio calculado de los datos de entrada
Dim CuasiVarianza As Double    ' Cuasivarianza calculada de los datos de entrada
Dim n As Long                  ' Número de datos de entrada
Dim i As Long                  ' Índice para bucles

n = VRango.Rows.Count
If (n <= 1) Then
   Estimador_Cuasivarianza = "Se necesita más de un valor"
   Exit Function
End If

ReDim datos(1 To n)

' Calcular la Cuasivarianza de los datos

valorM = Estimador_Media(VRango)

CuasiVarianza = 0
For i = 1 To n
    datos(i) = VRango.Cells(i, 1).Value   ' Lectura del vector datos() introducido como rango
    CuasiVarianza = CuasiVarianza + (datos(i) - valorM) ^ 2
Next i

CuasiVarianza = CuasiVarianza / (n - 1)

Estimador_Cuasivarianza = CuasiVarianza

End Function


' FUNCIÓN Estimador_Coef_Asimetria

Public Function Estimador_Coef_Asimetria(VRango As Range) As Variant
' Llama a la función Estimador_Media
' Llama a la función Estimador_Cuasivarianza
Dim datos() As Double          ' Datos de entrada
Dim valorM As Double           ' Valor medio calculado de los datos de entrada
Dim Desv_Estandar As Double    ' Desviación Estándar calculada de los datos de entrada
Dim Coef_Asimetria As Double    ' Coeficiente de Asimetría calculado de los datos de entrada
Dim n As Long                  ' Número de datos de entrada
Dim i As Long                  ' Índice para bucles
Dim eps As Double

eps = 0.0000001

n = VRango.Rows.Count
If (n <= 2) Then
   Estimador_Coef_Asimetria = "Se necesitan más de dos valores"
   Exit Function
End If

ReDim datos(1 To n)

' Calcular el Coeficiente de Asimetría de los datos

valorM = Estimador_Media(VRango)
Desv_Estandar = Sqr(Estimador_Cuasivarianza(VRango))

Coef_Asimetria = 0
For i = 1 To n
    datos(i) = VRango.Cells(i, 1).Value   ' Lectura del vector datos() introducido como rango
    Coef_Asimetria = Coef_Asimetria + (datos(i) - valorM) ^ 3
Next i

If Abs(Coef_Asimetria) < eps Then
   Estimador_Coef_Asimetria = 0
   Exit Function
End If

Coef_Asimetria = n * Coef_Asimetria / ((n - 1) * (n - 2) * Desv_Estandar ^ 3)

Estimador_Coef_Asimetria = Coef_Asimetria

End Function


' FUNCIÓN Estimador_Exceso_Curtosis

Public Function Estimador_Exceso_Curtosis(VRango As Range) As Variant
' Llama a la función Estimador_Media
' Llama a la función Estimador_Cuasivarianza
Dim datos() As Double            ' Datos de entrada
Dim valorM As Double             ' Valor medio calculado de los datos de entrada
Dim Desv_Estandar As Double      ' Desviación Estándar calculada de los datos de entrada
Dim Exceso_Curtosis As Double    ' Exceso de Curtosis calculado de los datos de entrada
Dim n As Long                    ' Número de datos de entrada
Dim i As Long                    ' Índice para bucles
Dim eps As Double

eps = 0.0000001

n = VRango.Rows.Count
If (n <= 3) Then
   Estimador_Exceso_Curtosis = "Se necesitan más de tres valores"
   Exit Function
End If

ReDim datos(1 To n)

' Calcular el Exceso de Curtosis de los datos

valorM = Estimador_Media(VRango)
Desv_Estandar = Sqr(Estimador_Cuasivarianza(VRango))

Exceso_Curtosis = 0
For i = 1 To n
    datos(i) = VRango.Cells(i, 1).Value   ' Lectura del vector datos() introducido como rango
    Exceso_Curtosis = Exceso_Curtosis + (datos(i) - valorM) ^ 4
Next i

If Abs(Exceso_Curtosis) < eps And Abs(Desv_Estandar) < eps Then
   Estimador_Exceso_Curtosis = 0
   Exit Function
End If

Exceso_Curtosis = n * (n + 1) * Exceso_Curtosis / ((n - 1) * (n - 2) * (n - 3) * Desv_Estandar ^ 4)
Exceso_Curtosis = Exceso_Curtosis - 3 * (n - 1) ^ 2 / ((n - 2) * (n - 3))

Estimador_Exceso_Curtosis = Exceso_Curtosis

End Function


