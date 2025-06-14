
' FUNCIÓN DE PROBABILIDAD

Public Function p_Pascal(x As Integer, p As Double) As Variant
' Función de probabilidad de la distribución de Pascal
' x es el número de fallos
' p la probabilidad de acertar en un ensayo

Dim q As Double, Eps As Double

Eps = 0.0000001

If p < 0 Or p > 1 Then
   p_Pascal = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If x < 0 Then
   p_Pascal = 0
   Exit Function
End If

If p < Eps Then
   p_Pascal = 0
   Exit Function
End If

q = 1 - p
p_Pascal = p * q ^ x

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Pascal(x As Double, p As Double) As Variant
' Función de distribución de la distribución de Pascal P(Psi<=x)
' Llama a la función p_Pascal

Dim i As Integer, ix As Integer
Dim Eps As Double

Eps = 0.0000001

If p < 0 Or p > 1 Then
   F_Pascal = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If x < 0 Then
   F_Pascal = 0
   Exit Function
End If

If p < Eps Then
   F_Pascal = 0
   Exit Function
End If

ix = Fix(x)
F_Pascal = 0
For i = 0 To ix
    F_Pascal = F_Pascal + p_Pascal(i, p)
Next i

End Function


