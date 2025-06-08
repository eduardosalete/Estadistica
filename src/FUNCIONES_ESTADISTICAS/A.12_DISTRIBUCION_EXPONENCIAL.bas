
' FUNCIÓN DE DENSIDAD

Public Function D_Exponencial(x As Double, Lambda As Double) As Variant
' Esta función calcula la función de densidad de la distribución Exponencial
' Devuelve 0 para x<0

If Lambda <= 0 Then
   D_Exponencial = "Lambda debe ser >0"
   Exit Function
End If

If x < 0 Then
   D_Exponencial = 0
   Exit Function
End If

D_Exponencial = Lambda * Exp(-Lambda * x)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Exponencial(x As Double, Lambda As Double) As Variant
' Esta función calcula la función de distribución de la distribución Exponencial
' Devuelve 0 para x<=0

If Lambda <= 0 Then
   FD_Exponencial = "Lambda debe ser >0"
   Exit Function
End If

If x <= 0 Then
   FD_Exponencial = 0
   Exit Function
End If

FD_Exponencial = 1 - Exp(-Lambda * x)

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Exponencial_Inv(Probabilidad As Double, Lambda As Double) As Variant
' Esta función obtiene la inversa de la función de distribución Exponencial
Dim Eps As Double

Eps = 0.0000001

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_Exponencial_Inv = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If Lambda <= 0 Then
   F_Exponencial_Inv = "Lambda debe ser >0"
   Exit Function
End If

If Probabilidad <= Eps Then
   F_Exponencial_Inv = 0
   Exit Function
End If

If Probabilidad >= 1 - Eps Then
   F_Exponencial_Inv = ChrW(8734)
   Exit Function
End If

F_Exponencial_Inv = -Log(1 - Probabilidad) / Lambda

End Function


