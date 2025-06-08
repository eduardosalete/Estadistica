
' FUNCIÓN DE DENSIDAD

Public Function D_Uniforme(x As Double, a As Double, b As Double) As Variant
' Esta función calcula la función de densidad de la distribución Uniforme continua
' a valor inferior, b valor superior
If b <= a Then
    D_Uniforme = "b debe ser mayor que a"
End If

If x < a Then
    D_Uniforme = 0
ElseIf x >= a And x <= b Then
    D_Uniforme = 1 / (b - a)
Else
    D_Uniforme = 0
End If

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Uniforme(x As Double, a As Double, b As Double) As Variant
' Esta función calcula la función de distribución de la distribución Uniforme continua
' a valor inferior, b valor superior
If b <= a Then
    FD_Uniforme = "b debe ser mayor que a"
End If

If x <= a Then
    FD_Uniforme = 0
ElseIf x > a And x < b Then
    FD_Uniforme = (x - a) / (b - a)
Else
    FD_Uniforme = 1
End If

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Uniforme_Inv(Probabilidad As Double, a As Double, b As Double) As Variant
' Esta función calcula la inversa de la función de distribución de la distribución Uniforme continua
' a valor inferior, b valor superior
If b <= a Then
    F_Uniforme_Inv = "b debe ser mayor que a"
End If

If Probabilidad < 0 Or Probabilidad > 1 Then
    F_Uniforme_Inv = "La probabilidad debe estar entre 0 y 1"
End If
F_Uniforme_Inv = (b - a) * Probabilidad + a

End Function


