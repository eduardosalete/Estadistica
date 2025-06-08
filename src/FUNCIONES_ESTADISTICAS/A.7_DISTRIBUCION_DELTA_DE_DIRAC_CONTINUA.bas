
' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Delta_Continua(x As Double, a As Double) As Double
' Esta función calcula la función de distribución de la distribución
' Delta de Dirac continua

If x < a Then
    FD_Delta_Continua = 0
Else
    FD_Delta_Continua = 1
End If

End Function


