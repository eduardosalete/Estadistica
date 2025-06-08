
' FUNCIÓN DE DENSIDAD

Public Function D_Normal_MS(x As Double, Mu As Double, Sigma As Double) As Variant
' Esta función calcula la función de densidad de la distribución N(Mu,Sigma)
' Llama a la función D_Normal_01

Dim y As Double

If Sigma <= 0 Then
   D_Normal_MS = "Sigma debe ser >0"
   Exit Function
End If

y = (x - Mu) / Sigma
D_Normal_MS = D_Normal_01(y) / Sigma

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Normal_MS(x As Double, Mu As Double, Sigma As Double, Optional Procedimiento As Double = 2) As Variant
' Esta función calcula la función de distribución de la distribución N(Mu,Sigma)
' Si Procedimiento=1 emplea la relación con la función Gamma Incompleta Inferior
' Si Procedimiento=2 emplea la aproximación de Hastings
' Llama a la función FD_Normal_01_G
' Llama a la función FD_Normal_01_H

Dim y As Double

If Sigma <= 0 Then
   FD_Normal_MS = "Sigma debe ser >0"
   Exit Function
End If

y = (x - Mu) / Sigma

If Abs(Procedimiento - 1) < 0.01 Then
   FD_Normal_MS = FD_Normal_01_G(y)
Else
   FD_Normal_MS = FD_Normal_01_H(y)
End If

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Normal_MS_Inv(Probabilidad As Double, Mu As Double, Sigma As Double, Optional Procedimiento As Double = 2) As Variant
' Esta función obtiene la inversa de la función de distribución N(Mu,Sigma)
' Si Procedimiento=1 emplea la relación con la función Gamma Incompleta Inferior
' Si Procedimiento=2 emplea la aproximación de Hastings
' Llama a la función F_Normal_01_Inv

If Sigma <= 0 Then
   F_Normal_MS_Inv = "Sigma debe ser >0"
   Exit Function
End If

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_Normal_MS_Inv = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

F_Normal_MS_Inv = Sigma * F_Normal_01_Inv(Probabilidad, Procedimiento) + Mu

End Function


