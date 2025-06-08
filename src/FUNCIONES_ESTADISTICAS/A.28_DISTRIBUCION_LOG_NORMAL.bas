
' FUNCIÓN DE DENSIDAD

Public Function D_LogNormal(x As Double, Mu As Double, Sigma As Double) As Variant
' Esta función calcula la función de densidad de la distribución LogNormal(Mu,Sigma)
' Llama a la función D_Normal_MS
Dim y As Double

If Sigma <= 0 Then
   D_LogNormal = "Sigma debe ser >0"
   Exit Function
End If

If x <= 0 Then
   ' Valor negativo
   D_LogNormal = 0
   Exit Function
End If

y = Log(x)
D_LogNormal = 1 / x * D_Normal_MS(y, Mu, Sigma)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_LogNormal(x As Double, Mu As Double, Sigma As Double, Optional Procedimiento As Double = 2) As Variant
' Esta función calcula la función de distribución de la distribución LogNormal(Mu,Sigma)
' Si Procedimiento=1 emplea la relación con la función Gamma Incompleta Inferior
' Si Procedimiento=2 emplea la aproximación de Hastings
' Llama a la función FD_Normal_MS
Dim y As Double

If Sigma <= 0 Then
   FD_LogNormal = "Sigma debe ser >0"
   Exit Function
End If

If x <= 0 Then
   ' Valor negativo
   FD_LogNormal = 0
   Exit Function
End If

y = Log(x)
FD_LogNormal = FD_Normal_MS(y, Mu, Sigma, Procedimiento)

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_LogNormal_Inv(Probabilidad As Double, Mu As Double, Sigma As Double, Optional Procedimiento As Double = 2) As Variant
' Esta función obtiene la inversa de la función de distribución LogNormal(Mu,Sigma)
' Si Procedimiento=1 emplea la relación con la función Gamma Incompleta Inferior
' Si Procedimiento=2 emplea la aproximación de Hastings
' Llama a la función F_Normal_MS_Inv
Dim Eps As Double

Eps = 0.0000001

If Sigma <= 0 Then
   F_LogNormal_Inv = "Sigma debe ser >0"
   Exit Function
End If

If Probabilidad < 0 Or Probabilidad > 1 Then
   F_LogNormal_Inv = "La probabilidad debe estar entre 0 y 1"
   Exit Function
End If

If Probabilidad <= Eps Then
   ' Valor negativo
   F_LogNormal_Inv = 0
   Exit Function
End If

If Abs(Probabilidad - 1) < Eps Then
   F_LogNormal_Inv = ChrW(8734)
   Exit Function
End If

F_LogNormal_Inv = Exp(F_Normal_MS_Inv(Probabilidad, Mu, Sigma, Procedimiento))

End Function


