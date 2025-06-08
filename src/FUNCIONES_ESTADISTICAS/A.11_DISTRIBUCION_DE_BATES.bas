
' FUNCIÓN DE DENSIDAD

Public Function D_Bates(x As Double, n As Long) As Variant
' Esta función calcula la función de densidad de la distribución de Bates
' Llama a la distribución de Irwin-Hall y ésta a su vez
'   Llama a la función auxiliar Mi_Factorial
'   Llama a la función auxiliar D_I_H
'   Llama a la función D_Normal_MS y ésta a D_Normal_01
' Cuando n>=25 utiliza la aproximación en el intervalo [0,n]
'
D_Bates = D_Irwin_Hall(n * x, n)

If IsNumeric(D_Bates) Then
   D_Bates = n * D_Bates
End If

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function FD_Bates(x As Double, n As Long) As Variant
' Esta función calcula la función de distribución de la distribución de Bates
' Llama a la distribución de Irwin-Hall y ésta a su vez
'   Llama a la función auxiliar Mi_Factorial
'   Llama a la función auxiliar D_I_H
'   Llama a la función FD_Normal_MS y ésta a FD_Normal_01_G y FD_Normal_01_H
' Cuando n>=25 utiliza la aproximación normal en el intervalo [0,n]
'
FD_Bates = FD_Irwin_Hall(n * x, n)

End Function


' INVERSA DE LA FUNCIÓN DE DISTRIBUCIÓN

Public Function F_Bates_Inv(Probabilidad As Double, n As Long) As Variant
' Esta función calcula la inversa de la función de distribución de la distribución de Bates
' Llama a la distribución de Irwin-Hall y ésta a su vez
'   Llama a la función Mi_ecuacion_Est

F_Bates_Inv = F_Irwin_Hall_Inv(Probabilidad, n)

If IsNumeric(F_Bates_Inv) Then
   F_Bates_Inv = F_Bates_Inv / n
End If

End Function


' FUNCIÓN F_Bates_Media

Public Function F_Bates_Media(n As Long) As Variant
' Calcula la media de la distribución de Bates

If n <= 0 Then
   F_Bates_Media = "n debe ser >0"
   Exit Function
End If

F_Bates_Media = 0.5

End Function


' FUNCIÓN F_Bates_Moda

Public Function F_Bates_Moda(n As Long) As Variant
' Calcula la moda de la distribución de Bates
If n <= 0 Then
   F_Bates_Moda = "n debe ser >0"
   Exit Function
End If

If n = 1 Then
   F_Bates_Moda = "Cualquier valor entre 0 y 1"
Else
   F_Bates_Moda = 0.5
End If

End Function


' FUNCIÓN F_Bates_Mediana

Public Function F_Bates_Mediana(n As Long) As Variant
' Calcula la mediana de la distribución de Bates
If n <= 0 Then
   F_Bates_Mediana = "n debe ser >0"
   Exit Function
End If

F_Bates_Mediana = 0.5

End Function


' FUNCIÓN F_Bates_DesvTip

Public Function F_Bates_DesvTip(n As Long) As Variant
' Calcula la desviación típica de la distribución de Bates

If n <= 0 Then
   F_Bates_DesvTip = "n debe ser >0"
   Exit Function
End If

F_Bates_DesvTip = Sqr(1 / 12 / n)

End Function


' FUNCIÓN F_Bates_Asimetria

Public Function F_Bates_Asimetria(n As Long) As Variant
' Calcula el coeficiente de asimetría (Fisher) de la distribución de Bates

If n <= 0 Then
   F_Bates_Asimetria = "n debe ser >0"
   Exit Function
End If

F_Bates_Asimetria = 0

End Function


' FUNCIÓN F_Bates_Curtosis

Public Function F_Bates_Curtosis(n As Long) As Variant
' Calcula la curtosis de la distribución de Bates

If n <= 0 Then
   F_Bates_Curtosis = "n debe ser >0"
   Exit Function
End If

F_Bates_Curtosis = 3 - 6 / 5 / n

End Function


