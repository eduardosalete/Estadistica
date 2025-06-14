
' FUNCIÓN F_EN

Public Function F_En(x As Double, N As Integer, Optional Nsumandos As Integer = 100) As Variant
' Llama a la función F_E1

Dim Eps As Double, Gamma As Double, F_En1 As Double, F_En2 As Double
Dim Sumando As Double, Nm1f As Double, Factor As Double
Dim i As Integer, j As Integer

Gamma = 0.577215664901532    ' Constante de Euler Mascheroni
Eps = 0.0000001              ' Margen alrededor del 0

If N <= 1 And Abs(x - 0) < Eps Then
   ' Punto singular
   F_En = "+" & " " & ChrW(8734)
   Exit Function
End If

If N > 1 And Abs(x - 0) < Eps Then
   F_En = 1 / (N - 1)
   Exit Function
End If

' Casos N=0, 1 ó 2
If N = 0 Then
   F_En = Exp(-x) / x
   Exit Function
ElseIf N = 1 Then
   F_En = F_E1(x, Nsumandos)
   Exit Function
ElseIf N = 2 Then
   F_En = -x * F_E1(x, Nsumandos) + Exp(-x)
   Exit Function
End If

' Casos N>2
F_En1 = F_E1(x, Nsumandos): Nm1f = 1
For i = 1 To N - 1
    F_En1 = F_En1 * (-x)
    Nm1f = Nm1f * i
Next

F_En2 = 0
For i = 0 To N - 2
    Factor = 1
    For j = 1 To N - i - 2
        Factor = Factor * j
    Next
    Sumando = Factor * (-x) ^ i
    F_En2 = F_En2 + Sumando
Next
F_En2 = Exp(-x) * F_En2

F_En = (F_En1 + F_En2) / Nm1f

End Function


' FUNCIÓN F_E1

Public Function F_E1(x As Double, Optional Nsumandos As Integer = 100) As Variant
' Llama a la función F_Ei

Dim Eps As Double

Eps = 0.0000001              ' Margen alrededor del 0

If Abs(x - 0) < Eps Then
   ' Punto singular
   F_E1 = "+" & " " & ChrW(8734)
   Exit Function
End If

F_E1 = -F_Ei(-x, Nsumandos)

End Function


' FUNCIÓN F_Ei

Public Function F_Ei(x As Double, Optional Nsumandos As Integer = 100) As Variant
' Llama a la función F_li

Dim xx As Double

xx = Exp(x)

F_Ei = F_li(xx, Nsumandos)

End Function


' FUNCIÓN F_li

Public Function F_li(x As Double, Optional Nsumandos As Integer = 100) As Variant
' Función "logaritmo integral"
' Calculada según el desarrollo de Ramanujan
' x es la variable y NSumandos el nº de sumandos a emplear.
' Llama a la función Aux_li

Dim Gamma As Double, i As Integer, Signo As Integer
Dim Lnx As Double, LnxN As Double, Factor As Double
Dim Eps As Double, Dx As Double, x1 As Double, Sumando As Double

Gamma = 0.577215664901532    ' Constante de Euler Mascheroni
Eps = 0.0000001              ' Margen alrededor del 1

If Abs(x - 0) < Eps Then
   F_li = 0
   Exit Function
End If

If Abs(x - 1) < Eps Then
   ' Punto singular
   F_li = "-" & " " & ChrW(8734)
   Exit Function
End If

If x < 1 Then
   ' Calculamos la función para x<1
   x1 = 1 / x
   Lnx = Log(x1)
   F_li = Gamma + Log(Lnx) - Lnx
   Signo = -1
   Sumando = Lnx
   For i = 2 To Nsumandos
       Signo = -Signo
       Sumando = Sumando * Lnx * (i - 1) / i / i
       F_li = F_li + Signo * Sumando
   Next
   Exit Function
End If

F_li = 0: Signo = -1: Lnx = Log(x): Factor = 2
For i = 1 To Nsumandos
  Signo = -Signo
  ' LnxN = LnxN * Lnx
  Factor = Factor * (Lnx / i / 2)
  F_li = F_li + Signo * Factor * Aux_li(i)
Next

F_li = Gamma + Log(Lnx) + Sqr(x) * F_li

End Function


' FUNCIÓN Aux_li

Public Function Aux_li(N As Integer) As Double
' Función auxiliar para la serie de Ramanujan
Dim i As Long, NN As Long

NN = (N - 1) \ 2
Aux_li = 0
For i = 0 To NN
  Aux_li = Aux_li + 1 / (2 * i + 1)
Next

End Function


