
' FUNCIÓN DE PROBABILIDAD

Public Function p_HiperG(x As Long, N As Long, N_ As Long, p As Double) As Variant
' Función de probabilidad de la distribución Hipergeométrica
' N_ número de opciones
' N  número de ensayos
' p  relación entre las opciones acertadas y el número total de opciones N_
' x  número total de elecciones acertadas
' Llama a la función N_Combinatorio

Dim Nf As Long, Ndf As Long
Dim q As Double, Eps As Double

Eps = 0.0000001

If p < 0 Or p > 1 Then
   p_HiperG = "La prop. de op. acert. ha de estar entre 0 y 1"
   Exit Function
End If
If N_ <= 1 Then
   p_HiperG = "El número de opciones debe ser > 1"
   Exit Function
End If
If N_ < N Then
   p_HiperG = "Debe ser núm. opciones >= núm. ensayos"
   Exit Function
End If

If x < 0 Then
   p_HiperG = 0
   Exit Function
End If

If p < Eps Then
   p_HiperG = 0
   Exit Function
End If

If x > p * N_ Then
   ' No podemos pedir sacar más bolas de éxito que las contenidas en la urna
   p_HiperG = 0
   Exit Function
End If

q = 1 - p
Nf = p * N_
Ndf = N_ - Nf

p_HiperG = N_Combinatorio(Nf, x) / N_Combinatorio(N_, N) * N_Combinatorio(Ndf, N - x)

End Function


' FUNCIÓN DE DISTRIBUCIÓN

Public Function F_HiperG(x As Double, N As Long, N_ As Long, p As Double) As Variant
' Función de distribución de la distribución Hipergeométrica P(Psi<=x)
' Llama a la función p_HiperG

Dim i As Long, ix As Long
Dim Eps As Double

Eps = 0.0000001

If p < 0 Or p > 1 Then
   F_HiperG = "La prop. de op. acert. ha de estar entre 0 y 1"
   Exit Function
End If
If N_ <= 1 Then
   F_HiperG = "El número de opciones debe ser > 1"
   Exit Function
End If
If N_ < N Then
   F_HiperG = "Debe ser núm. opciones >= núm. ensayos"
   Exit Function
End If

If x < 0 Then
   F_HiperG = 0
   Exit Function
End If

If p < Eps Then
   F_HiperG = 0
   Exit Function
End If

If Abs(p - 1) < Eps Then
   F_HiperG = 1
   Exit Function
End If

ix = Fix(x)
F_HiperG = 0
For i = 0 To ix
    F_HiperG = F_HiperG + p_HiperG(i, N, N_, p)
Next i

End Function


' FUNCIÓN Hiperg_Mu

Public Function Hiperg_Mu(N As Long, N_ As Long, p As Double) As Double
' Calcula la media de la distribución Hipergeométrica

If p < 0 Or p > 1 Then
   Hiperg_Mu = "La prop. de op. acert. ha de estar entre 0 y 1"
   Exit Function
End If
If N_ <= 1 Then
   Hiperg_Mu = "El número de opciones debe ser > 1"
   Exit Function
End If
If N_ < N Then
   Hiperg_Mu = "Debe ser núm. opciones >= núm. ensayos"
   Exit Function
End If

Hiperg_Mu = N * p

End Function


' FUNCIÓN Hiperg_M

Public Function Hiperg_M(N As Long, N_ As Long, p As Double) As Long
' Calcula la moda de la distribución Hipergeométrica

If p < 0 Or p > 1 Then
   Hiperg_M = "La prop. de op. acert. ha de estar entre 0 y 1"
   Exit Function
End If
If N_ <= 1 Then
   Hiperg_M = "El número de opciones debe ser > 1"
   Exit Function
End If
If N_ < N Then
   Hiperg_M = "Debe ser núm. opciones >= núm. ensayos"
   Exit Function
End If

Hiperg_M = Fix((N + 1) / (N_ + 2) * (p * N_ + 1))

End Function


' FUNCIÓN Hiperg_S

Public Function Hiperg_S(N As Long, N_ As Long, p As Double) As Double
' Calcula la desviación típica de la distribución Hipergeométrica
Dim q As Double

If p < 0 Or p > 1 Then
   Hiperg_S = "La prop. de op. acert. ha de estar entre 0 y 1"
   Exit Function
End If
If N_ <= 1 Then
   Hiperg_S = "El número de opciones debe ser > 1"
   Exit Function
End If
If N_ < N Then
   Hiperg_S = "Debe ser núm. opciones >= núm. ensayos"
   Exit Function
End If

q = 1 - p
Hiperg_S = Sqr(N * p * q * (N_ - N) / (N_ - 1))

End Function


' FUNCIÓN Hiperg_Gf

Public Function Hiperg_Gf(N As Long, N_ As Long, p As Double) As Double
' Calcula la asimetría de la distribución Hipergeométrica
' Llama a la función Hiperg_S
Dim q As Double, S As Double

If p < 0 Or p > 1 Then
   Hiperg_Gf = "La prop. de op. acert. ha de estar entre 0 y 1"
   Exit Function
End If
If N_ <= 1 Then
   Hiperg_Gf = "El número de opciones debe ser > 1"
   Exit Function
End If
If N_ < N Then
   Hiperg_Gf = "Debe ser núm. opciones >= núm. ensayos"
   Exit Function
End If

q = 1 - p
S = Hiperg_S(N, N_, p)
Hiperg_Gf = (N_ - 2 * N) * (q - p) / S / (N_ - 2)

End Function


' FUNCIÓN Hiperg_k

Public Function Hiperg_k(N As Long, N_ As Long, p As Double) As Double
' Calcula la curtosis de la distribución Hipergeométrica
Dim q As Double, k As Double

If p < 0 Or p > 1 Then
   Hiperg_k = "La prop. de op. acert. ha de estar entre 0 y 1"
   Exit Function
End If
If N_ <= 1 Then
   Hiperg_k = "El número de opciones debe ser > 1"
   Exit Function
End If
If N_ < N Then
   Hiperg_k = "Debe ser núm. opciones >= núm. ensayos"
   Exit Function
End If

q = 1 - p
k = (N_ - 1) * (N_ * N_ * (1 - 6 * p * q) + N_ * (1 - 6 * N) + 6 * N * N)
k = k / N / p / q / (N_ - N) / (N_ - 2) / (N_ - 3)
Hiperg_k = k + 6 * N_ * N_ / (N_ - 2) / (N_ - 3) - 3

End Function


