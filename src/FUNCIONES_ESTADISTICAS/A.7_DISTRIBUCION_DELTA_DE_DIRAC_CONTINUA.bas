
' FUNCI�N DE DISTRIBUCI�N
Public Function FD_Delta_Continua(x As Double, a As Double) As Double' Esta funci�n calcula la funci�n de distribuci�n de la distribuci�n' Delta de Dirac continuaIf x < a Then    FD_Delta_Continua = 0Else    FD_Delta_Continua = 1End IfEnd Function
