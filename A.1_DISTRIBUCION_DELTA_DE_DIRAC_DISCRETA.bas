
' FUNCI�N DE PROBABILIDAD
Public Function p_Delta_Discreta(x As Double, a As Double) As Double' Funci�n de probabilidad de la distribuci�n Delta de Dirac discreta' a es el punto donde se concentra la masa probabil�sticap_Delta_Discreta = 0If x = a Then    p_Delta_Discreta = 1End IfEnd Function
' FUNCI�N DE DISTRIBUCI�N
Public Function F_Delta_Discreta(x As Double, a As Double) As Double' Funci�n de distribuci�n de la distribuci�n Delta de Dirac discretaIf x < a Then    F_Delta_Discreta = 0Else    F_Delta_Discreta = 1End IfEnd Function
