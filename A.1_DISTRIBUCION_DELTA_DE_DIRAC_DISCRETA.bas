
' FUNCIÓN DE PROBABILIDAD
Public Function p_Delta_Discreta(x As Double, a As Double) As Double' Función de probabilidad de la distribución Delta de Dirac discreta' a es el punto donde se concentra la masa probabilísticap_Delta_Discreta = 0If x = a Then    p_Delta_Discreta = 1End IfEnd Function
' FUNCIÓN DE DISTRIBUCIÓN
Public Function F_Delta_Discreta(x As Double, a As Double) As Double' Función de distribución de la distribución Delta de Dirac discretaIf x < a Then    F_Delta_Discreta = 0Else    F_Delta_Discreta = 1End IfEnd Function
