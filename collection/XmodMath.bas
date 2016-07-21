Attribute VB_Name = "XmodMath"
Option Explicit

'****************************************************************************************************
Sub SwapValue(a As Variant, b As Variant)

'====================================================================================================
'Swaps two values of any type variable
'Juraj Ahel, 2015-04-30, for general purposes
'Last update 2015-04-30
'====================================================================================================

Dim C

C = a
a = b
b = C

End Sub

'****************************************************************************************************
Function RoundToNearestX( _
    ByVal NumberToRound As Double, _
    ByVal RoundingFactor As Double, _
    Optional ByVal RoundDown As Boolean = False) As Double

'====================================================================================================
'Rounds input X to the nearest multiple of input Y
'
'Juraj Ahel, 2015-04-23, for general purposes
'Last update 2015-04-23
'2016-06-13 add RoundDown Flag
'====================================================================================================

If RoundDown Then
    RoundToNearestX = RoundingFactor * Int(NumberToRound / RoundingFactor)
Else
    RoundToNearestX = RoundingFactor * Round(NumberToRound / RoundingFactor)
End If

End Function



'****************************************************************************************************
Function Lg(a As Double) As Double
'====================================================================================================
'Logarithm base 10
'Juraj Ahel, 2015-02-11
'Last update 2015-02-11
'====================================================================================================

Lg = Log(a) / Log(10#)

End Function
'****************************************************************************************************
Function Ln(a As Double) As Double
'====================================================================================================
'Logarithm base e (natural logarithm)
'Juraj Ahel, 2015-02-11
'Last update 2015-02-11
'====================================================================================================

Ln = Log(a)

End Function
'****************************************************************************************************
Function Lb(a As Double) As Double
'====================================================================================================
'Logarithm base 2
'Juraj Ahel, 2015-02-11
'Last update 2015-02-11
'====================================================================================================

Lb = Log(a) / Log(2)

End Function

'****************************************************************************************************
Public Function FMod(a As Double, b As Double) As Double

    FMod = a - Fix(a / b) * b

    'http://en.wikipedia.org/wiki/Machine_epsilon
    'Unfortunately, this function can only be accurate when `a / b` is outside [-2.22E-16,+2.22E-16]
    'Without this correction, FMod(.66, .06) = 5.55111512312578E-17 when it should be 0
    If FMod >= -2 ^ -52 And FMod <= 2 ^ -52 Then '+/- 2.22E-16
        FMod = 0
    End If
    
End Function


