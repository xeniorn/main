Attribute VB_Name = "modMath"
Option Explicit


'****************************************************************************************************
Function RoundToSignificantDigits( _
    ByVal NumberToRound As Double, _
    ByVal SignificantDigits As Long _
    ) As Double

'====================================================================================================
'Rounds input X to Y significant digits
'
'Juraj Ahel, 2017-02-07, for general purposes
'====================================================================================================
    
    If NumberToRound = 0 Then
        RoundToSignificantDigits = 0
    Else
        RoundToSignificantDigits = Round(NumberToRound, SignificantDigits - Int(Lg(NumberToRound) + 1))
    End If

End Function

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

