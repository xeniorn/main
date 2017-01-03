Attribute VB_Name = "XmodColor"
Option Explicit

Private Const Pi As Double = 3.14159265258979

Sub testColors()


    Dim R As Long, G As Long, b As Long
    Dim H As Double, S As Double, V As Double
    
    
    Dim agh As Range
    Dim i As Long
    Dim j As Long
    
    Set agh = Range("A1").Resize(6, 10)
    
    H = Pi / 3
    S = 1
    
    For j = 0 To 5
        H = j * Pi / 3
        For i = 1 To 10
            'H = H + 0.025
            V = i / 10
            agh.Cells(j + 1, i).Interior.Color = ColorFromHSV(H, S, V)
        Next i
    Next j

End Sub

Function ColorFromHSV(H As Double, S As Double, V As Double) As Long

    Dim R As Long, G As Long, b As Long
    
    HSV2RGB H, S, V, R, G, b
    
    ColorFromHSV = RGB(R, G, b)

End Function

Sub Color2RGB(ByVal InputColor As Long, _
                ByRef R As Long, _
                ByRef G As Long, _
                ByRef b As Long)
    
    Dim sColor As String
        
    R = InputColor Mod 256
    InputColor = InputColor \ 256
    G = InputColor Mod 256
    InputColor = InputColor \ 256
    b = InputColor Mod 256
    
    
End Sub


Private Sub test1()

Dim R As Long, G As Long, b As Long
Dim H As Double, S As Double, V As Double

R = 63
G = 127
b = 127

RGB2HSV R, G, b, H, S, V

Debug.Print ("H: " & H)
Debug.Print ("S: " & S)
Debug.Print ("V: " & V)

End Sub

Private Sub test2()

Dim R As Long, G As Long, b As Long
Dim H As Double, S As Double, V As Double

H = Pi * 3 / 2
S = 1
V = 1

HSV2RGB H, S, V, R, G, b

Debug.Print ("R: " & R)
Debug.Print ("G: " & G)
Debug.Print ("B: " & b)

End Sub




Private Function MaximumValue(ParamArray Inputs() As Variant) As Variant

    Dim InputVarType As VbVarType
    Dim i As Long
    Dim tempmax As Variant
    
    InputVarType = VarType(Inputs(0))
    tempmax = Inputs(0)
    
    'TODO: do a check of all the input elements
    
    If IsNumeric(Inputs(0)) Then
        
        For i = 1 To UBound(Inputs)
            If Inputs(i) > tempmax Then
                tempmax = Inputs(i)
            End If
        Next i
        
    MaximumValue = tempmax
    
    End If
    

End Function

Private Function MinimumValue(ParamArray Inputs() As Variant) As Variant

    Dim InputVarType As VbVarType
    Dim i As Long
    Dim tempmin As Variant
    
    InputVarType = VarType(Inputs(0))
    tempmin = Inputs(0)
    
    'TODO: do a check of all the input elements
    
    If IsNumeric(Inputs(0)) Then
        
        For i = 1 To UBound(Inputs)
            If Inputs(i) < tempmin Then
                tempmin = Inputs(i)
            End If
        Next i
        
    MinimumValue = tempmin
    
    End If
    

End Function

Sub RGB2HSV(ByVal R As Long, _
            ByVal G As Long, _
            ByVal b As Long, _
            ByRef H As Double, _
            ByRef S As Double, _
            ByRef V As Double)
            
'red green blue (0-255) to
'hue saturation value (0-2pi, 0-1, 0-1)
            
    Dim rR As Double, rG As Double, rB As Double
    Dim Cmax As Double, Cmin As Double
    Dim delta As Double
    
    rR = R / 255
    rG = G / 255
    rB = b / 255
    
    Cmax = MaximumValue(rR, rG, rB)
    Cmin = MinimumValue(rR, rG, rB)
    
    delta = Cmax - Cmin
    
    If delta = 0 Then
        H = 0
    Else
        Select Case Cmax
            Case rR
                H = Pi / 3 * (((rG - rB) / delta) Mod 6)
            Case rG
                H = Pi / 3 * (((rB - rR) / delta) + 2)
            Case rB
                H = Pi / 3 * (((rR - rG) / delta) + 4)
            Case Else
                Err.Raise 1
        End Select
    End If
    
    If Cmax = 0 Then
        S = 0
    Else
        S = delta / Cmax
    End If
        
    V = Cmax

End Sub

Sub HSV2RGB(ByVal H As Double, _
            ByVal S As Double, _
            ByVal V As Double, _
            ByRef R As Long, _
            ByRef G As Long, _
            ByRef b As Long)
            
    Dim rR As Double, rG As Double, rB As Double
    Dim rH As Double
    Dim c As Double
    Dim x As Double
    Dim m As Double
    Dim Hdegree As Double
    
    Hdegree = H / Pi * 180
    rH = H / (Pi / 3)
    
    'total color intensity is value reduced by saturation %
    c = V * S
    
    'ratio in the second color
    x = V * S * (1 - Abs(FMod(rH, 2) - 1))
    
    'neutral color intensity = total value - saturated value
    'added to all colors (r,g,b)
    m = V * (1 - S)
    
    
    If rH >= 0 And rH < 1 Then
        rR = c
        rG = x
        rB = 0
    End If
    
    If rH >= 1 And rH < 2 Then
        rR = x
        rG = c
        rB = 0
    End If
    
    If rH >= 2 And rH < 3 Then
        rR = 0
        rG = c
        rB = x
    End If
    
    If rH >= 3 And rH < 4 Then
        rR = 0
        rG = x
        rB = c
    End If
    
    If rH >= 4 And rH < 5 Then
        rR = x
        rG = 0
        rB = c
    End If
    
    If rH >= 5 And rH < 6 Then
        rR = c
        rG = 0
        rB = x
    End If
    
    R = (rR + m) * 255
    G = (rG + m) * 255
    b = (rB + m) * 255

End Sub


