VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAxisFormatSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public AxisWeight As Long
Public AxisColor As Long
Public MajorTickType As Long
Public MinorTickType As Long
Public TickLabelSize As Long
Public TitleFontSize As Long
Public TitleColor As Long
Public HasMajorGridlines As Boolean
Public HasMinorGridlines As Boolean
Public HasTitle As Boolean

Private Sub Class_Initialize()

        AxisWeight = xlMedium
        AxisColor = RGB(147, 149, 152) 'from Illustrator testing
        MajorTickType = xlTickMarkOutside
        MinorTickType = xlTickMarkOutside
        TickLabelSize = 25
        TitleFontSize = 25
        TitleColor = AxisColor
        HasMajorGridlines = False
        HasMinorGridlines = False
        HasTitle = True

End Sub

