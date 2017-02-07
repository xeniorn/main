VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSeriesFormatSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LineColor As Long
Public LineWeight As Long
Public MarkerColor As Long
Public MarkerSize As Long
Public MarkerType As Long
Public LineVisible As Boolean
Public LineSmooth As Boolean
Public Transparency As Double


Private Sub Class_Initialize()

    LineColor = RGB(0, 0, 0)
    LineWeight = 2
    MarkerColor = RGB(0, 0, 0)
    MarkerSize = 5 'standard for new graphs
    'MarkerType = XlMarkerStyle.xlMarkerStyleCircle
    MarkerType = XlMarkerStyle.xlMarkerStyleNone
    'MarkerType = xlmarkerstyle.xlMarkerStyleAutomatic
    LineVisible = msoTrue
    LineSmooth = msoTrue
    Transparency = 1 '100 % opaque

End Sub


