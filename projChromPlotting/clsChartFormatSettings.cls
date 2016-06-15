VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartFormatSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ChartAreaBgVisible As Boolean
Public PlotAreaBgVisible As Boolean
Public BorderLineStyle As Long
Public ChartType As Long

Private Sub Class_Initialize()

    BorderLineStyle = xlNone
    ChartAreaBgVisible = False
    PlotAreaBgVisible = False
    ChartType = xlXYScatterSmoothNoMarkers
    
End Sub

