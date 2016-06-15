VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGraphData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2016-05-07, for this and that purpose
'Last update 2016-05-23
'====================================================================================================
Option Explicit

Const conClassName As String = "clsGraphData"

'[Var Declaration]
Private pXAxisInfo As AxisDataType
Private pYAxisInfo As AxisDataType
Private pXDataArray() As Double
Private pYDataArray() As Double

'[Object Initialization]

'[Object Dereferencing]

'[Property Gets and Lets and Sets]
Public Property Get XAxisLabel() As String
    XAxisLabel = pXAxisInfo.Label
End Property

Public Property Get YAxisLabel() As String
    YAxisLabel = pYAxisInfo.Label
End Property

Public Property Get XAxisUnit() As String
    XAxisUnit = pXAxisInfo.Unit
End Property

Public Property Get YAxisUnit() As String
    YAxisUnit = pYAxisInfo.Unit
End Property


Public Property Get XAxisInfo() As AxisDataType
    XAxisInfo = pXAxisInfo
End Property

Public Property Let XAxisInfo(inp_XAxis As AxisDataType)
    pXAxisInfo = inp_XAxis
End Property

Public Property Get YAxisInfo() As AxisDataType
    YAxisInfo = pYAxisInfo
End Property

Public Property Let YAxisInfo(inp_YAxis As AxisDataType)
    pYAxisInfo = inp_YAxis
End Property


Public Property Get XData(ByVal i As Long) As Double
    XData = pXDataArray(i)
End Property

Public Property Let XData(ByVal i As Long, C As Double)
    pXDataArray(i) = C
End Property


Public Property Get YData(ByVal i As Long) As Double
    YData = pYDataArray(i)
End Property

Public Property Let YData(ByVal i As Long, C As Double)
    pYDataArray(i) = C
End Property


Public Property Get XDataArray() As Double()
    XDataArray = pXDataArray
End Property

Public Property Let XDataArray(inp_XDataArray() As Double)
    pXDataArray = inp_XDataArray
End Property

Public Property Get YDataArray() As Double()
    YDataArray = pYDataArray
End Property

Public Property Let YDataArray(inp_YDataArray() As Double)
    pYDataArray = inp_YDataArray
End Property

Public Property Get Count() As Long
    Count = 1 + UBound(pXDataArray) - LBound(pXDataArray)
End Property

Private Sub ErrorReport(Optional ErrorNumber As Long = 0, Optional ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub

Private Sub Class_Initialize()

End Sub
