VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGraphAnnotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2016-05-09, for this and that purpose
'Last update 2016-05-24
'====================================================================================================
Option Explicit

'TODO: add AxisTypeInfo for X axis!!! 2016-06-10

Const conClassName As String = "clsGraphAnnotation"

'[Var Declaration]
Private pXDataArray() As Double
Private pAnnotationArray() As String
Private pAnnotationType As String

'[Object Initialization]

'[Object Dereferencing]

'[Property Gets and Lets and Sets]
Public Property Get AnnotationType() As String
    AnnotationType = pAnnotationType
End Property

Public Property Get XDataArray() As Double()
    XDataArray = pXDataArray
End Property

Public Property Get XData(ByVal i As Long) As Double
    XData = pXDataArray(i)
End Property

Public Property Let XDataArray(inp_XData() As Double)
    pXDataArray = inp_XData
End Property

Public Property Get Annotation(ByVal i As Long) As String
    Annotation = pAnnotationArray(i)
End Property

Public Property Get AnnotationArray() As String()
    AnnotationArray = pAnnotationArray
End Property

Public Property Let AnnotationArray(inp_YData() As String)
    pAnnotationArray = inp_YData
End Property

Public Property Get Count() As Long
    Count = 1 + UBound(pAnnotationArray) - LBound(pAnnotationArray)
End Property

Public Sub SetType(ByVal inp_AnnotationType As String)
    pAnnotationType = inp_AnnotationType
End Sub

'****************************************************************************************************
Public Sub ShiftX(ByVal ShiftAmount As Double)

    Dim i As Long
    
    Dim tempArray() As Double
    
    tempArray = XDataArray
    
    For i = 1 To UBound(XDataArray)
        tempArray(i) = tempArray(i) + ShiftAmount
    Next i
    
    XDataArray = tempArray
    
End Sub

Private Sub ErrorReport(Optional ErrorNumber As Long = 0, Optional ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub


