VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGraphRegions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2016-05-13, for this and that purpose
'Last update 2016-05-24
'====================================================================================================
Option Explicit

'Has a 2D array to be filled in the format (i,1) = start of region (i,2) = end of region that is annotated

Const conClassName As String = "clsGraphRegions"

'[Var Declaration]
Private pXStartEndArray() As Double
Private pAnnotation() As String
Private pAnnotationType As String


'[Object Initialization]

'[Object Dereferencing]

'[Property Gets and Lets and Sets]
Public Property Get AnnotationType() As String
    AnnotationType = pAnnotationType
End Property

Public Property Get Xstart(ByVal i As Long) As Double
    Xstart = pXStartEndArray(i, 1)
End Property

Public Property Get Xend(ByVal i As Long) As Double
    Xend = pXStartEndArray(i, 2)
End Property

Public Property Get XStartEndArray() As Double()
    XStartEndArray = pXStartEndArray
End Property

Public Property Let XStartEndArray(ByRef inp_XData() As Double)
    pXStartEndArray = inp_XData
End Property

Public Property Get Annotation(ByVal i As Long) As String
    Annotation = pAnnotation(i)
End Property

Private Property Let Annotation(ByVal i As Long, ByVal S As String)
    pAnnotation(i) = Annotation
End Property

Public Property Get AnnotationArray() As String()
    AnnotationArray = pAnnotation
End Property

Public Property Let AnnotationArray(ByRef inp_YData() As String)
    pAnnotation = inp_YData
End Property

Public Sub SetType(ByVal inp_AnnotationType As String)
    pAnnotationType = inp_AnnotationType
End Sub

Public Property Get Count() As Long
    Count = 1 + UBound(pXStartEndArray) - LBound(pXStartEndArray)
End Property

'****************************************************************************************************
Public Sub ShiftX(ByVal ShiftAmount As Double)

    Dim i As Long
    
    Dim tempArray() As Double
    
    tempArray = XStartEndArray
    
    For i = 1 To UBound(XStartEndArray, 1)
        tempArray(i, 1) = tempArray(i, 1) + ShiftAmount
        tempArray(i, 2) = tempArray(i, 2) + ShiftAmount
    Next i
    
    XStartEndArray = tempArray
    
End Sub

Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub



