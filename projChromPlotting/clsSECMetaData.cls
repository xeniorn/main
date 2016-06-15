VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSECMetaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2016-05-09, for this and that purpose
'Last update 2016-05-11
'====================================================================================================
Option Explicit

Private Const conClassName As String = "clsClassTemplate"

'[Var Declaration]
Public ExperimentName As String
Public ExperimentDate As String
Public ExperimentScientist As String
Private pSampleVolume As Double
Public SystemUsed As String
Public ColumnVolume As Double


'[Object Initialization]
Private Sub Class_Initialize()


End Sub

'[Object Dereferencing]
Private Sub Class_Terminate()


End Sub

Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    'If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    If Len(ErrorString) = 0 Then
        Err.Raise vbError + ErrorNumber, conClassName, ErrorString
    Else
        Err.Raise vbError + ErrorNumber, conClassName
    End If

End Sub

'[Property Gets and Lets and Sets]
Public Property Get SampleVolume() As Double
    SampleVolume = pSampleVolume
End Property

Public Property Let SampleVolume(ByVal inp_SampleVolume As Double)
    pSampleVolume = inp_SampleVolume
End Property




