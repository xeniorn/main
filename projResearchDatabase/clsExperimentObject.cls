VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExperimentObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2016-06-28, for this and that purpose
'Last update 2016-06-28
'====================================================================================================
Option Explicit

Implements IResearchObject

Private Const conClassName As String = "clsExperimentObject"

'[Var Declaration]
Private pActualExperiment As IExperiment

'[Object Initialization]
Private Sub Class_Initialize()

    Set pActualExperiment = New IExperiment

End Sub

'[Object Dereferencing]
Private Sub Class_Terminate()

    Set pActualExperiment = Nothing

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
Public Property Get ActualExperiment() As IExperiment
    Set ActualExperiment = pActualExperiment
End Property

Public Property Set ActualExperiment(inp_ActualExperiment As IExperiment)
    Set pActualExperiment = inp_ActualExperiment
End Property

Private Sub IResearchObject_Add(ObjectToAdd As Object)
    Set pActualExperiment = ObjectToAdd
End Sub

Private Property Let IResearchObject_EntityType(RHS As String)

End Property

Private Property Get IResearchObject_EntityType() As String

End Property

Private Property Let IResearchObject_FilePath(RHS As String)

End Property

Private Property Get IResearchObject_FilePath() As String

End Property

Private Property Let IResearchObject_Name(RHS As String)
    ActualExperiment.ExpName = RHS
End Property

Private Property Get IResearchObject_Name() As String

End Property
