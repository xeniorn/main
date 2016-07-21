VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResearchDatabaseObject"
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

Implements IResearchDatabaseEntry

Private Const conClassName As String = "clsResearchDatabaseObject"

'[Var Declaration]
Private pActualObject As IResearchObject
Public ThingRelationships As clsRelationships
Public ExperimentRelationships As clsRelationships

'[Object Initialization]
Private Sub Class_Initialize()
    
    Set ThingRelationships = New clsRelationships
    Set ExperimentRelationships = New clsRelationships

End Sub

'[Object Dereferencing]
Private Sub Class_Terminate()

    Set pActualObject = Nothing
    Set ThingRelationships = Nothing
    Set ExperimentRelationships = Nothing

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
Public Property Get ActualObject() As IResearchObject
    Set ActualObject = pActualObject
End Property

Public Property Set ActualObject(inp_ActualObject As IResearchObject)
    Set pActualObject = inp_ActualObject
End Property


Private Property Get IResearchDatabaseEntry_EntityPath() As String

End Property

Private Property Get IResearchDatabaseEntry_EntityType() As String

End Property

Private Property Get IResearchDatabaseEntry_ExperimentRelationships() As clsRelationships

End Property

Private Property Get IResearchDatabaseEntry_ThingRelationships() As clsRelationships

End Property
