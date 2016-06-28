VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGenericPhysicalEntity"
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

Implements IPhysicalEntity
Implements IAllowedInDatabase

Private Const conClassName As String = "clsGenericPhysicalEntity"

'[Var Declaration]
Public Name As String

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

Private Property Let IPhysicalEntity_EntityType(RHS As String)

End Property

Private Property Get IPhysicalEntity_EntityType() As String

End Property

Private Property Let IPhysicalEntity_Name(RHS As String)
    Name = RHS
End Property

'[Property Gets and Lets and Sets]


Private Property Get IPhysicalEntity_Name() As String

End Property
