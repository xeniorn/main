VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRelationships"
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

Private Const conClassName As String = "clsRelationships"

'[Var Declaration]
Private pParents As VBA.Collection
Private pSiblings As VBA.Collection
Private pChildren As VBA.Collection

'[Object Initialization]
Private Sub Class_Initialize()

    Set pParents = New VBA.Collection
    Set pSiblings = New VBA.Collection
    Set pChildren = New VBA.Collection

End Sub

'[Object Dereferencing]
Private Sub Class_Terminate()

    Set pParents = Nothing
    Set pSiblings = Nothing
    Set pChildren = Nothing

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
Public Property Get Parents() As VBA.Collection
    Set Parents = pParents
End Property

Public Property Set Parents(inp_Parents As VBA.Collection)
    Set pParents = inp_Parents
End Property

Public Property Get Siblings() As VBA.Collection
    Set Siblings = pSiblings
End Property

Public Property Set Siblings(inp_Siblings As VBA.Collection)
    Set pSiblings = inp_Siblings
End Property

Public Property Get Children() As VBA.Collection
    Set Children = pChildren
End Property

Public Property Set Children(inp_Children As VBA.Collection)
    Set pChildren = inp_Children
End Property


