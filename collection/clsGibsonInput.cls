VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGibsonInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2017-01-23, for this and that purpose
'Last update 2017-01-23
'====================================================================================================
Option Explicit

Private Const conClassName As String = "clsGibsonInput"

'[Var Declaration]
Private pSource As clsDNA
Private pInsertBefore As clsDNA
Private pFragment As clsDNA
Private pInsertAfter As clsDNA
Private pForbiddenRegions As String

'[Object Initialization]
Private Sub Class_Initialize()

    Set pSource = New clsDNA
    Set pInsertBefore = New clsDNA
    Set pFragment = New clsDNA
    Set pInsertAfter = New clsDNA

End Sub

'[Object Dereferencing]
Private Sub Class_Terminate()

    Set pSource = Nothing
    Set pInsertBefore = Nothing
    Set pFragment = Nothing
    Set pInsertAfter = Nothing

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
Public Property Get Source() As clsDNA
    Set Source = pSource
End Property

Public Property Set Source(ByVal inp_Source As clsDNA)
    Set pSource = inp_Source
End Property

Public Property Get InsertBefore() As clsDNA
    Set InsertBefore = pInsertBefore
End Property

Public Property Set InsertBefore(ByVal inp_InsertBefore As clsDNA)
    Set pInsertBefore = inp_InsertBefore
End Property

Public Property Get Fragment() As clsDNA
    Set Fragment = pFragment
End Property

Public Property Set Fragment(ByVal inp_Fragment As clsDNA)
    Set pFragment = inp_Fragment
End Property

Public Property Get InsertAfter() As clsDNA
    Set InsertAfter = pInsertAfter
End Property

Public Property Set InsertAfter(ByVal inp_InsertAfter As clsDNA)
    Set pInsertAfter = inp_InsertAfter
End Property

Public Property Get ForbiddenRegions() As String
    ForbiddenRegions = pForbiddenRegions
End Property

Public Property Let ForbiddenRegions(ByVal inp_ForbiddenRegions As String)
    pForbiddenRegions = inp_ForbiddenRegions
End Property




