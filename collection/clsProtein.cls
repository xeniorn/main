VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProtein"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2017-01-25, for this and that purpose
'Last update 2017-01-25
'====================================================================================================
Option Explicit

Private Const conClassName As String = "clsProtein"

'[Var Declaration]
Private pSequence As String
Private pName As String


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
Public Property Get Sequence() As String
    Sequence = pSequence
End Property

Public Property Let Sequence(ByVal inp_Sequence As String)
    pSequence = inp_Sequence
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal inp_Name As String)
    pName = inp_Name
End Property

Public Function Length() As Long
    Length = Len(Sequence)
End Function




