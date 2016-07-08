VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGenericExperiment"
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

Implements IExperiment
Implements IAllowedInDatabase

Private Const conClassName As String = "clsGenericExperiment"

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

Private Property Let IExperiment_ExpDate(RHS As String)

End Property

Private Property Get IExperiment_ExpDate() As String

End Property

'[Property Gets and Lets and Sets]



Private Property Get IExperiment_ExpName() As String
    IExperiment_ExpName = Name
End Property

Private Property Let IExperiment_ExpName(RHS As String)
    Name = RHS
End Property

Private Property Let IExperiment_ExpType(RHS As String)

End Property

Private Property Get IExperiment_ExpType() As String

End Property

Private Property Let IExperiment_Scientist(RHS As String)

End Property

Private Property Get IExperiment_Scientist() As String

End Property
