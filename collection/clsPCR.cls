VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPCR"
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

Private Const conClassName As String = "clsPCR"

'[Var Declaration]
Private pSourceDNA As clsDNA
Private pFinalDNA As clsDNA
Private pForwardPrimer As clsDNA
Private pReversePrimer As clsDNA

'[Object Initialization]
Private Sub Class_Initialize()

    Set pSourceDNA = New clsDNA
    Set pFinalDNA = New clsDNA
    Set pForwardPrimer = New clsDNA
    Set pReversePrimer = New clsDNA

End Sub

'[Object Dereferencing]
Private Sub Class_Terminate()

    Set pSourceDNA = Nothing
    Set pFinalDNA = Nothing
    Set pForwardPrimer = Nothing
    Set pReversePrimer = Nothing

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
Public Property Get SourceDNA() As clsDNA
    Set SourceDNA = pSourceDNA
End Property

Public Property Set SourceDNA(ByVal inp_SourceDNA As clsDNA)
    Set pSourceDNA = inp_SourceDNA
End Property

Public Property Get FinalDNA() As clsDNA
    Set FinalDNA = pFinalDNA
End Property

Public Property Set FinalDNA(ByVal inp_FinalDNA As clsDNA)
    Set pFinalDNA = inp_FinalDNA
End Property

Public Property Get ForwardPrimer() As clsDNA
    Set ForwardPrimer = pForwardPrimer
End Property

Public Property Set ForwardPrimer(ByVal inp_ForwardPrimer As clsDNA)
    Set pForwardPrimer = inp_ForwardPrimer
End Property

Public Property Get ReversePrimer() As clsDNA
    Set ReversePrimer = pReversePrimer
End Property

Public Property Set ReversePrimer(ByVal inp_ReversePrimer As clsDNA)
    Set pReversePrimer = inp_ReversePrimer
End Property




