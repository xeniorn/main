VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSeqReaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPrimer As clsPrimer
Private pTemplate As clsDNA

Private Sub Class_Initialize()

    If pPrimer Is Nothing Then
        Set pPrimer = New clsPrimer
    End If
    
    If pTemplate Is Nothing Then
        Set pTemplate = New clsDNA
    End If

End Sub

Private Sub Class_Terminate()

    Set pPrimer = Nothing
    Set pTemplate = Nothing

End Sub

Public Property Get Primer() As clsPrimer
    Set Primer = pPrimer
End Property

Public Property Set Primer(c As clsPrimer)
    Set pPrimer = c
End Property

Public Property Get Template() As clsDNA
    Set Template = pTemplate
End Property

Public Property Set Template(c As clsDNA)
    Set pTemplate = c
End Property

Public Sub AddPrimer(AddedPrimer As clsPrimer)
    pPrimers.AddPrimer AddedPrimer, AddedPrimer.Name
End Sub



