VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSeqSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPrimers As clsPrimers
Private pTemplate As clsDNA
Private pName As String

Private Sub Class_Initialize()

    If pPrimers Is Nothing Then
        Set pPrimers = New clsPrimers
    End If
    
    If pTemplate Is Nothing Then
        Set pTemplate = New clsDNA
    End If

End Sub

Private Sub Class_Terminate()

    Set pPrimers = Nothing
    Set pTemplate = Nothing

End Sub

Public Property Get Primers() As clsPrimers
    Set Primers = pPrimers
End Property

Public Property Set Primers(c As clsPrimers)
    Set pPrimers = c
End Property

Public Property Get Template() As clsDNA
    Set Template = pTemplate
End Property

Public Property Set Template(c As clsDNA)
    Set pTemplate = c
    If pName = "" Then pName = pTemplate.Name
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(S As String)
    pName = S
End Property

Public Sub AddPrimer(AddedPrimer As clsPrimer)
    pPrimers.AddPrimer AddedPrimer
End Sub




