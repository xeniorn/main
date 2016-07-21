VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String
Private pSequence As String
Private pSeqSamples As clsSeqSamples

Private Sub Class_Initialize()

    If pSeqSamples Is Nothing Then
        Set pSeqSamples = New clsSeqSamples
    End If

End Sub


Private Sub Class_Terminate()

    Set pSeqSamples = Nothing

End Sub


Public Property Let Name(S As String)
    pName = S
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Sequence(S As String)
    pSequence = S
End Property

Public Property Get Sequence() As String
    Sequence = pSequence
End Property

Public Property Get SeqSamples() As clsSeqSamples
    Set SeqSamples = pSeqSamples
End Property

Public Property Set SeqSamples(c As clsSeqSamples)
    Set pSeqSamples = c
End Property

Public Sub AddSeqSample(AddedSeqSample As clsSeqSample)
    pSeqSamples.Add AddedSeqSample
End Sub

Public Sub Define( _
        Optional Name As String = "", _
        Optional Sequence As String = "" _
        )
    Me.Name = Name
    Me.Sequence = Sequence
    
End Sub

Public Property Get Length() As Long
    Length = Len(pSequence)
End Property

