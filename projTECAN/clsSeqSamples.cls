VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSeqSamples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'Private pSeqSamples As New Collection
Private pSeqSamples As Collection
Private pSeqSamplesKeys As Collection

'not sure what would be advantage of this with "pSeqSamples as collection"
'why would I separate the instantiation and declaration in this case?

Private Sub Class_Initialize()

    If pSeqSamples Is Nothing Or pSeqSamplesKeys Is Nothing Then
        Set pSeqSamples = New Collection
        Set pSeqSamplesKeys = New Collection
    End If

End Sub

Private Sub Class_Terminate()

    Set pSeqSamples = Nothing
    Set pSeqSamplesKeys = Nothing

End Sub

Public Property Get SeqSamples() As Collection
    Set SeqSamples = pSeqSamples
End Property

Public Property Get SeqSample(SeqSampleKey) As clsSeqSample
    Set SeqSample = pSeqSamples.Item(SeqSampleKey)
End Property

Public Sub Add(AddedSeqSample As clsSeqSample)
    pSeqSamples.Add AddedSeqSample, AddedSeqSample.Name
    pSeqSamplesKeys.Add AddedSeqSample.Name
End Sub

Public Property Get Count() As Integer
    Count = pSeqSamples.Count
End Property



'Was causing weird crashes
'Public Function NewEnum() As IUnknown
'    R "enum"
'    Set NewEnum = pSeqSamples.[_NewEnum]
'End Function

