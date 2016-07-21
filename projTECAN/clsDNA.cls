VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDNA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String
Private pSequence As String
Private pCircular As Boolean

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

Public Property Let Circular(S As Boolean)
    pCircular = S
End Property

Public Property Get Circular() As Boolean
    Circular = pCircular
End Property

Public Property Get Length() As Long
    Length = Len(pSequence)
End Property

Public Sub Define( _
        Optional Name As String = "", _
        Optional Sequence As String = "", _
        Optional Circular As Boolean = False _
        )
    Me.Name = Name
    Me.Sequence = Sequence
    Me.Circular = Circular
End Sub
        

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub
