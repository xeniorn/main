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

Public Property Let Name(ByVal S As String)
    pName = S
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Sequence(ByVal S As String)
    pSequence = S
End Property

Public Property Get Sequence() As String
    Sequence = pSequence
End Property

Public Property Let Circular(ByVal S As Boolean)
    pCircular = S
End Property

Public Property Get Circular() As Boolean
    Circular = pCircular
End Property

Public Property Get Length() As Long
    Length = Len(pSequence)
End Property

Public Sub Define( _
        Optional ByVal Name As String = "", _
        Optional ByVal Sequence As String = "", _
        Optional ByVal Circular As Boolean = False _
        )
    Me.Name = Name
    Me.Sequence = Sequence
    Me.Circular = Circular
End Sub

Public Function DefineNew( _
        Optional ByVal Name As String = "", _
        Optional ByVal Sequence As String = "", _
        Optional ByVal Circular As Boolean = False _
        ) As clsDNA
        
    Dim tDNA As clsDNA
    
    Set tDNA = New clsDNA
    
    With tDNA
        .Name = Name
        .Sequence = Sequence
        .Circular = Circular
    End With
    
    Set DefineNew = tDNA
    
    Set tDNA = Nothing
    
End Function

        
Public Function ChangeName(ByVal NewName As String) As clsDNA
    
    Me.Name = NewName
    Set ChangeName = Me

End Function


