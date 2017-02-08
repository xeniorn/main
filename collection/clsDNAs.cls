VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDNAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'2017-02-08 add ContainsDNASeq function

Option Explicit

'Private pDNAs As New Collection
Private pDNAs As Collection
Private pDNAsKeys As Collection

'not sure what would be advantage of this with "pDNAs as collection"
'why would I separate the instantiation and declaration in this case?

Private Sub Class_Initialize()

    If pDNAs Is Nothing Or pDNAsKeys Is Nothing Then
        Set pDNAs = New VBA.Collection
        Set pDNAsKeys = New VBA.Collection
    End If

End Sub

Private Sub Class_Terminate()

    Set pDNAs = Nothing
    Set pDNAsKeys = Nothing

End Sub

Public Property Get DNAs() As VBA.Collection
    Set DNAs = pDNAs
End Property

Public Property Get DNA(ByVal DNAKey As Variant) As clsDNA
    Set DNA = pDNAs.Item(DNAKey)
End Property

'Public Property Set DNA(ByVal DNAKey As Variant, ByVal tDNA As clsDNA)
'    Set pDNAs.Item(DNAKey) = tDNA
'End Property

Public Sub AddNewDNA(ByVal Sequence As String, Optional ByVal Name As String = "no_name", Optional ByVal Circular As Boolean = False)

    Dim tDNA As clsDNA
    
    Set tDNA = New clsDNA
    
    If Len(Name) > 0 Then
        tDNA.Name = Name
    Else
        tDNA.Name = "no_name"
    End If
    
    tDNA.Circular = Circular
    
    tDNA.Sequence = DNAParseTextInput(Sequence)
    
    Call AddDNA(tDNA)
    
    Set tDNA = Nothing

End Sub

Public Sub AddDNA(AddedDNA As clsDNA, Optional ByVal Key As String = "", Optional ByVal AllowRedundancy As Boolean = False)
    
    Dim AddName As String
    Dim DNAExists As Boolean
    
    If Key = "" Then
        AddName = AddedDNA.Name
    Else
        AddName = Key
    End If
    
    DNAExists = IsElementOf(AddName, DNAs)
    
    If AllowRedundancy Then
        If DNAExists Then
            If Right(AddName, 5) Like "_r###" Then
                If Right(AddName, 3) = "999" Then
                    Err.Raise 1002
                Else
                    'add suffix to allow adding "same" DNA, increase suffix
                    AddName = Left(AddName, (Len(AddName) - 3)) & Format((CInt(Right(AddName, 3)) + 1), "000")
                End If
            Else
                'add suffix to allow adding "same" DNA
                AddName = AddName & "_r1"
            End If
        End If
    Else
        If DNAExists Then
            Err.Raise 1003
        End If
    End If
    
    pDNAs.Add AddedDNA, AddName
    pDNAsKeys.Add AddName
        
            
End Sub

Public Function ContainsDNASeq(ByVal inpDNA As clsDNA) As Boolean

    Dim i As Long
    
    For i = 1 To Count
        If inpDNA.Sequence = DNA(i).Sequence Then
            ContainsDNASeq = True
            Exit Function
        End If
    Next i

End Function

Public Function GetDNABySeq(ByVal inpSeq As String) As clsDNA

    Dim i As Long
        
    For i = 1 To Count
        If DNA(i).Sequence = inpSeq Then
            Set GetDNABySeq = DNA(i)
            Exit Function
        End If
    Next i

End Function

Public Property Get Count() As Integer
    Count = pDNAs.Count
End Property



'Was causing weird crashes
'Public Function NewEnum() As IUnknown
'    R "enum"
'    Set NewEnum = pDNAs.[_NewEnum]
'End Function

