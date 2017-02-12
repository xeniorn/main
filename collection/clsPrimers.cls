VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'Private pPrimers As New Collection
Private pPrimers As Collection
Private pPrimersKeys As Collection

'not sure what would be advantage of this with "pPrimers as collection"
'why would I separate the instantiation and declaration in this case?

Private Sub Class_Initialize()

    If pPrimers Is Nothing Or pPrimersKeys Is Nothing Then
        Set pPrimers = New Collection
        Set pPrimersKeys = New Collection
    End If

End Sub

Private Sub Class_Terminate()

    Set pPrimers = Nothing
    Set pPrimersKeys = Nothing

End Sub

Public Property Get Primers() As Collection
    Set Primers = pPrimers
End Property

Public Property Get Primer(ByVal PrimerKey) As clsPrimer
    Set Primer = pPrimers.Item(PrimerKey)
End Property

Public Sub AddPrimer(ByVal AddedPrimer As clsPrimer, Optional ByVal AllowRedundancy As Boolean = False)
    
    Dim AddName As String
    Dim PrimerExists As Boolean
    
    AddName = AddedPrimer.Name
    PrimerExists = IsElementOf(AddName, Primers)
    
    If AllowRedundancy Then
        If PrimerExists Then
            If Right(AddName, 5) Like "_r###" Then
                If Right(AddName, 3) = "999" Then
                    ErrorReport 1002
                Else
                    'add suffix to allow adding "same" primer, increase suffix
                    AddName = Left(AddName, (Len(AddName) - 3)) & Format((CInt(Right(AddName, 3)) + 1), "000")
                End If
            Else
                'add suffix to allow adding "same" primer
                AddName = AddName & "_r1"
            End If
        End If
    Else
        If PrimerExists Then
            ErrorReport 1003
        End If
    End If
    
    pPrimers.Add AddedPrimer, AddName
    pPrimersKeys.Add AddName
        
            
End Sub

Public Property Get Count() As Integer
    Count = pPrimers.Count
End Property



'Was causing weird crashes
'Public Function NewEnum() As IUnknown
'    R "enum"
'    Set NewEnum = pPrimers.[_NewEnum]
'End Function
