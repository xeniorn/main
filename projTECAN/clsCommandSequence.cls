VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCommands As Collection
Private pCommandIDs As Collection
Private pLastID As Integer

Private Sub Class_Initialize()
    
    Set pCommands = New Collection
    Set pCommandIDs = New Collection

End Sub

Private Sub Class_Terminate()

    Set pCommands = Nothing
    Set pCommandIDs = Nothing
    
End Sub


Public Property Get Output(Optional Delimiter As String = vbCrLf) As String

    Dim i As Integer
    Dim tempArray() As String
    Dim tempResult As String
    
    tempResult = ""
    
    If Not (pCommands Is Nothing) Then
        If pCommands.Count > 0 Then
        
            ReDim tempArray(1 To pCommands.Count)
            For i = LBound(tempArray) To UBound(tempArray)
                tempArray(i) = pCommands.Item(i)
            Next i
            tempResult = Join(tempArray, Delimiter)
        
        End If
    End If
                
    
    
    Output = tempResult

End Property

Public Sub Purge()

    Set pCommands = Nothing
    Set pCommandIDs = Nothing
    
    Set pCommands = New Collection
    Set pCommandIDs = New Collection
    
    pLastID = 0

End Sub

Public Sub Append(CommandText As String) 'append a Command class?

    Dim tempID As Integer
    
    tempID = MakeNextID
    
    pCommands.Add CommandText, CStr(tempID)
    pCommandIDs.Add tempID, CStr(tempID)

End Sub

Private Function MakeNextID() As Integer
    MakeNextID = LastID + 1
    pLastID = MakeNextID
End Function

Private Property Get LastID() As Integer
    LastID = pLastID
End Property


