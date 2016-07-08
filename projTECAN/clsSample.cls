VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String
Private pLocation As clsWorktableSampleLocation
Private pObject As Object 'to hold the ACTUAL object instead of placeholder clsSample. ###TODO later

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    Set pLocation = Nothing
End Sub

Public Property Get Object() As Object

    

End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(N As String)
    pName = N
End Property

Public Property Get Location() As clsWorktableSampleLocation
    Set Location = pLocation
End Property

Public Property Set Location(inLoc As clsWorktableSampleLocation)
    Set pLocation = inLoc
End Property

Public Sub Locate(Worktable As clsWorktableSetup)
    
    Worktable.Locate Me 'kehe
    
    'Set Worktable = Nothing ' do I need to dereference it?

End Sub
