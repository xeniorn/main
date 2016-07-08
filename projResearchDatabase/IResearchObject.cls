VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IResearchObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This is an interface all objects in the database, whether physical or not
'it's to provide an umbrella term for them
'Juraj Ahel, 2016-06-28
'Last update 2016-06-28
'====================================================================================================
Option Explicit

Public Property Get FilePath() As String
End Property

Public Property Let FilePath(inpFilepath As String)
End Property

Public Property Get Name() As String
End Property

Public Property Let Name(inpFilepath As String)
End Property

Public Property Get EntityType() As String
End Property

Public Property Let EntityType(inpFilepath As String)
End Property

Public Sub Add(ByRef ObjectToAdd As Object)
End Sub
