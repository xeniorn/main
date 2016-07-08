VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMyObjectCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2016-05-09, for this and that purpose
'Last update 2016-05-09
'====================================================================================================
Option Explicit

Private Const conClassName As String = "clsMyObjectCollection"

'this is to be used a s template for all single-object-type collections
'just replace all instances of "clsMyObject" with the name of the target object class!

Private pCollection As Collection
Private pKeys As Collection

Private Sub Class_Initialize()

    Set pCollection = New Collection
    Set pKeys = New Collection

End Sub

Private Sub Class_Terminate()

    Set pCollection = Nothing
    Set pKeys = Nothing

End Sub

Private Sub ErrorReport(Optional ErrorNumber As Integer = 0, Optional ErrorString As String = 0)

    Const conDefaultErrorN As Integer = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub

Public Property Get Keys() As Collection
    Set Keys = pKeys
End Property

Public Property Get Collection() As Collection
    Set Collection = pCollection
End Property

Public Property Get Item(Key) As clsMyObject

    Item = Collection.Item(Key)

End Property

Public Sub Add(AddedItem As clsMyObject, Optional Key As String)
    
    If Key = "" Then Key = CStr(Collection.Count + 1) ' need to make this safer... This can lead overlapping keys
        
    Collection.Add AddedItem, Key
    Keys.Add Key, Key
                
End Sub

Public Sub Remove(Remparameter)

    If VarType(Remparameter) = vbInteger Then
        RemoveIndex (Remparameter)
    ElseIf VarType(Remparameter) = vbString Then
        RemoveKey (Remparameter)
    End If

End Sub

Public Sub RemoveKey(Key As String)
    
    If IsElementOf(Key, Me.Collection) Then
        Me.Collection.Remove (Key)
        Keys.Remove (Key)
    End If
        
End Sub

Public Sub RemoveIndex(Index As Integer)
    
    If pCollection.Count >= Index Then
        Collection.Remove (Index)
        Keys.Remove (Index)
    End If
        
End Sub

Public Property Get Count() As Integer
    Count = pCollection.Count
End Property



