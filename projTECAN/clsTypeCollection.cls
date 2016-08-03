VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTypeCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'Private pCollection As New Collection
Private pCollection As Collection
Private pKeys As Collection
Private pType As VbVarType
Private pState As String 'make an enum, one for strict, one for loose, possibly other modes, to allow e.g. Longs and Doubles to be stored together

Private Sub Class_Initialize()

    If pCollection Is Nothing Or pKeys Is Nothing Then
        Set pCollection = New Collection
        Set pKeys = New Collection
    End If

End Sub

Public Property Get Keys() As Collection
    Set Keys = pKeys
End Property

Public Property Get Collection() As Collection
    Set Collection = pCollection
End Property

Public Property Get MyType() As VbVarType
    MyType = pType
End Property

Public Property Let MyType(CT As VbVarType)
    If pCollection.Count = 0 Then
        pType = CT
    Else
        ErrorReport 1001
    End If
End Property

Public Property Get Item(Key) As Variant

    Item = Collection.Item(Key)

End Property


Private Function TypeCheckOK(Variable) As Boolean

    If VarType(Variable) = pType Then
        TypeCheckOK = True
    Else
        TypeCheckOK = False
    End If

End Function



Public Sub Add(Added As Variant, Optional Key As String)
    
    If pType = vbEmpty Then pType = VarType(Added)
    
    If TypeCheckOK(Added) Then
        
        If Key = "" Then Key = CStr(Me.Collection.Count + 1)
        
        Me.Collection.Add Added, Key
        Keys.Add Key, Key
        
    Else
        ErrorReport 1000
    End If
                
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
        Me.Collection.Remove (Index)
        Keys.Remove (Index)
    End If
        
End Sub

Public Property Get Count() As Integer
    Count = pCollection.Count
End Property

Private Sub Class_Terminate()

    Set pCollection = Nothing
    Set pKeys = Nothing

End Sub
