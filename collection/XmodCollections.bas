Attribute VB_Name = "XmodCollections"
Option Explicit

'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2016-03-09
'Last update 2016-03-09
'====================================================================================================
Public Function IsElementOf(sKey As String, oCollection As Collection) As Boolean
    'Const scSOURCE As String = "IsElementOf"

    Dim lErrNumber As Long
    Dim sErrDescription As String

    lErrNumber = 0
    sErrDescription = "unknown error occurred"
    Err.Clear
    On Error Resume Next
        ' note: just access the item - no need to assign it to a dummy value
        ' and this would not be so easy, because we would need different
        ' code depending on the type of object
        ' e.g.
        '   Dim vItem as Variant
        '   If VarType(oCollection.Item(sKey)) = vbObject Then
        '       Set vItem = oCollection.Item(sKey)
        '   Else
        '       vItem = oCollection.Item(sKey)
        '   End If
        oCollection.Item sKey
        lErrNumber = CLng(Err.Number)
        sErrDescription = Err.Description
    On Error GoTo 0

    If lErrNumber = 5 Then ' 5 = not in collection
        IsElementOf = False
    ElseIf (lErrNumber = 0) Then
        IsElementOf = True
    'Else
        ' Re-raise error
    '    err.Raise lErrNumber, mscMODULE & ":" & scSOURCE, sErrDescription
    End If
End Function

'****************************************************************************************************
Sub CollectionAppend( _
    ByRef Collection1 As VBA.Collection, _
    ByVal Collection2 As VBA.Collection _
    )
    
'====================================================================================================
'Merges 2 Collections into the first one
'Juraj Ahel, 2016-06-28
'
'====================================================================================================

    Dim i As Long
            
    If Collection1 Is Nothing Then
        Set Collection1 = New VBA.Collection
    End If
    
    
    If Not Collection2 Is Nothing Then
        For i = 1 To Collection2.Count
            Collection1.Add Collection2.Item(1)
            Collection2.Remove (1)
        Next i
    End If
    
End Sub

'****************************************************************************************************
Sub SortStringCollectionByLength(ByRef StringCollection As VBA.Collection)
'====================================================================================================
'sorts. string. collection. by. length. figures....
'Juraj Ahel, 2016-06-28
'
'====================================================================================================
'AllowORFOverlap not yet implemented - but it can be easily acquired by just running DNALongestORF

    Dim Lengths As VBA.Collection
    Dim SortedCollection As VBA.Collection
    Dim i As Long
    Dim tempLen As Long
    Dim tempIndex As Long
    
    If Not StringCollection Is Nothing Then
        If StringCollection.Count > 0 Then
        
            Set Lengths = New VBA.Collection
            Set SortedCollection = New VBA.Collection
            
            For i = 1 To StringCollection.Count
                Lengths.Add Len(StringCollection.Item(i))
            Next i
                        
            Do While StringCollection.Count > 0
                
                tempLen = 0
                tempIndex = 1
                
                For i = 1 To StringCollection.Count
                    If Lengths.Item(i) > tempLen Then
                        tempLen = Lengths.Item(i)
                        tempIndex = i
                    End If
                Next i
                
                SortedCollection.Add StringCollection.Item(tempIndex)
                StringCollection.Remove (tempIndex)
                Lengths.Remove (tempIndex)
                                
            Loop
        
        End If
    End If
    
    Set Lengths = Nothing
    Set StringCollection = SortedCollection
    Set SortedCollection = Nothing
          
End Sub

