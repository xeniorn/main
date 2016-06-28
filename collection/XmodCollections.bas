Attribute VB_Name = "XmodCollections"
Option Explicit

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

