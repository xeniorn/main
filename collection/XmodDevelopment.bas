Attribute VB_Name = "XmodDevelopment"
Option Explicit

Sub testaaaa()

    Dim a
    
    Set a = DNAFindORFs("ATGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACGTACATCGACTATCGATCGATCGATCGACGCGCCATGCATCGATCGACTAGCTAGTCGATCGTAGCTACGTACGTAGTAG")
    
    

End Sub



'************************************************************************************
Function CloningMakeConstructs( _
         ByVal ProteinSequence As String, _
         ByVal DNASource As String, _
         ByVal TruncationList As String, _
         Optional ByVal Circular As Boolean = True, _
         Optional ByVal CheckReverseComplement As Boolean = True, _
         Optional ByVal Interactive As Boolean = True _
         ) As VBA.Collection

'====================================================================================================
'Takes in a protein sequence, DNA source, and list of truncations to introduce
'Formulates a cloning strategy - fragments to clone out + primers to get these fragments from the soruce
'Juraj Ahel, 2017-01-03
'====================================================================================================
    
    Const MyName As String = "CloningMakeConstructs"
        
'TODO: parse protein / DNA sequences / Trunc list
    
    Dim i As Long
    
    Dim ProteinLength As Long
    Dim DNALength As Long
    Dim ORFLength As Long
    
    Dim ORFLocus As Long
    Dim IsReverse As Boolean
    
    Dim tColl As VBA.Collection
    
    ProteinLength = Len(ProteinSequence)
    DNALength = Len(DNASource)
    ORFLength = 3 * ProteinLength + 3
    
    If ProteinLength = 0 Or DNALength = 0 Then
        Call ApplyNewError(jaErr + 18, MyName, "Empty input")
        If Interactive Then
            ErrReraise
        Else
            Set CloningMakeConstructs = Nothing
            Exit Function
        End If
    End If
        
    '1 confirm DNA encodes for full protein
    
    Set tColl = DNAFindProteinInTemplate(ProteinSequence, DNASource, Circular, CheckReverseComplement, False)
    
    If Err.Number <> 0 Then
        If Err.Source = "DNAFindProteinInTemplate" Then
            Err.Source = MyName
        End If
        If Interactive Then
            ErrReraise
        Else
            Set CloningMakeConstructs = Nothing
            Exit Function
        End If
    Else
        If Not tColl Is Nothing Then
            If tColl.Count = 2 Then
                ORFLocus = tColl.Item(1)
                IsReverse = tColl.Item(2)
            End If
        End If
        If IsReverse Then DNASource = DNAReverseComplement(DNASource)
    End If
        
        
    Debug.Print ("DNA encodes for protein at locus: " & ORFLocus & " Reverse Strand: " & IsReverse)
       
       
       
    
            
        '2 formulate truncated sequences
    
        '3 in silico truncate DNA
    
        '4 design Gibson assembly
    
        '5 confirm PCR / gibson / translation of assembly
    
    

    End Function


'************************************************************************************
Function SmoothData(DataArray As Variant, WindowSize As Long) As Variant

    Dim DataLength As Long
    Dim TempOutput() As Variant
    Dim tempIndex As Long
    Dim i As Long, j As Long
    Dim tempsum As Double
        
        
    DataLength = 1 + UBound(DataArray) - LBound(DataArray)
    
    ReDim TempOutput(1 To DataLength)
        
    
            
    For i = 1 To DataLength - WindowSize
        
        tempIndex = i + WindowSize \ 2
        tempsum = 0
        
        For j = i To i + WindowSize - 1
            tempsum = tempsum + DataArray(j)
        Next j
        
        TempOutput(tempIndex) = tempsum / WindowSize
        
    Next i
        
        
    For i = 1 To WindowSize \ 2
        TempOutput(i) = 0
        TempOutput(DataLength - i + 1) = 0
    Next i
    
    SmoothData = TempOutput

End Function

'************************************************************************************
Function GetMaxLetterCount(InputString As String) As Long

    Dim i As Byte
    Dim Char As String
    Dim tout As String
    Dim tempCount As Long
    Dim CharCount As Long
    
    tempCount = 0
    
    For i = 65 To 90
    
        Char = Chr(i)
        CharCount = StringCharCount(InputString, Char)
        If CharCount > tempCount Then tempCount = CharCount
        
    Next i
    
    GetMaxLetterCount = tempCount
        

End Function


