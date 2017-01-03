Attribute VB_Name = "XmodDevelopment"
Option Explicit




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



