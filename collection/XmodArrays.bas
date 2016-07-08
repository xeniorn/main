Attribute VB_Name = "XmodArrays"
Option Explicit

'************************************************************************************************
Function IsArrayAllocated(Arr As Variant) As Boolean
'===============================================================================
'taken from Chip Pearson    http://www.cpearson.com/excel/isarrayallocated.aspx
'Juraj Ahel, 2016-06-08
'Last update 2016-06-08
'===============================================================================
    
    
    On Error Resume Next
    IsArrayAllocated = IsArray(Arr) And _
        Not IsError(LBound(Arr, 1)) And _
        LBound(Arr, 1) <= UBound(Arr, 1)

End Function

'************************************************************************************************
Sub ArrayCopy(ByVal SourceArray As Variant, _
              ByRef TargetArray As Variant, _
              Optional ByVal TargetStartIndex As Long = 1, _
              Optional ByVal Overwrite As Boolean = True)

'===============================================================================
'
'Juraj Ahel, 2016-06-08
'Last update 2016-06-08
'2016-06-12 add Overwrite flag and make overwriting default
'===============================================================================
'non-overwriting mode isn't tested and doesn't have good behavior in general
'so don't use it.
              
    Dim i As Long
    Dim IndexOffset As Long
    Dim SourceStartIndex As Long
    Dim SourceEndIndex As Long
    
    If MatrixDimesionNumber(SourceArray) <> 1 Then
        Err.Raise 1001, , "Source variable is not a 1D array"
    End If
    
    SourceStartIndex = LBound(SourceArray)
    SourceEndIndex = UBound(SourceArray)
        
    IndexOffset = TargetStartIndex - SourceStartIndex
    
    'if the array exists and overwrite flag is false, then don't delete old data!
    If Not Overwrite Then
        If IsArrayAllocated(TargetArray) Then
            If MatrixDimesionNumber(TargetArray) <> 1 Then
                Err.Raise 1001, , "Target variable is not a 1D array"
            Else
                ReDim Preserve TargetArray(TargetStartIndex To SourceEndIndex + IndexOffset)
            End If
        Else
        'if it's not allocated, it should be treated as if overwrite is on (to generate it)
            Overwrite = True
        End If
    End If
    
    If Overwrite Then
        ReDim TargetArray(TargetStartIndex To SourceEndIndex + IndexOffset)
    End If
    
    'TODO: add check for datatype compatibility
    
    For i = SourceStartIndex To SourceEndIndex
        TargetArray(i + IndexOffset) = SourceArray(i)
    Next i
              
End Sub

