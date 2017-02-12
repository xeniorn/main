Attribute VB_Name = "XmodStrings"
Option Explicit

'****************************************************************************************************
Public Function StringOffsetCircular(ByVal InputString As String, ByVal Offset As Long) As String

'====================================================================================================
'a string is assumed to be circular - and the origin is shifted by "Offset"
'Juraj Ahel, 2016-07-21

'====================================================================================================

    Dim NewLeft As String, NewRight As String
    Dim StringLength As Long
        
    StringLength = Len(InputString)
    
    Offset = Offset Mod StringLength
    
    Select Case True
        Case (Offset = 0)
            NewLeft = InputString
            NewRight = ""
        Case (Offset > 0)
            NewLeft = Right(InputString, StringLength - Offset)
            NewRight = Left(InputString, Offset)
        Case (Offset < 0)
            NewLeft = Right(InputString, Offset)
            NewRight = Left(InputString, StringLength - Offset)
    End Select

    StringOffsetCircular = NewLeft & NewRight
    

End Function

'****************************************************************************************************
Public Function SequenceRangeSelect(ByVal InputString As String, ByVal IndexRange As String, Optional ByVal DNA As Boolean = False, Optional ByVal Separator As String = "-") As String

'====================================================================================================
'Like SubSequenceSelect, but taking a string formatted as a "range" (e.g. 15-150)
'When end < start, gives reverse
'If DNA is true, the reverse complement for end < start
'Juraj Ahel, 2015-02-16, for vector subselection and general purposes
'Last update 2015-02-16
'====================================================================================================

    Dim StartIndex As Long, EndIndex As Long, SeparatorIndex As Long
    
    SeparatorIndex = InStr(1, IndexRange, Separator)
    
    StartIndex = CInt(Left(IndexRange, SeparatorIndex - 1))
    EndIndex = CInt(Right(IndexRange, Len(IndexRange) - SeparatorIndex))
    
    SequenceRangeSelect = SubSequenceSelect(InputString, StartIndex, EndIndex, DNA)

End Function

'****************************************************************************************************
Public Function SubSequenceSelect( _
    ByVal InputString As String, _
    ByVal StartIndex As Long, _
    ByVal EndIndex As Long, _
    Optional ByVal DNA As Boolean = False _
    ) As String

'====================================================================================================
'Like "Mid" function, but taking indices as arguments, not start index + length
'When end < start, gives reverse
'If DNA is true, the reverse complement for end < start
'Juraj Ahel, 2015-02-16, for vector subselection and general purposes
'Last update 2015-02-16
'====================================================================================================
'2016-12-23 make byval

    Dim tempString As String
    
    If StartIndex <= EndIndex Then
    
        tempString = Mid(InputString, StartIndex, EndIndex - StartIndex + 1)
    
    Else
        
        
        tempString = Mid(InputString, EndIndex, StartIndex - EndIndex + 1)
        
        Dim N As Long, i As Long
        Dim TempStringChars() As String
        
        N = Len(tempString)
        ReDim TempStringChars(1 To N)
        
        If DNA Then
            tempString = DNAReverseComplement(tempString)
        Else
            For i = 1 To N: TempStringChars(i) = Mid(tempString, N - i + 1, 1): Next i
            tempString = Join(TempStringChars, "")
        End If
        
    End If
    
    SubSequenceSelect = tempString

End Function
'****************************************************************************************************
Public Function StringCharCount(ByVal InputString As String, ParamArray Substrings() As Variant) As Long

'====================================================================================================
'Counts the total number of occurrences of any of the listed characters in the given string
'also works for occurrences of longer substrings, but it is "stupid" and it will count overlapping
'substrings regardless of overlap!
'Juraj Ahel, 2015-01-28, for Mutagenesis table programs
'Last update 2015-02-04
'====================================================================================================

    Dim i As Long
    Dim temp() As Long
    Dim N As Long
    
    N = UBound(Substrings) - LBound(Substrings) + 1
    ReDim temp(1 To N)
    
    Dim StringLength As Long
    StringLength = Len(InputString)
    
    For i = 1 To N
        temp(i) = (StringLength - Len(Replace(InputString, Substrings(i - 1), ""))) / Len(Substrings(i - 1))
    Next i
    
    Dim Result As Long
    Result = WorksheetFunction.Sum(temp)
    StringCharCount = Result

End Function

'****************************************************************************************************
Public Function StringCharCount_IncludeOverlap(ByVal InputString As String, ParamArray Substrings() As Variant) As Long

'====================================================================================================
'Counts independetly and sums the number of ocurrences of the given sequences in the main sequence
'Counts with overlaps, i.e. AAA counts as two times "AA".
'Juraj Ahel, 2015-02-18, for OligoTm calculations
'Last update 2015-02-18
'2015-03-24 Result was resetting after each iteration, moved Result = 0 outside of loop
'====================================================================================================
'2016-12-23 protect against 0-length substrings
'TODO: add support for arrays and collections as substrings...

    Dim i As Long, j As Long
    Dim Result As Long
    Dim N As Long
    
    N = UBound(Substrings) - LBound(Substrings) + 1
    
    Dim StringLength As Long, SubstringLength As Long, Limit As Long
    StringLength = Len(InputString)
    
    Result = 0
    
    For i = 1 To N
            
        SubstringLength = Len(Substrings(i - 1))
        
        If SubstringLength > 0 Then
            
            j = InStr(1, InputString, Substrings(i - 1))
                    
            Do While j > 0
                Result = Result + 1
                j = InStr(j + 1, InputString, Substrings(i - 1))
            Loop
            
        End If
             
    Next i
    
    StringCharCount_IncludeOverlap = Result

End Function

'****************************************************************************************************
Public Function StringCompare(ByVal a As String, ByVal b As String, Optional ByVal Limit As Long = 10, Optional ByVal mode As String = "Verbose") As String

'====================================================================================================
'Compares two strings and lists their differences, very raw so far
'Juraj Ahel, 2015-02-12, for comparing of protein sequences to find point mutations
'Last update 2015-02-12
'====================================================================================================

    Dim i As Long, j As Long
    Dim Result As String, S As String
    Dim LA As Long, Lb As Long
    Dim counter As Long: counter = 0
    Dim cA As String, cB As String
    
    LA = Len(a): Lb = Len(b)
    
    S = "; "
    
    Select Case UCase(mode)
    
        Case "SHORT", "S"
    
            Do
                i = i + 1
                cA = Mid(a, i, 1)
                cB = Mid(b, i, 1)
                
                If cA <> cB Then
                    counter = counter + 1
                    Result = Result & S & i
                End If
            Loop Until i = LA Or i = Lb Or ((counter > Limit) And (Limit > 0))
    
    
        Case "VERBOSE", "V"
            GoTo 50
    
        Case Else
50
            Do
                i = i + 1
                cA = Mid(a, i, 1)
                cB = Mid(b, i, 1)
                
                If cA <> cB Then
                    counter = counter + 1
                    Result = Result & S & i & "(" & cA & ">" & cB & ")"
                End If
            Loop Until i = LA Or i = Lb Or ((counter > Limit) And (Limit > 0))
            
            If counter = 0 And LA = Lb Then
                Result = "Exact Copy!"
                GoTo 99
            End If
    
    End Select
    
    If LA <> Lb Then Result = Result & S & "LenDiff=" & LA - Lb
    
    If Len(Result) > 0 Then Result = Right(Result, Len(Result) - Len(S))
    
    If counter > Limit And Limit > 0 Then Result = "Threshold (" & Limit & ") reached!"
    
99     StringCompare = Result

End Function

'****************************************************************************************************
Public Function StringRemoveNonPrintable(ByVal InputString As String) As String
'====================================================================================================
'Removes all the nonprintable characters from a string
'Juraj Ahel, 2016-03-09, for automatic handling of UNICORN 3.1 res files
'Last update 2016-03-09
'====================================================================================================

    StringRemoveNonPrintable = StringSubstract(InputString, _
        Chr(0), Chr(1), Chr(2), Chr(3), Chr(4), Chr(5), Chr(6), Chr(7), _
        Chr(8), Chr(9), Chr(10), Chr(11), Chr(12), Chr(13), Chr(14), Chr(15), _
        Chr(16), Chr(17), Chr(18), Chr(19), Chr(20), Chr(21), Chr(22), Chr(23), _
        Chr(24), Chr(25), Chr(26), Chr(27), Chr(28), Chr(29), Chr(30), Chr(31) _
        )
    
End Function

'****************************************************************************************************
Public Function StringJoin(ByVal RangeToJoin As Range, Optional ByVal Separator As String = "", Optional ByVal Direction As Long) As String

'====================================================================================================
'Joins all the cell values in an array as strings
'Juraj Ahel, 2015-02-16, for general purposes
'Last update 2015-04-13
'====================================================================================================
'Direction not yet implemented

    Dim tempString As String
    Dim cell As Range
    
    For Each cell In RangeToJoin
        tempString = tempString & cell.Value & Separator
    Next cell
    
    StringJoin = tempString

End Function




'****************************************************************************************************
Public Function StringFindSubstringLocations(ByVal Probe As String, ByVal Target As String) As VBA.Collection

'====================================================================================================
'Finds all the instances of Probe in Target and gives indices as a collection
'Juraj Ahel, 2017-01-23
'====================================================================================================

    Dim i As Long
    Dim tColl As VBA.Collection
    
    i = 0
    
    Set tColl = New VBA.Collection
    
    Do
        i = InStr(i + 1, Target, Probe)
        
        If i > 0 Then
            tColl.Add i
        Else
            Exit Do
        End If
    
    Loop
    
    Set StringFindSubstringLocations = tColl
    
    Set tColl = Nothing

End Function

'****************************************************************************************************
Public Function StringFindOverlap(ByVal Probe As String, ByVal Target As String, Optional Interactive As Boolean = True)
'====================================================================================================
'Finds the (largest) continuous perfectoverlap between two strings
'Juraj Ahel, 2015-04-30, for general purposes
'Last update 2015-04-30
'2016-06-28 explicit variable declaration
'====================================================================================================
'2016-12-21 add Interactive option, make byval
'2017-01-03 correct the limit of the number of overlaps found to the actual theoretical maximum
'2017-01-20 Rewrite for efficiency, it was damn slow on long strings
'2017-01-22 Rewrite was full of bugs and simply wrong in most cases, corrected

    Dim ProbeLength As Long, TargetLength As Long
    
    Dim MaxPossibleOverlap As Long
    Dim MinPossibleOverlap As Long
    Dim CurrentOverlap As Long
    
    Dim tempProbe As String
    
    Dim tempIndex As Long
    Dim LastIndexFound As Long
    Dim LastOverlap As Long
    
    Dim OverlapCount As Long
    Dim Indices As VBA.Collection
    Dim tColl As VBA.Collection
    
    Dim TempResultAsStrings() As String
    Dim tempResult As String
    
    Dim OverlapFound As Boolean
    
    Dim j As Long
    
    'Dim ExpansionLength As Long
    'Dim FirstIndexToSearch As Long
    'Dim LastIndexToSearch As Long
    
    ProbeLength = Len(Probe)
    TargetLength = Len(Target)
    
    If ProbeLength > TargetLength Then
        Call SwapValue(Probe, Target)
        Call SwapValue(ProbeLength, TargetLength)
    End If
    
    MinPossibleOverlap = 0
    MaxPossibleOverlap = ProbeLength
    
    CurrentOverlap = MaxPossibleOverlap
    LastOverlap = -1
    LastIndexFound = 1
    
    OverlapFound = False
    
    'first, find the longest overlap length, later we'll check how many there are
    Do
    
            OverlapFound = False
            
            'check if we searched the entire space
NextSize:   If MinPossibleOverlap = MaxPossibleOverlap Then
                
                If LastOverlap = MinPossibleOverlap Then
                    GoTo YesOverlap
                Else
                    GoTo NoOverlap
                End If
                
            Else
                
                'ExpansionLength = CurrentOverlap - LastOverlap
                'FirstIndexToSearch = LastIndexFound - ExpansionLength
                'If FirstIndexToSearch < 1 Then FirstIndexToSearch = 1
                'LastIndexToSearch = LastIndexFound + ExpansionLength
                'If LastIndexToSearch + CurrentOverlap - 1 > ProbeLength Then LastIndexToSearch = ProbeLength - CurrentOverlap + 1
                            
                For j = 1 To ProbeLength - CurrentOverlap + 1
                    tempProbe = Mid(Probe, j, CurrentOverlap)
                
                    tempIndex = InStr(1, Target, tempProbe)
                
                    If tempIndex > 0 Then
                        'overlap of this length was found - search higher
                        LastOverlap = CurrentOverlap
                        LastIndexFound = tempIndex
                        MinPossibleOverlap = CurrentOverlap
                        CurrentOverlap = (CurrentOverlap + MaxPossibleOverlap + 1) \ 2
                        OverlapFound = True
                        Exit For
                    End If
                Next j
                
                If Not OverlapFound Then
                    'no overlap of this length was found, search lower
                    If CurrentOverlap > MinPossibleOverlap Then
                        MaxPossibleOverlap = CurrentOverlap - 1
                        CurrentOverlap = (MinPossibleOverlap + CurrentOverlap) \ 2
                    End If
                End If
                
            End If
    
    Loop


YesOverlap:
    
    tempIndex = 1
    
    Set Indices = New VBA.Collection
    
    For j = 1 To ProbeLength - CurrentOverlap + 1
    
        tempProbe = Mid(Probe, j, LastOverlap)
        tempIndex = 0
    
        Do
        
            tempIndex = InStr(tempIndex + 1, Target, tempProbe)
            
            If tempIndex <> 0 Then
                Set tColl = New VBA.Collection
                tColl.Add tempIndex
                tColl.Add tempProbe
                Indices.Add tColl
            End If
            
        Loop While tempIndex <> 0
            
    Next j
        
NoOverlap:
    
    If Indices Is Nothing Then Set Indices = New VBA.Collection
    
    Select Case Indices.Count
    
        Case 0
            tempResult = "#! No overlap found."
        Case 1
            tempResult = Indices.Item(1).Item(2)
        Case Is > 1
            'ReDim TempResultAsStrings(1 To Indices.Count)
            'For j = 1 To Indices.Count
            '    TempResultAsStrings(j) = CStr(Indices.Item(j).Item(1))
            'Next j
        
            If Interactive Then
                tempResult = "Multiple equivalent results of length " _
                            & LastOverlap '& " at positions: " _
                            '& Join(TempResultAsStrings, ";")
            Else
                StringFindOverlap = Indices.Item(1).Item(2)
                Call ApplyNewError(jaErr + 1, "StringFindOverlap", "(" & Indices.Count & ") overlaps of same length found")
                Exit Function
            End If
    End Select
    
    StringFindOverlap = tempResult
    
    Set Indices = Nothing
    Set tColl = Nothing

End Function



'****************************************************************************************************
Public Function LongestCommonSubstring(ByVal S1 As String, ByVal S2 As String) As String

    Dim MaxSubstrStart
    Dim MaxLenFound
    Dim i1
    Dim i2
    Dim x
    

    MaxSubstrStart = 1
    MaxLenFound = 0
    For i1 = 1 To Len(S1)
        For i2 = 1 To Len(S2)
            x = 0
            While i1 + x <= Len(S1) And _
                i2 + x <= Len(S2) And _
                    Mid(S1, i1 + x, 1) = Mid(S2, i2 + x, 1)
                    x = x + 1
            Wend
            If x > MaxLenFound Then
                MaxLenFound = x
                MaxSubstrStart = i1
            End If
        Next
    Next
    LongestCommonSubstring = Mid(S1, MaxSubstrStart, MaxLenFound)
End Function

'****************************************************************************************************
Public Function StringSubstract(ByVal Template As String, _
                        ParamArray Substractions() As Variant _
                        ) As String

'====================================================================================================
'Removes all instances of given substrings from the template sequence, even if overlapping
'Juraj Ahel, 2015-04-30, for general purposes
'Last update 2015-04-30
'====================================================================================================
'2016-12-22 add control for empty string in substractions / template

    Dim TemplateLength As Long, SubstractionLengths() As Long
    Dim TemplateArray() As String
    Dim NumberOfSubstractions As Long
    Dim i As Long, j As Long, k As Long
    Dim FoundTarget As Boolean
    
    TemplateLength = Len(Template)
    
    If TemplateLength = 0 Then
        StringSubstract = vbNullString
        Exit Function
    End If
    
    ReDim TemplateArray(1 To TemplateLength)
    
    For i = 1 To TemplateLength
        TemplateArray(i) = Mid(Template, i, 1)
    Next i
    
    NumberOfSubstractions = UBound(Substractions) - LBound(Substractions) + 1
    
    For i = 1 To NumberOfSubstractions
        If Len(Substractions(i - 1)) > 0 Then
            j = 0
            Do
                j = InStr(j + 1, Template, Substractions(i - 1))
                FoundTarget = (j > 0)
                If FoundTarget Then
                    For k = 1 To Len(Substractions(i - 1))
                        TemplateArray(j + k - 1) = ""
                    Next k
                End If
            Loop Until Not FoundTarget
        End If
    Next i
    
    StringSubstract = Join(TemplateArray, "")

End Function



