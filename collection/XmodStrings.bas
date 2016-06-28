Attribute VB_Name = "XmodStrings"
Option Explicit

'****************************************************************************************************
Function SequenceRangeSelect(InputString As String, IndexRange As String, Optional DNA As Boolean = False, Optional Separator As String = "-") As String

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
Function SubSequenceSelect(InputString As String, StartIndex As Long, EndIndex As Long, Optional DNA As Boolean = False) As String

'====================================================================================================
'Like "Mid" function, but taking indices as arguments, not start index + length
'When end < start, gives reverse
'If DNA is true, the reverse complement for end < start
'Juraj Ahel, 2015-02-16, for vector subselection and general purposes
'Last update 2015-02-16
'====================================================================================================

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
Function StringCharCount(InputString As String, ParamArray Substrings() As Variant) As Long

'====================================================================================================
'Counts the total number of occurrences of any of the listed characters in the given string
'also works for occurrences of longer substrings, but it is "stupid" and it will count overlapping
'substrings regardless of overlap!
'Juraj Ahel, 2015-01-28, for Mutagenesis table programs
'Last update 2015-02-04
'====================================================================================================

Dim i As Long
Dim temp() As Long

N = UBound(Substrings) - LBound(Substrings) + 1
ReDim temp(1 To N)

Dim StringLength As Long
StringLength = Len(InputString)

For i = 1 To N
    temp(i) = (StringLength - Len(Replace(InputString, Substrings(i - 1), ""))) / Len(Substrings(i - 1))
Next i

Dim result As Long
result = WorksheetFunction.Sum(temp)
StringCharCount = result

End Function

'****************************************************************************************************
Function StringCharCount_IncludeOverlap(InputString As String, ParamArray Substrings() As Variant) As Long

'====================================================================================================
'Counts independetly and sums the number of ocurrences of the given sequences in the main sequence
'Counts with overlaps, i.e. AAA counts as two times "AA".
'Juraj Ahel, 2015-02-18, for OligoTm calculations
'Last update 2015-02-18
'2015-03-24 Result was resetting after each iteration, moved Result = 0 outside of loop
'====================================================================================================

Dim i As Long, j As Long
Dim result As Long
Dim N As Long

N = UBound(Substrings) - LBound(Substrings) + 1

Dim StringLength As Long, SubstringLength As Long, Limit As Long
StringLength = Len(InputString)

result = 0

For i = 1 To N

    SubstringLength = Len(Substrings(i - 1))
    
    j = InStr(1, InputString, Substrings(i - 1))
            
    Do While j > 0
        result = result + 1
        j = InStr(j + 1, InputString, Substrings(i - 1))
    Loop
         
Next i

StringCharCount_IncludeOverlap = result

End Function

'****************************************************************************************************
Function StringCompare(a As String, b As String, Optional Limit As Long = 10, Optional mode As String = "Verbose") As String

'====================================================================================================
'Compares two strings and lists their differences, very raw so far
'Juraj Ahel, 2015-02-12, for comparing of protein sequences to find point mutations
'Last update 2015-02-12
'====================================================================================================

Dim i As Long, j As Long
Dim result As String, S As String
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
        result = result & S & i
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
        result = result & S & i & "(" & cA & ">" & cB & ")"
    End If
Loop Until i = LA Or i = Lb Or ((counter > Limit) And (Limit > 0))

If counter = 0 And LA = Lb Then
    result = "Exact Copy!"
    GoTo 99
End If

End Select

If LA <> Lb Then result = result & S & "LenDiff=" & LA - Lb

If Len(result) > 0 Then result = Right(result, Len(result) - Len(S))

If counter > Limit And Limit > 0 Then result = "Threshold (" & Limit & ") reached!"

99 StringCompare = result

End Function

'****************************************************************************************************
Function StringRemoveNonPrintable(InputString As String) As String
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
Function StringJoin(RangeToJoin As Range, Optional Separator As String = "", Optional Direction As Long) As String

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
Function StringFindOverlap(Probe As String, Target As String)

'====================================================================================================
'Finds the (largest) continuous perfectoverlap between two strings
'Juraj Ahel, 2015-04-30, for general purposes
'Last update 2015-04-30
'2016-06-28 explicit variable declaration
'====================================================================================================

Dim ProbeLength As Long, TargetLength As Long
Dim Results() As Long
Dim wStart As Long
Dim tempResult As String
Dim OverlapWidth As Long

ProbeLength = Len(Probe)
TargetLength = Len(Target)

If ProbeLength > TargetLength Then
    Call SwapValue(Probe, Target)
    Call SwapValue(ProbeLength, TargetLength)
End If
    
wStart = ProbeLength

If wStart = 0 Then
    tempResult = "Zero-string probe or target."
    GoTo 999
End If

'- if I want to map them all
'ReDim Results(1 To wStart, 1 To wStart)

'- if I want to extract the longest ones only
ReDim Results(1 To wStart)

Dim i As Long, j As Long, k As Long, W As Long
Dim TempProbe As String
Dim FoundOverlap As Boolean

W = wStart

Do
    k = 0
    
    For i = 1 To 1 + (wStart - W)
    
        TempProbe = Mid(Probe, i, W)
        
        j = 0
        Do
            j = InStr(j + 1, Target, TempProbe)
            FoundOverlap = (j > 0)
            
            'k = k + FoundOverlap
            'Results(w, k) = FoundOverlap * j
            
            If FoundOverlap Then
                k = k + 1
                Results(k) = j
            End If
        Loop Until Not FoundOverlap
        
    Next i
    
    W = W - 1
    
Loop Until k <> 0 Or W = 0

OverlapWidth = W + 1

Dim TempResultAsStrings() As String

Select Case k
    Case 0
        tempResult = "#! No overlap found."
    Case 1
        tempResult = Mid(Target, Results(1), OverlapWidth)
    Case Is > 1
        ReDim TempResultAsStrings(1 To k)
        For i = 1 To k
            TempResultAsStrings(i) = CStr(Results(i))
        Next i
    
        
        tempResult = "Multiple equivalent results of length " _
                    & OverlapWidth & " at positions: " _
                    & Join(TempResultAsStrings, ";")
End Select

999 StringFindOverlap = tempResult

End Function

'****************************************************************************************************
Function StringSubstract(Template As String, _
                        ParamArray Substractions() As Variant _
                        ) As String

'====================================================================================================
'Removes all instances of given substrings from the template sequence, even if overlapping
'Juraj Ahel, 2015-04-30, for general purposes
'Last update 2015-04-30
'====================================================================================================

Dim TemplateLength As Long, SubstractionLengths() As Long
Dim TemplateArray() As String
Dim NumberOfSubstractions As Long
Dim i, j As Long
Dim FoundTarget As Boolean

TemplateLength = Len(Template)
ReDim TemplateArray(1 To TemplateLength)

For i = 1 To TemplateLength
    TemplateArray(i) = Mid(Template, i, 1)
Next i

NumberOfSubstractions = UBound(Substractions) - LBound(Substractions) + 1

For i = 1 To NumberOfSubstractions
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
Next i

StringSubstract = Join(TemplateArray, "")

End Function

