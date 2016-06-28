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

Dim StartIndex As Integer, EndIndex As Integer, SeparatorIndex As Integer

SeparatorIndex = InStr(1, IndexRange, Separator)

StartIndex = CInt(Left(IndexRange, SeparatorIndex - 1))
EndIndex = CInt(Right(IndexRange, Len(IndexRange) - SeparatorIndex))

SequenceRangeSelect = SubSequenceSelect(InputString, StartIndex, EndIndex, DNA)

End Function

'****************************************************************************************************
Function SubSequenceSelect(InputString As String, StartIndex As Integer, EndIndex As Integer, Optional DNA As Boolean = False) As String

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
    
    Dim N As Integer, i As Integer
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
Function StringCharCount(InputString As String, ParamArray Substrings() As Variant) As Integer

'====================================================================================================
'Counts the total number of occurrences of any of the listed characters in the given string
'also works for occurrences of longer substrings, but it is "stupid" and it will count overlapping
'substrings regardless of overlap!
'Juraj Ahel, 2015-01-28, for Mutagenesis table programs
'Last update 2015-02-04
'====================================================================================================

Dim i As Integer
Dim temp() As Integer

N = UBound(Substrings) - LBound(Substrings) + 1
ReDim temp(1 To N)

Dim StringLength As Long
StringLength = Len(InputString)

For i = 1 To N
    temp(i) = (StringLength - Len(Replace(InputString, Substrings(i - 1), ""))) / Len(Substrings(i - 1))
Next i

Dim Result As Integer
Result = WorksheetFunction.Sum(temp)
StringCharCount = Result

End Function

'****************************************************************************************************
Function StringCharCount_IncludeOverlap(InputString As String, ParamArray Substrings() As Variant) As Integer

'====================================================================================================
'Counts independetly and sums the number of ocurrences of the given sequences in the main sequence
'Counts with overlaps, i.e. AAA counts as two times "AA".
'Juraj Ahel, 2015-02-18, for OligoTm calculations
'Last update 2015-02-18
'2015-03-24 Result was resetting after each iteration, moved Result = 0 outside of loop
'====================================================================================================

Dim i As Integer, j As Integer
Dim Result As Integer

N = UBound(Substrings) - LBound(Substrings) + 1

Dim StringLength As Integer, SubstringLength As Integer, Limit As Integer
StringLength = Len(InputString)

Result = 0

For i = 1 To N

    SubstringLength = Len(Substrings(i - 1))
    
    j = InStr(1, InputString, Substrings(i - 1))
            
    Do While j > 0
        Result = Result + 1
        j = InStr(j + 1, InputString, Substrings(i - 1))
    Loop
         
Next i

StringCharCount_IncludeOverlap = Result

End Function

'****************************************************************************************************
Function StringCompare(a As String, b As String, Optional Limit As Integer = 10, Optional Mode As String = "Verbose") As String

'====================================================================================================
'Compares two strings and lists their differences, very raw so far
'Juraj Ahel, 2015-02-12, for comparing of protein sequences to find point mutations
'Last update 2015-02-12
'====================================================================================================

Dim i As Integer, j As Integer
Dim Result As String, s As String
Dim LA As Integer, Lb As Integer
Dim Counter As Integer: Counter = 0
Dim cA As String, cB As String

LA = Len(a): Lb = Len(b)

s = "; "

Select Case UCase(Mode)

Case "SHORT", "S"

Do
    i = i + 1
    cA = Mid(a, i, 1)
    cB = Mid(b, i, 1)
    
    If cA <> cB Then
        Counter = Counter + 1
        Result = Result & s & i
    End If
Loop Until i = LA Or i = Lb Or ((Counter > Limit) And (Limit > 0))


Case "VERBOSE", "V"
GoTo 50

Case Else
50
Do
    i = i + 1
    cA = Mid(a, i, 1)
    cB = Mid(b, i, 1)
    
    If cA <> cB Then
        Counter = Counter + 1
        Result = Result & s & i & "(" & cA & ">" & cB & ")"
    End If
Loop Until i = LA Or i = Lb Or ((Counter > Limit) And (Limit > 0))

If Counter = 0 And LA = Lb Then
    Result = "Exact Copy!"
    GoTo 99
End If

End Select

If LA <> Lb Then Result = Result & s & "LenDiff=" & LA - Lb

If Len(Result) > 0 Then Result = Right(Result, Len(Result) - Len(s))

If Counter > Limit And Limit > 0 Then Result = "Threshold (" & Limit & ") reached!"

99 StringCompare = Result

End Function
