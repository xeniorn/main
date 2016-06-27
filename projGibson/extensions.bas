Attribute VB_Name = "extensions"
Sub Testy()

Dim a As String, b As String, te As String, c As String, d As String

a = DNALongestORF(Range("AC17"))
'a = DNALongestORF("AAAAAAAAATGGGGGGGGGGGGGGGGTGACCC")
k = Len(a)
hh = DNATranslate(Range("AK17"))
End Sub
'****************************************************************************************************
Function DNALongestORF(Sequence As String, Optional Circular As Boolean = True, _
                        Optional Skip As Integer = 1) As String
'====================================================================================================
'Finds the longest ORF in a DNA sequence, read in forward direction, assuming it's circular by default
'Juraj Ahel, 2015-09-29
'Last update 2016-01-14
'====================================================================================================

Const MinimumORF As Long = 0

Dim TempStart As Long, TempEnd As Long, BestStart As Long
Dim SequenceLength As Long
Dim MaxEnd As Long

Dim BestLength As Long, CurrentLength As Long

SequenceLength = Len(Sequence)

If Circular Then
    'Sequence = Right(Sequence, SequenceLength \ 2 + 1) & _
    '            Sequence & _
    '        Left(Sequence, SequenceLength \ 2 + 1)
    Sequence = Sequence & Left(Sequence, SequenceLength - 1)
End If

TempStart = 0
BestStart = 0
Do

    TempStart = InStr(TempStart + 1, Sequence, "ATG") 'beginning of start codon
    TempEnd = TempStart
    MaxEnd = TempStart + SequenceLength - 3 'the start of end codon must
    
    If MaxEnd > Len(Sequence) Then MaxEnd = Len(Sequence) - 2
    j = 0
    Do
        TempEnd = TempEnd + 3
        Codon = Mid(Sequence, TempEnd, 3)
    Loop Until Codon = "TGA" Or Codon = "TAA" Or Codon = "TAG" Or TempEnd > MaxEnd
    
    CurrentLength = TempEnd - TempStart
    If CurrentLength > BestLength And CurrentLength <= SequenceLength Then
        BestLength = CurrentLength
        BestStart = TempStart
    End If

Loop Until TempStart = 0 Or TempStart > (SequenceLength - MinimumORF)

DNALongestORF = Mid(Sequence, BestStart, BestLength + 3)

End Function

'****************************************************************************************************
Function DNAGibsonLigation(ParamArray DNAList() As Variant) As String

'====================================================================================================
'Ligates a number of DNA sequences, requiring the final product to be circular
'Juraj Ahel, 2015-09-27
'Last update 2015-09-28
'====================================================================================================
'demonstrated to work 2015-09-28 on pJA1K and PLS46 (Mys1b in pFastBAC1 from 1-2, 3-5, 6-7, DF14)

Const MinOverlap = 15           'overlap should be at least this
Const MaxOverlapCheck = 250     'max meaningful to check, could be arbitrarily long code-wise, but no reason
Const MinTm = 48                'Tm should be at least this

Dim FragmentCount As Integer
Dim OverlapLength As Integer
Dim TempResult As String
Dim i As Integer, j As Integer
Dim Tm As Double

FragmentCount = 1 + UBound(DNAList) - LBound(DNAList)

TempResult = DNAList(0)


For i = 0 To FragmentCount - 1
    j = MaxOverlapCheck
    Do While (Right(DNAList(i), j) <> Left(DNAList((i + 1) Mod FragmentCount), j))
        j = j - 1
    Loop
    OverlapLength = j
    Tm = OligoTm(Right(DNAList(i), j))
    If (OverlapLength < MinOverlap) Or (Tm < MinTm) Then
        TempResult = "#ERROR! Overlap " & (1 + i) & "-" & (1 + ((i + 1) Mod FragmentCount)) & " faulty!"
        GoTo 999
    Else
        DNAList(i) = Left(DNAList(i), Len(DNAList(i)) - OverlapLength)
    End If
Next i

TempResult = Join(DNAList, "")

999 DNAGibsonLigation = TempResult

End Function

'****************************************************************************************************
Function DNAGCContent(Sequence As String) As Double
'====================================================================================================
'Calculates GC % as sum(G+C) / total length
'Juraj Ahel, 2015-09-28, for general purposes
'Last update 2015-09-28
'====================================================================================================

DNAGCContent = StringCharCount(UCase(Sequence), "G", "C", "S") / Len(Sequence)

End Function

'****************************************************************************************************
Function StringJoin(RangeToJoin As Range, Optional Separator As String = "", Optional Direction As Integer) As String

'====================================================================================================
'Joins all the cell values in an array as strings
'Juraj Ahel, 2015-02-16, for general purposes
'Last update 2015-04-13
'====================================================================================================

Dim TempString As String
Dim cell As Range

For Each cell In RangeToJoin
    TempString = TempString & cell.Value & Separator
Next cell

StringJoin = TempString

End Function


'****************************************************************************************************
Function PCRWithOverhangs(Template As String, _
                    ForwardPrimer As String, ReversePrimer As String, _
                    Optional Circular = False, _
                    Optional Perfect = True, _
                    Optional IgnoreBestMatch = False, _
                    Optional Details = False, _
                    Optional MinimalOverlap = 15 _
                    ) As String

'====================================================================================================
'Allows for PCR using primers that have 5' overhangs, introducing extra nucleotides at the
'termini of the amplified DNA sequence. Does some basic checks for whether it would work -
'it needs at least 15 nt overlap, a single most prominent binding site, and doesn't allow
'insertions or deletions after the annealing locus
'Juraj Ahel, 2015-06-14, to be able to quickly generate fragments for in-silico cloning
'Last update 2015-06-29
'====================================================================================================

Dim OverhangF As String, OverhangR As String
Dim OverlapF As String, OverlapR As String
Dim ReversePrimerRC As String, TempFrag As String

Dim NCheck As Integer: NCheck = 3
Dim ErrorMsg() As String
Dim CtrlF() As Boolean, CtrlR() As Boolean
Dim ErrMF() As String, ErrMR() As String
ReDim ErrMF(1 To NCheck)
ReDim ErrMR(1 To NCheck)
ReDim ErrorMsg(1 To NCheck)
ReDim CtrlF(1 To NCheck)
ReDim CtrlR(1 To NCheck)
Dim CtrlSum As Integer

ErrorMsg(1) = "no overlap"
ErrorMsg(2) = "overlap <" & MinimalOverlap & " bp"
ErrorMsg(3) = "insertion after overlap"

ReversePrimerRC = DNAReverseComplement(ReversePrimer)

'if stringent, looks for best match, otherwise looks for maximum overlap at terminus
If Not IgnoreBestMatch Then
    OverlapF = StringFindOverlap(ForwardPrimer, Template)
    OverlapR = StringFindOverlap(ReversePrimerRC, Template)
Else
    i = 0
    Do
        i = i + 1
        TempFrag = Right(ForwardPrimer, i)
    Loop Until InStr(1, Template, TempFrag) = 0 Or i = Len(ForwardPrimer)
    OverlapF = Right(ForwardPrimer, i - 1)
    i = 0
    Do
        i = i + 1
        TempFrag = Left(ReversePrimerRC, i)
    Loop Until InStr(1, Template, TempFrag) = 0 Or i = Len(ReversePrimer)
    OverlapR = Left(ReversePrimerRC, i - 1)
End If

'is there overlap at all?
If Left(OverlapF, 2) = "#!" Then CtrlF(1) = True
If Left(OverlapR, 2) = "#!" Then CtrlR(1) = True

'is the overlap at least MinimalOverlap bp?
If Len(OverlapF) < MinimalOverlap Then CtrlF(2) = True
If Len(OverlapR) < MinimalOverlap Then CtrlR(2) = True

'is the overlapping region at the 3' end of the primer?
If Right(ForwardPrimer, Len(OverlapF)) <> OverlapF Then CtrlF(3) = True
If Left(ReversePrimerRC, Len(OverlapR)) <> OverlapR Then CtrlR(3) = True

For i = 1 To NCheck
    CtrlSum = CtrlSum + CtrlF(i) + CtrlR(i)
    If CtrlF(i) Then ErrMF(i) = ErrorMsg(i)
    If CtrlR(i) Then ErrMR(i) = ErrorMsg(i)
Next i

'TRUE IS -1, NOT 1 AS INTEGER!!!!
If CtrlSum < 0 Then
    TempResult = Abs(CtrlSum) & "#!:"
    TempResult = TempResult & " for: " & Join(ErrMF, ", ")
    TempResult = TempResult & " rev: " & Join(ErrMR, ", ")
    GoTo 999
End If

OverhangF = Left(ForwardPrimer, Len(ForwardPrimer) - Len(OverlapF))
If Len(OverlapR) < Len(ReversePrimer) Then
    OverhangR = DNAReverseComplement(Left(ReversePrimer, Len(ReversePrimer) - Len(OverlapR)))
End If

If Not Details Then
    TempResult = PCRSimulate(Template, OverlapF, DNAReverseComplement(OverlapR), Circular, Perfect)
    TempResult = OverhangF & TempResult & OverhangR
Else
    TempResult = "F:" & OligoTm(OverlapF) & " °C, " & Len(OverlapF)
    TempResult = TempResult & " R:" & OligoTm(OverlapR) & " °C, " & Len(OverlapR)
End If

999 PCRWithOverhangs = TempResult

End Function

Sub SingleColumn()

RC = Selection.Rows.Count
CC = Selection.Columns.Count

Dim InputData(), OutList()
ReDim InputData(1 To RC, 1 To CC)
ReDim OutList(1 To RC * CC, 1 To 1)

InputData = Selection.Value

For i = 1 To CC
    For j = 1 To RC
        OutList((i - 1) * RC + j, 1) = InputData(j, i)
    Next j
Next i

Range(Selection.Offset(RC + 1, CC + 1).Resize(RC * CC, 1)).Value = OutList

End Sub


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
Dim NumberOfSubstractions As Integer
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
'****************************************************************************************************
Sub SwapValue(a As Variant, b As Variant)

'====================================================================================================
'Swaps two values of any type variable
'Juraj Ahel, 2015-04-30, for general purposes
'Last update 2015-04-30
'====================================================================================================

Dim c

c = a
a = b
b = c

End Sub
'****************************************************************************************************
Function StringFindOverlap(Probe As String, Target As String)

'====================================================================================================
'Finds the (largest) continuous perfectoverlap between two strings
'Juraj Ahel, 2015-04-30, for general purposes
'Last update 2015-04-30
'====================================================================================================

Dim ProbeLength As Long, TargetLength As Long
Dim Results() As Long

ProbeLength = Len(Probe)
TargetLength = Len(Target)

If ProbeLength > TargetLength Then
    Call SwapValue(Probe, Target)
    Call SwapValue(ProbeLength, TargetLength)
End If
    
wStart = ProbeLength

If wStart = 0 Then
    TempResult = "Zero-string probe or target."
    GoTo 999
End If

'- if I want to map them all
'ReDim Results(1 To wStart, 1 To wStart)

'- if I want to extract the longest ones only
ReDim Results(1 To wStart)

Dim i As Long, j As Long, k As Long, w As Long
Dim TempProbe As String
Dim FoundOverlap As Boolean

w = wStart

Do
    k = 0
    
    For i = 1 To 1 + (wStart - w)
    
        TempProbe = Mid(Probe, i, w)
        
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
    
    w = w - 1
    
Loop Until k <> 0 Or w = 0

OverlapWidth = w + 1

Dim TempResultAsStrings() As String

Select Case k
    Case 0
        TempResult = "#! No overlap found."
    Case 1
        TempResult = Mid(Target, Results(1), OverlapWidth)
    Case Is > 1
        ReDim TempResultAsStrings(1 To k)
        For i = 1 To k
            TempResultAsStrings(i) = CStr(Results(i))
        Next i
    
        
        TempResult = "Multiple equivalent results of length " _
                    & OverlapWidth & " at positions: " _
                    & Join(TempResultAsStrings, ";")
End Select

999 StringFindOverlap = TempResult

End Function


'****************************************************************************************************
Sub RangeColumnInvert()

'====================================================================================================
'Inverts the row order within the selected columns
'Copies formulas "stupidly", raw-copy style, so "A2 + A3" will stay as such and won't switch to "A3 + A4"
'Juraj Ahel, 2015-04-17, for general purposes
'Last update 2015-04-17
'====================================================================================================

Dim TargetRange As Range
Dim NRow As Long, NColumn As Long
Dim RangeDataIn() As Variant, RangeDataOut() As Variant
Dim i As Long, j As Long, k As Long

Set TargetRange = Selection

NRow = TargetRange.Rows.Count
NColumn = TargetRange.Columns.Count

ReDim RangeDataIn(1 To NRow, 1 To NColumn)
ReDim RangeDataOut(1 To NRow, 1 To NColumn)

RangeDataIn = TargetRange.Formula

For i = 1 To NRow

    k = NRow - i + 1
    
    For j = 1 To NColumn
        RangeDataOut(k, j) = RangeDataIn(i, j)
    Next j
    
Next i

TargetRange.Formula = RangeDataOut

End Sub

'****************************************************************************************************
Function OptimizePrimer(TargetSequence As String, Optional TargetTm As Double = 60, Optional MinLength As Integer = 15) As String

'====================================================================================================
'Designs a simple primer for regular PCR amplification, trying to optimize the Tm and trying to
'keep the termini either G or C
'Always does a forward primer - do DNAReverseTranslate to Target to get the reverse. Might implement
'it as an option later
'In the future might be made more robust
'Juraj Ahel, 2015-03-24, general purposes
'Last update 2015-03-24
'====================================================================================================

Const NumberOfVariants = 40

Dim Result As String
Dim Tm As Double
Dim Length As Integer
Dim Score() As Double, MaxScore As Integer
Dim Variants() As String
Dim i As Integer, j As Integer
Dim PrimerStart As String, PrimerEnd As String

ReDim Score(1 To NumberOfVariants)
ReDim Variants(1 To NumberOfVariants)

j = 0
MaxScore = -30000

For i = 1 To NumberOfVariants

    Variants(i) = Left(TargetSequence, MinLength + i - 1)
    Score(i) = -((OligoTm(Variants(i)) - TargetTm)) ^ 2
    PrimerStart = Left(Variants(i), 1)
    PrimerEnd = Right(Variants(i), 1)
    If PrimerStart = "A" Or PrimerStart = "T" Then Score(i) = Score(i) - 4
    If PrimerEnd = "A" Or PrimerEnd = "T" Then Score(i) = Score(i) - 10
    If Score(i) > MaxScore Then
        MaxScore = Score(i)
        j = i
    End If

Next i

OptimizePrimer = Variants(j)

End Function

'****************************************************************************************************
Function PCRSimulate(Template As String, _
                    ForwardPrimer As String, ReversePrimer As String, _
                    Optional Circular = False, _
                    Optional Perfect = True _
                    ) As String

'====================================================================================================
'Simulates a PCR using selected primers. So far supports only perfect primers (no overhang, no mismatch)
'Can simulate PCR of circular templates
'Juraj Ahel, 2015-03-24, for Gibson assembly and general purposes
'Last update 2015-05-05
'====================================================================================================

Dim ErrorPrefix As String
ErrorPrefix = "#! "

Dim PrimerFCount As Integer, PrimerRCount As Integer
Dim Result As String

PrimerFCount = StringCharCount_IncludeOverlap(Template, ForwardPrimer, DNAReverseComplement(ForwardPrimer))
PrimerRCount = StringCharCount_IncludeOverlap(Template, DNAReverseComplement(ReversePrimer))

If PrimerFCount <> 1 Or PrimerRCount <> 1 Then

    If PrimerFCount > 1 Or PrimerRCount > 1 Then
        Result = "Primer target sites not unique: Forward: " & PrimerFCount & " Reverse: " & PrimerRCount
    ElseIf PrimerFCount = 0 And PrimerRCount = 0 Then
        Result = "No binding site found for either primer!"
    ElseIf PrimerFCount = 0 Then
        Result = "No binding site found for Forward primer."
    ElseIf PrimerRCount = 0 Then
        Result = "No binding site found for Reverse primer."
    End If
    
    Result = ErrorPrefix & Result
    
    GoTo 999
End If

Dim FSite As Integer, RSite As Integer, FLen As Integer, RLen As Integer
Dim Reverse As Boolean

Reverse = False
FSite = InStr(1, Template, ForwardPrimer)
RSite = InStr(1, Template, DNAReverseComplement(ReversePrimer))

'If circular, pretend it's linear that starts exactly where F primer starts
'and remap the indexing
If Circular Then
    Template = SubSequenceSelect(Template, FSite, Len(Template)) & _
                SubSequenceSelect(Template, 1, FSite - 1)
    RSite = RSite - FSite + 1
    FSite = 1
    If RSite < 1 Then RSite = Len(Template) + RSite
End If

'###correcting for if Forward primer and Reverse primer have been swapped
'If FSite = 0 Or RSite = 0 Then
'
'    ForwardPrimer = DNAReverseComplement(ForwardPrimer)
'    ReversePrimer = DNAReverseComplement(ReversePrimer)
'
'    FSite = InStr(1, Template, ForwardPrimer)
'    RSite = InStr(1, Template, DNAReverseComplement(ReversePrimer))
'
'    Reverse = True
'
'End If

FLen = Len(ForwardPrimer)
RLen = Len(ReversePrimer)

Result = ForwardPrimer & SubSequenceSelect(Template, FSite + FLen, RSite - 1) & DNAReverseComplement(ReversePrimer)

If Len(Result) < FLen + RLen Then Result = ErrorPrefix & "Primers too close."

If FSite > RSite Then Result = ErrorPrefix & "Reverse primer anneals upstream of Forward primer, check sequences."

999 PCRSimulate = Result

End Function

'****************************************************************************************************
Function SubSequenceSelect(InputString As String, StartIndex, EndIndex, Optional DNA As Boolean = False) As String

'====================================================================================================
'Like "Mid" function, but taking indices as arguments, not start index + length
'When end < start, gives reverse
'If DNA is true, the reverse complement for end < start
'Juraj Ahel, 2015-02-16, for vector subselection and general purposes
'Last update 2015-02-16
'====================================================================================================

Dim TempString As String

If StartIndex <= EndIndex Then

    TempString = Mid(InputString, StartIndex, EndIndex - StartIndex + 1)

Else
    
    
    TempString = Mid(InputString, EndIndex, StartIndex - EndIndex + 1)
    
    Dim N As Integer, i As Integer
    Dim TempStringChars() As String
    
    N = Len(TempString)
    ReDim TempStringChars(1 To N)
    
    If DNA Then
        TempString = DNAReverseComplement(TempString)
    Else
        For i = 1 To N: TempStringChars(i) = Mid(TempString, N - i + 1, 1): Next i
        TempString = Join(TempStringChars, "")
    End If
    
End If

SubSequenceSelect = TempString

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

Dim StringLength As Integer
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
'Last update 2015-03-24
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
Sub testtt()

a = OligoTm("AAAAAAAA")
b = OligoTm("GAAAAAAC")

c = OligoTm1("AAAAAAAA")
d = OligoTm1("GAAAAAAC")

End Sub


'****************************************************************************************************
'****************************************************************************************************
Function OligoTm( _
                 Sequence As String, _
                 Optional EffectiveMonovalentCation_mM As Double = 50, _
                 Optional OligoConcentration_nM As Double = 500, _
                 Optional Mode As String = "DNA", _
                 Optional TargetSequence As String = "" _
                ) As Double

'====================================================================================================
'Returns DNA melting temperature using Nearest Neighbour thermodynamics (NN)
'Works as EMBOSS dan, except it doesn't implement % formamide / DMSO and mismatches yet
'Replicated from Florian Weissman's script for Gibson assembly | originaly by Sebastina Bassi
'Juraj Ahel, 2015-02-11, for more proper oligo Tm calculations than with the older naive algorithm
'Last update 2015-03-24
'====================================================================================================
'still lacks additional energy by terminal GC or AT on either side (can take also from PrecisePrimer manual)
'for this, I would first implement the possibility of selecting the subsequence that actually anneals, + mismatches
'Also, I would like to implement the effect of Magnesium (and other divalent) ions, and possibly DMSO

Dim Pairs() As Variant, dHTable() As Variant, dSTable() As Variant
Pairs = Array("AA", "TT", "AT", "TA", "CA", "TG", "GT", "AC", "CT", "AG", "GA", "TC", "CG", "GC", "GG", "CC")
dHTable = Array(7.9, 7.9, 7.2, 7.2, 8.5, 8.5, 8.4, 8.4, 7.8, 7.8, 8.2, 8.2, 10.6, 9.8, 8, 8)
dSTable = Array(22.2, 22.2, 20.4, 21.3, 22.7, 22.7, 22.4, 22.4, 21, 21, 22.2, 22.2, 27.2, 24.4, 19.9, 19.9)

Sequence = UCase(Sequence)

Dim i As Integer
Dim Seq() As String, Seqp() As String
Dim N As Integer
Dim salt As Double, DNAc As Double
Dim r As Double, LogDNA As Double

Dim dH As Double, dS As Double
Dim Pair As String, PairCount As Integer

salt = EffectiveMonovalentCation_mM / 1000#
DNAc = OligoConcentration_nM / 1000000000#
N = Len(Sequence)

dG = 0: dS = 0

For i = 0 To 15
    Pair = Pairs(i)
    PairCount = StringCharCount_IncludeOverlap(Sequence, Pairs(i))
    If PairCount > 0 Then
        dH = dH + PairCount * dHTable(i)
        dS = dS + PairCount * dSTable(i)
    End If
    Counter = Counter + PairCount
Next i
    
r = 1.98717

'### Florian's version
'LogDNA = r * Ln(DNAc / 4)
    
'### Version from PrecisePrimer (different assumptions, focusing on the initial state where [primer]>>[template]
'### and also additional effect of terminal nucleotides (from SantaLucia et al.)
LogDNA = r * Ln(DNAc)
'Dim Termini As String: Termini = Left(Sequence, 1) & Right(Sequence, 1)
Dim STerminal As Double, HTerminal As Double

'HTerminal = 100 * StringCharCount(Termini, "G", "C") + 2300 * StringCharCount(Termini, "A", "T")
'STerminal = -2.8 * StringCharCount(Termini, "G", "C") + 4.1 * StringCharCount(Termini, "A", "T")
HTerminal = 0: STerminal = 0

    
Dim Entropy As Double, Enthalpy As Double, Tm As Double

'Entropy = -10.8 - dS + 0.368 * (N - 1) * Lg(salt)
Entropy = -10.8 - dS + 0.368 * (N - 1) * Lg(salt) + STerminal
Enthalpy = -dH * 1000 + HTerminal

Tm = Enthalpy / (Entropy + LogDNA) - 273.15              'Lol, error was that it said "275.15".... -.-'

OligoTm = Round(Tm, 1)

End Function


'****************************************************************************************************
Function DNAReverseComplement(InputSequence As String) As String

'====================================================================================================
'Outputs a DNA reverse complement of a given input sequence
'Juraj Ahel, 2015-02-04, for checking primers
'Last update 2015-02-04
'====================================================================================================
'So far, always UPPERCASE output. Non-ACGT are preserved.

Dim i As Integer, StringLength As Integer
Dim OutputSequence() As String

StringLength = Len(InputSequence)
ReDim OutputSequence(1 To StringLength)
 InputSequence = UCase(InputSequence)

For i = 1 To StringLength
    
    j = StringLength - i + 1
    
    Select Case Mid(InputSequence, i, 1)
        Case "A": OutputSequence(j) = "T"
        Case "C": OutputSequence(j) = "G"
        Case "G": OutputSequence(j) = "C"
        Case "T": OutputSequence(j) = "A"
        Case Else: OutputSequence(j) = Mid(InputSequence, i, 1)
    End Select
Next i

DNAReverseComplement = Join(OutputSequence, "")

End Function



'****************************************************************************************************
Function DNATranslate(ByVal InputSequence As String) As String

'====================================================================================================
'Translates a DNA sequence to a protein sequence, using standard code
'sequence is given as a single-line string, output is also a single-line string
'other date formats possibly to be added
'Juraj Ahel, 2015-01-25, for general purposes
'Last update 2016-01-15
'====================================================================================================
    
    Dim i As Integer, SequenceLength As Long, ProteinLength As Long
    Dim AminoAcid As String, OutputSequence As String, Codon As String
    Dim AminoAcids() As String
    
    SequenceLength = Len(InputSequence)
    OutputSequence = ""
    AminoAcid = ""
    i = 0
    
    InputSequence = Replace(UCase(InputSequence), "U", "T")
    
    ProteinLength = SequenceLength \ 3
    
    If SequenceLength Mod 3 <> 0 Then
        OutputSequence = "Input is not a valid coding sequence (len = 3k, k€N)"
        GoTo 99
    End If
    
    ReDim AminoAcids(1 To ProteinLength)
    
    For i = 1 To ProteinLength
    
        Codon = Mid(InputSequence, 3 * i - 2, 3)
        
        Select Case Codon
            Case "GCA", "GCC", "GCG", "GCT"
            AminoAcid = "A"
            Case "AGA", "AGG", "CGA", "CGC", "CGG", "CGT"
            AminoAcid = "R"
            Case "AAC", "AAT"
            AminoAcid = "N"
            Case "GAC", "GAT"
            AminoAcid = "D"
            Case "TGC", "TGT"
            AminoAcid = "C"
            Case "CAA", "CAG"
            AminoAcid = "Q"
            Case "GAA", "GAG"
            AminoAcid = "E"
            Case "GGA", "GGC", "GGG", "GGT"
            AminoAcid = "G"
            Case "CAC", "CAT"
            AminoAcid = "H"
            Case "ATA", "ATC", "ATT"
            AminoAcid = "I"
            Case "CTA", "CTC", "CTG", "CTT", "TTA", "TTG"
            AminoAcid = "L"
            Case "AAA", "AAG"
            AminoAcid = "K"
            Case "ATG"
            AminoAcid = "M"
            Case "TTC", "TTT"
            AminoAcid = "F"
            Case "CCA", "CCC", "CCG", "CCT"
            AminoAcid = "P"
            Case "AGC", "AGT", "TCA", "TCC", "TCG", "TCT"
            AminoAcid = "S"
            Case "ACA", "ACC", "ACG", "ACT"
            AminoAcid = "T"
            Case "TGG"
            AminoAcid = "W"
            Case "TAC", "TAT"
            AminoAcid = "Y"
            Case "GTA", "GTC", "GTG", "GTT"
            AminoAcid = "V"
            Case "TAA", "TAG", "TGA"
            AminoAcid = "*"
            Case Else
            AminoAcid = "X"
        End Select
        
        OutputSequence = OutputSequence & AminoAcid
        AminoAcids(i) = AminoAcid
        
    Next i
    
'99     DNATranslate = OutputSequence
99     DNATranslate = Join(AminoAcids, "")

End Function


