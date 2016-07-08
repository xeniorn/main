Attribute VB_Name = "modCodeCollection"
'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2016-03-09
'Last update 2016-03-09
'====================================================================================================


Sub testoror()

Dim dict As Dictionary


End Sub

Sub Table96x1To8x12()

Set InputRange = Selection

For k = 1 To InputRange.Columns.Count

Set OutputRange = InputRange.Offset(98 + (k - 1) * 9, 0).Resize(8, 12)

InputTable = InputRange.Columns(k).Value
Dim OutputTable(1 To 8, 1 To 12)

For i = 1 To 8
    For j = 1 To 12
        OutputTable(i, j) = InputTable(8 * (j - 1) + i, 1)
    Next j
Next i

OutputRange.Value = OutputTable

Next k


End Sub

'****************************************************************************************************
Function DNALongestORF(Sequence As String, Optional Circular As Boolean = True, _
                        Optional Skip As Integer = 1) As String
'====================================================================================================
'Finds the longest ORF in a DNA sequence, read in forward direction, assuming it's circular by default
'Juraj Ahel, 2015-09-29
'Last update 2015-11-06
'====================================================================================================

Const MinimumORF As Long = 0

Dim TempStart As Long, TempEnd As Long, BestStart As Long
Dim SequenceLength As Long
Dim MaxEnd As Long

Dim BestLength As Long, CurrentLength As Long

SequenceLength = Len(Sequence)

If Circular Then
    Sequence = Right(Sequence, SequenceLength \ 2 + 1) & _
                Sequence & _
            Left(Sequence, SequenceLength \ 2 + 1)
End If

TempStart = 0
BestStart = 0
Do

    TempStart = InStr(TempStart + 1, Sequence, "ATG")
    TempEnd = TempStart
    MaxEnd = TempStart + SequenceLength - 3
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
Function DNAGCContent(Sequence As String) As Double
'====================================================================================================
'Calculates GC % as sum(G+C) / total length
'Juraj Ahel, 2015-09-28, for general purposes
'Last update 2015-09-28
'====================================================================================================

DNAGCContent = StringCharCount(UCase(Sequence), "G", "C", "S") / Len(Sequence)

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
Dim tempResult As String
Dim i As Integer, j As Integer
Dim Tm As Double

FragmentCount = 1 + UBound(DNAList) - LBound(DNAList)

tempResult = DNAList(0)


For i = 0 To FragmentCount - 1
    j = MaxOverlapCheck
    Do While (Right(DNAList(i), j) <> Left(DNAList((i + 1) Mod FragmentCount), j))
        j = j - 1
    Loop
    OverlapLength = j
    Tm = OligoTm(Right(DNAList(i), j))
    If (OverlapLength < MinOverlap) Or (Tm < MinTm) Then
        tempResult = "#ERROR! Overlap " & (1 + i) & "-" & (1 + ((i + 1) Mod FragmentCount)) & " faulty!"
        GoTo 999
    Else
        DNAList(i) = Left(DNAList(i), Len(DNAList(i)) - OverlapLength)
    End If
Next i

tempResult = Join(DNAList, "")

999 DNAGibsonLigation = tempResult

End Function

'****************************************************************************************************
Function DNAReindex(DNASequence As String, NewStartBase As Long) As String

'====================================================================================================
'Reindexes a circular DNA sequence
'Juraj Ahel, 2015-09-27
'Last update 2015-09-28
'====================================================================================================

Dim SeqLength As Long, Offset As Long

SeqLength = Len(DNASequence)

Offset = NewStartBase - 1

Select Case Offset
    Case 0
        DNAReindex = DNASequence
    Case Is > 0
        DNAReindex = Right(DNASequence, SeqLength - Offset) & Left(DNASequence, Offset)
    Case Else
        DNAReindex = Right(DNASequence, -Offset) & Right(DNASequence, SeqLength + Offset)
End Select

End Function

'****************************************************************************************************
Sub GenerateCodingFromProtein()

'====================================================================================================
'Generates all possible nucleotide sequences that produce a given protein sequence
'
'Juraj Ahel, 2015-09-24, for finding ideal Gibson overlaps
'Last update 2015-09-25
'====================================================================================================

Dim Codons(1 To 20, 1 To 6) As String

Codons(1, 1) = "GCA": Codons(16, 5) = "TCG"
Codons(2, 1) = "TGC": Codons(1, 2) = "GCC"
Codons(3, 1) = "GAC": Codons(2, 2) = "TGT"
Codons(4, 1) = "GAA": Codons(3, 2) = "GAT"
Codons(5, 1) = "TTC": Codons(4, 2) = "GAG"
Codons(6, 1) = "GGA": Codons(5, 2) = "TTT"
Codons(7, 1) = "CAC": Codons(6, 2) = "GGC"
Codons(8, 1) = "ATA": Codons(7, 2) = "CAT"
Codons(9, 1) = "AAA": Codons(8, 2) = "ATC"
Codons(10, 1) = "CTA": Codons(9, 2) = "AAG"
Codons(11, 1) = "ATG": Codons(10, 2) = "CTC"
Codons(12, 1) = "AAC": Codons(12, 2) = "AAT"
Codons(13, 1) = "CCA": Codons(13, 2) = "CCC"
Codons(14, 1) = "CAA": Codons(14, 2) = "CAG"
Codons(15, 1) = "AGA": Codons(15, 2) = "AGG"
Codons(16, 1) = "AGC": Codons(16, 2) = "AGT"
Codons(17, 1) = "ACA": Codons(17, 2) = "ACC"
Codons(18, 1) = "GTA": Codons(18, 2) = "GTC"
Codons(19, 1) = "TGG": Codons(20, 2) = "TAT"
Codons(20, 1) = "TAC": Codons(1, 4) = "GCT"
Codons(1, 3) = "GCG": Codons(6, 4) = "GGT"
Codons(6, 3) = "GGG": Codons(10, 4) = "CTT"
Codons(8, 3) = "ATT": Codons(13, 4) = "CCT"
Codons(10, 3) = "CTG": Codons(15, 4) = "CGC"
Codons(13, 3) = "CCG": Codons(16, 4) = "TCC"
Codons(15, 3) = "CGA": Codons(17, 4) = "ACT"
Codons(16, 3) = "TCA": Codons(18, 4) = "GTT"
Codons(17, 3) = "ACG": Codons(10, 6) = "TTG"
Codons(18, 3) = "GTG": Codons(15, 6) = "CGT"
Codons(10, 5) = "TTA": Codons(16, 6) = "TCT"
Codons(15, 5) = "CGG"


Dim ProteinSequence As String
Dim NumberOfVariants As Integer, ProteinSequenceLength As Integer
Dim Variants() As String
Dim AminoAcidIndex() As Integer
Dim Multiplicity() As Integer
Dim counter As Long, CumulativeIndex As Long
Dim Codon As String
Dim CodonIndex As Integer

'ProteinSequence = InputBox("Gimme da Sequence:")
ProteinSequence = CStr(Selection.Resize(1, 1))

ProteinSequenceLength = Len(ProteinSequence)

ReDim AminoAcidIndex(1 To ProteinSequenceLength)
ReDim Multiplicity(1 To ProteinSequenceLength)

NumberOfVariants = 1 ^ StringCharCount(ProteinSequence, "M", "W") * _
                2 ^ StringCharCount(ProteinSequence, "C", "D", "E", "F", "H", "K", "N", "Q", "Y") * _
                3 ^ StringCharCount(ProteinSequence, "I") * _
                4 ^ StringCharCount(ProteinSequence, "A", "G", "P", "T", "V") * _
                6 ^ StringCharCount(ProteinSequence, "L", "R", "S")
        
ReDim Variants(1 To NumberOfVariants, 1 To 1)
        
For i = 1 To ProteinSequenceLength
    Select Case Mid(ProteinSequence, i, 1)
        Case "M", "W"
            Multiplicity(i) = 1
        Case "C", "D", "E", "F", "H", "K", "N", "Q", "Y"
            Multiplicity(i) = 2
        Case "I"
            Multiplicity(i) = 3
        Case "A", "G", "P", "T", "V"
            Multiplicity(i) = 4
        Case "L", "R", "S"
            Multiplicity(i) = 6
    End Select
Next i

For i = 1 To ProteinSequenceLength
    Select Case Mid(ProteinSequence, i, 1)
        Case "A": AminoAcidIndex(i) = 1
        Case "C": AminoAcidIndex(i) = 2
        Case "D": AminoAcidIndex(i) = 3
        Case "E": AminoAcidIndex(i) = 4
        Case "F": AminoAcidIndex(i) = 5
        Case "G": AminoAcidIndex(i) = 6
        Case "H": AminoAcidIndex(i) = 7
        Case "I": AminoAcidIndex(i) = 8
        Case "K": AminoAcidIndex(i) = 9
        Case "L": AminoAcidIndex(i) = 10
        Case "M": AminoAcidIndex(i) = 11
        Case "N": AminoAcidIndex(i) = 12
        Case "P": AminoAcidIndex(i) = 13
        Case "Q": AminoAcidIndex(i) = 14
        Case "R": AminoAcidIndex(i) = 15
        Case "S": AminoAcidIndex(i) = 16
        Case "T": AminoAcidIndex(i) = 17
        Case "V": AminoAcidIndex(i) = 18
        Case "W": AminoAcidIndex(i) = 19
        Case "Y": AminoAcidIndex(i) = 20
    End Select
Next i

CumulativeIndex = 1

For i = 1 To ProteinSequenceLength

    For counter = 1 To NumberOfVariants
        'CodonIndex = 1 + Counter Mod Multiplicity(i)
        CodonIndex = 1 + ((counter - 1) \ CumulativeIndex) Mod Multiplicity(i)
        Codon = Codons(AminoAcidIndex(i), CodonIndex)
        Variants(counter, 1) = Variants(counter, 1) & Codon
    Next counter
    
    CumulativeIndex = Multiplicity(i) * CumulativeIndex

Next i

Selection.Offset(1, 0).Resize(NumberOfVariants, 1).Value = Variants

End Sub

'****************************************************************************************************
Sub ExportSeqToTXT()

'====================================================================================================
'Exports cell column pairs formated as [HEADER][SEQUENCE] in simple FASTA format
'under a file name corresponding to the header
'
'Juraj Ahel, 2015-02-10, for exporting sequences, to have an external database
'Last update 2015-08-26
'====================================================================================================

Dim FilePath As String, OutputFile As String
Dim DataSource As Range
Dim HeaderLine As String
Dim Sequence As String

FilePath = "C:\Excel_outputs\Sequences\"

Set DataSource = Selection

For i = 1 To DataSource.Rows.Count
    HeaderLine = ">" & CStr(DataSource(i, 1).Value)
    Sequence = DataSource(i, 2).Value
    OutputFile = FilePath & CStr(DataSource(i, 1).Value) & "_seq.txt"
    Call ExportDataToTextFile(HeaderLine & vbCrLf & Sequence, OutputFile)
Next i

End Sub

'****************************************************************************************************
Function AnnotateMutationsManual(ReferenceSequence As String, ResidueIndex As Long, MutationType As String, _
                                   ResultNucleotide As String, Optional ProteinAsWell As Boolean = True) As String
                                   
'====================================================================================================
'Annotates mutations in DNA in a standard way, using a table of inputs consisting of
'residue index of mutation, type (del, ins, sub), resulting nucleotide (for ins/sub)
'gives nucleotide annotation, and optionally protein one
'
'Juraj Ahel, 2015-07-25
'Last update 2015-07-25
'====================================================================================================
'so far, handles only point mutations, no indels, range-deletions, and such things
'also, no mutations of terminus and start methionine


Dim SequenceLength As Long
Dim ResIndexP As Long
Dim RefSeqArray(), CurrentArray() As String
Dim TargetSequence As String
Dim TranslationWT As String, TranslationMUT As String
Dim i As Long

SequenceLength = Len(ReferenceSequence)

'ReDim RefSeqArray(1 To SequenceLength)
'ReDim CurrentArray(1 To SequenceLength)

'For i = 1 To SequenceLength
    'RefSeqArray(i) = Mid(ReferenceSequence, i, 1)
    'CurrentArray(i) = RefSeqArray(i)
'Next i

MutationType = UCase(MutationType)

Select Case MutationType
    Case "DEL", "D", "DELETION"
        
        TargetSequence = Left(ReferenceSequence, ResidueIndex - 1) & _
                         Right(ReferenceSequence, SequenceLength - ResidueIndex)
        AnnotationNucleotide = "c." & ResidueIndex & "del" & Mid(ReferenceSequence, ResidueIndex, 1)
        
        ResIndexP = Int((ResidueIndex + 2) / 3)
        
        TranslationWT = DNATranslate(ReferenceSequence)
        TranslationMUT = DNATranslate(TargetSequence, True)
        
        AnnotationProtein = "p." & Mid(TranslationWT, ResIndexP, 1) & ResIndexP & _
                            Mid(TranslationMUT, ResIndexP, 1) & "fs" & "*" & _
                            (InStr(ResIndexP, TranslationMUT, "*") - ResIndexP)
    Case "INS", "I", "INSERTION", "INSERT"
        
        TargetSequence = Left(ReferenceSequence, ResidueIndex) & ResultNucleotide & _
                         Right(ReferenceSequence, SequenceLength - ResidueIndex)
        AnnotationNucleotide = "c." & ResidueIndex & "ins" & ResultNucleotide
        
        ResIndexP = Int((ResidueIndex + 2) / 3)
        
        TranslationWT = DNATranslate(ReferenceSequence)
        TranslationMUT = DNATranslate(TargetSequence, True)
        
        AnnotationProtein = "p." & Mid(TranslationWT, ResIndexP, 1) & ResIndexP & _
                            Mid(TranslationMUT, ResIndexP, 1) & "fs" & "*" & _
                            (InStr(ResIndexP, TranslationMUT, "*") - ResIndexP)
                         
    Case "SUB", "SUBSTITUTION", "S"
        
        TargetSequence = Left(ReferenceSequence, ResidueIndex - 1) & ResultNucleotide & _
                         Right(ReferenceSequence, SequenceLength - ResidueIndex)
        AnnotationNucleotide = "c." & ResidueIndex & Mid(ReferenceSequence, ResidueIndex, 1) & _
                                ">" & ResultNucleotide
        
        ResIndexP = Int((ResidueIndex + 2) / 3)
        
        TranslationWT = DNATranslate(ReferenceSequence)
        TranslationMUT = DNATranslate(TargetSequence)
        
        If Mid(TranslationWT, ResIndexP, 1) = Mid(TranslationMUT, ResIndexP, 1) Then
            AnnotationProtein = "p.="
        Else
            AnnotationProtein = "p." & Mid(TranslationWT, ResIndexP, 1) & ResIndexP & _
                            Mid(TranslationMUT, ResIndexP, 1)
        End If
        
    Case Else
End Select
                            

If ProteinAsWell Then
    AnnotateMutationsManual = AnnotationNucleotide & " (" & AnnotationProtein & ")"
Else
    AnnotateMutationsManual = AnnotationNucleotide
End If
                                   
End Function

'****************************************************************************************************
Function GCRich(InputSequence As String, GCType, HalfWindowSize As Integer) As String

'====================================================================================================
'
'Juraj Ahel, 2015-07-09
'Last update 2015-07-09
'====================================================================================================

Dim SequenceLength As Long
Dim StartIndex As Integer, EndIndex As Integer
Dim GCRichness() As Double
Dim GCRichnessIndex() As String
Dim TempGCRich As Integer

Dim CutoffRich As Double, CutoffPoor As Double

CutoffRich = 0.55
CutoffPoor = 0.45

SequenceLength = Len(InputSequence)

ReDim GCRichness(1 To SequenceLength)
ReDim GCRichnessIndex(1 To SequenceLength)

For i = HalfWindowSize + 1 To SequenceLength - HalfWindowSize
    StartIndex = i - HalfWindowSize
    EndIndex = i + HalfWindowSize
    TempGC = StringCharCount_IncludeOverlap(SubSequenceSelect(InputSequence, StartIndex, EndIndex), "G", "C")
    GCRichness(i) = TempGC / HalfWindowSize
Next i

For i = 1 To HalfWindowSize
    GCRichness(i) = 0.5
    GCRichness(SequenceLength - i + 1) = 0.5
Next i

For i = 1 To SequenceLength
    If GCRichness(i) > CutoffRich Then
        GCRichnessIndex(i) = "2"
    Else
        If GCRichness(i) < CutoffPoor Then
            GCRichnessIndex(i) = "1"
        Else
            GCRichnessIndex(i) = "0"
        End If
    End If
Next i
  
GCRich = StringSubRegions(Join(GCRichnessIndex, ""), CStr(GCType), False)

End Function

'****************************************************************************************************
Sub ModifyTableData()

'====================================================================================================
'Truncates a data table in column or row dimension
'Thins columns and rows from the table using selected thinning factor
'Juraj Ahel, 2015-04-28
'Last update 2015-04-28
'====================================================================================================


Dim InputTable As Range, OutputTable As Range
Dim InputArray() As Variant, OutputArray() As Variant

Dim ParameterString As String
Dim StartRow As Long, EndRow As Long
Dim StartColumn As Long, EndColumn As Long
Dim RowThinFactor As Long, ColumnThinFactor As Long

Set InputTable = Application.InputBox("Select the table (or a part of it) to modify:", "Input selection", Type:=8)

ParameterString = InputBox("Start row; End row; Start column; End column; Keep every N-th row; Keep every N-th column." _
                        & vbCrLf & "Defaults give out an exact copy. Zero-end means full-length.", _
                        "Input parameters", _
                        "1;0;1;0;1;1")

Set OutputTable = Application.InputBox("Select the top left corner of the location where you want the processed table to be output.", _
                                        "Target output", Type:=8)


Parameters = Split(ParameterString, ";")

StartRow = Parameters(0)
EndRow = Parameters(1)
StartColumn = Parameters(2)
EndColumn = Parameters(3)
RowThinFactor = Parameters(4)
ColumnThinFactor = Parameters(5)

Dim InputTableWidth As Long, InputTableHeight As Long
Dim OutputTableWidth As Long, OutputTableHeight As Long

InputTableWidth = InputTable.Columns.Count
InputTableHeight = InputTable.Rows.Count

If EndColumn > InputTableWidth Or EndColumn = 0 Then EndColumn = InputTableWidth
If EndRow > InputTableHeight Or EndRow = 0 Then EndRow = InputTableHeight

OutputTableWidth = 1 + EndColumn - StartColumn
OutputTableWidth = 1 + Int((OutputTableWidth - 1) / ColumnThinFactor)
OutputTableHeight = 1 + EndRow - StartRow
OutputTableHeight = 1 + Int((OutputTableHeight - 1) / RowThinFactor)

'If OutputTableWidth * OutputTableHeight = 0 Then
'    MsgBox ("Zero-length output. Check input parameters and table selection.")
'    GoTo 999
'End If

ReDim InputArray(1 To InputTableHeight, 1 To InputTableWidth)
ReDim OutputArray(1 To OutputTableHeight, 1 To OutputTableWidth)

Dim TempFormula As String, TempColor As Long, TempPattern As Long

TempFormula = OutputTable.Value
TempColor = OutputTable.Interior.Color
TempPattern = OutputTable.Interior.Pattern

OutputTable.Formula = "!SELECTION!"
OutputTable.Interior.Color = vbRed
OutputTable.Interior.Pattern = xlSolid


ProceedCheck = MsgBox( _
                      "The macro will replace the area " & _
                      OutputTableHeight & " columns below and " & _
                      OutputTableWidth & " rows to the right of the top-left cell of the selection" _
                      & vbCrLf & "Any contents will be lost. Proceed?", _
                      vbYesNo, _
                      "Confirm replacement:" _
                     )

OutputTable.Formula = TempFormula
OutputTable.Interior.Color = TempColor
OutputTable.Interior.Pattern = TempPattern

If ProceedCheck = vbNo Then
    MsgBox ("Input aborted.")
    GoTo 999
End If

InputArray = InputTable.Value

Dim i As Long, j As Long, m As Long, N As Long

For i = 1 To OutputTableHeight
    m = StartRow + (i - 1) * RowThinFactor
    For j = 1 To OutputTableWidth
        N = StartColumn + (j - 1) * ColumnThinFactor
        OutputArray(i, j) = InputArray(m, N)
    Next j
Next i

Set OutputTable = OutputTable.Resize(OutputTableHeight, OutputTableWidth)

OutputTable.Value = OutputArray

999 End Sub

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
'Finds the (largest) continuous perfect overlap between two strings
'Juraj Ahel, 2015-04-30, for general purposes
'Last update 2015-04-30
'====================================================================================================

Dim ProbeLength As Long, TargetLength As Long
Dim Results As New Collection

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
'ReDim Results(1 To wStart)

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
                Results.Add j
            End If
        Loop Until Not FoundOverlap
        
    Next i
    
    W = W - 1
    
Loop Until k <> 0 Or W = 0

OverlapWidth = W + 1

Dim TempResultAsStrings() As String

Select Case k
    Case 0
        tempResult = "No overlap found."
    Case 1
        tempResult = Mid(Target, Results(1), OverlapWidth)
    Case Is > 1
        ReDim TempResultAsStrings(1 To k)
        i = 0
        For Each tempVar In Results
            i = i + 1
            TempResultAsStrings(i) = CStr(tempVar)
        Next tempVar
    
        
        tempResult = "Multiple equivalent results of length " _
                    & OverlapWidth & " at positions: " _
                    & Join(TempResultAsStrings, ";")
End Select

999 StringFindOverlap = tempResult

End Function


'****************************************************************************************************
Sub DeleteAllChartsOnSheet()
'====================================================================================================
'Deletes all charts on the active sheet
'Juraj Ahel, 2015-04-24
'Last update 2015-04-24
'====================================================================================================


Dim chtObj As ChartObject
For Each chtObj In ActiveSheet.ChartObjects
chtObj.Delete
Next

End Sub

'****************************************************************************************************
Sub FoldIndexDraw(WindowSize As Long, PlotRange As Range, _
                    LeftOffset, GraphWidth, _
                    TopOffset, GraphHeight, _
                    GraphMaximum, GraphMinimum, _
                    TickSpace, LabelSpace, DisplayGrid, _
                    Mode, Series As Integer)

'====================================================================================================
'Draws the graphs for FoldIndexMacro
'Juraj Ahel, 2015-04-24, for more automated FoldIndex-ing
'Last update 2015-11-09
'====================================================================================================

Dim myChart As Object
Dim srs As Series

'if I want it in the sheet
'If mode = 0 Then Set myChart = ActiveSheet.ChartObjects.Add(Left:=0, Width:=800, Top:=0, Height:=500).Chart

Select Case Mode

    Case 1
        SeriesNumber = 1
    Case 2
        SeriesNumber = 2
    Case 3
        SeriesNumber = Series
End Select

Set myChart = ActiveSheet.ChartObjects.Add(Left:=LeftOffset, Width:=GraphWidth, _
                                            Top:=TopOffset, Height:=GraphHeight) _
                                            .Chart



'How big the labels and markers on axes will be
TitleSize = 25
TickLabelSize = 25

Dim ChartColor()
Dim Data() As Range
ReDim Data(1 To SeriesNumber)
ReDim ChartColor(1 To SeriesNumber)

If Mode = 1 Then
    ChartColor(1) = 13998939 '-that bluish color
    Set Data(1) = PlotRange
Else
    'Green and Red, respectively for positive and negative series from FoldIndexMacro
    For i = 1 To SeriesNumber
        If i Mod 2 = 1 Then ChartColor(i) = RGB(25, 190, 25)
        If i Mod 2 = 0 Then ChartColor(i) = RGB(200, 25, 25)
        Set Data(i) = PlotRange.Offset(0, i - 1)
    Next i
End If



'Set Data(1) = PlotRange
'Set Data(2) = PlotRange.Offset(0, 1)

With myChart
    '.ChartTitle.text = "NiNTA"
    If Mode = 3 Then
        .HasTitle = False
    Else
        .HasTitle = True
        .ChartTitle.Text = CStr(WindowSize)
    End If
    '.Type = xlXYScatter
         
    'remove possible old series
    For Each srs In .SeriesCollection
        srs.Delete
    Next srs
    
    .ChartType = xlArea
        
    For i = 1 To SeriesNumber
        'introduce the series
        .SeriesCollection.NewSeries
        
           
        With .SeriesCollection(i)
            .Values = Data(i)
            .Format.Fill.ForeColor.RGB = ChartColor(i)
        End With
    Next i
    
    If DisplayGrid Then
        With .Axes(xlValue, 1)
            .HasTitle = True
            .MinimumScale = GraphMinimum
            .MaximumScale = GraphMaximum
            With .AxisTitle
                .Caption = "Fold Index"
                .Font.Size = TitleSize
            End With
            .MajorTickMark = xlTickMarkOutside
            .MinorTickMark = xlTickMarkOutside
            .Border.Weight = xlThick
            .Border.Color = RGB(0, 0, 0)
        End With
        Else
        With .Axes(xlValue, 1)
            .MinimumScale = GraphMinimum
            .MaximumScale = GraphMaximum
        End With
    End If
    
    If DisplayGrid Then
        With .Axes(xlPrimary)
            .HasTitle = True
            '.MinimumScale = 1
            '.MaximumScale = PlotRange.Rows.Count
            With .AxisTitle
                .Caption = "residue number"
                .Font.Size = TitleSize
            End With
            '.MinorUnit = .MajorUnit / 2
            .MajorTickMark = xlTickMarkOutside
            .MinorTickMark = xlTickMarkOutside
            .Border.Weight = xlThick
            .Border.Color = RGB(0, 0, 0)
            .TickMarkSpacing = TickSpace
            .TickLabelSpacing = LabelSpace
        End With
    Else
        With .Axes(xlPrimary)
            .MajorTickMark = xlNone
            .MinorTickMark = xlNone
            .Border.Weight = xlThin
            .Border.Color = RGB(0, 0, 0)
            .TickLabelPosition = xlTickLabelPositionNone
        End With
    End If
    
                    
    .Axes(xlCategory).TickLabels.Font.Size = TickLabelSize
    .Axes(xlValue, 1).TickLabels.Font.Size = TickLabelSize
    .Axes(xlValue).MajorGridlines.Delete
    '.Axes(xlValue).MinorGridlines.Delete
            
    '.Legend.Font.Size = 20
    .Legend.Delete
    .ChartArea.Border.LineStyle = xlNone
    .ChartArea.Format.Fill.Visible = msoFalse
    .PlotArea.Format.Fill.Visible = msoFalse
        
    'For Each srs In .SeriesCollection
    '    srs.Format.Line.Weight = 1
    'Next srs
        
    If Not DisplayGrid Then
        '.HasAxis(xlPrimary) = False
        .HasAxis(xlValue) = False
    End If
    
End With


End Sub

'****************************************************************************************************
Sub FoldIndexMacro()

'====================================================================================================
'Performs the FoldIndex calculation and generates the graphs to be imported in photoshop for overlaying
'Plots positive and negative values separately (different colors!)
'All the graphs have the same min / max x and y axes, so should be easy to overlay!
'The idea is to export the images using Daniel's XL Toolbox, and import them to Photoshop
'and overlaying them, with blend mode "Multiply" and then finetuning opacity to get optimal saturation
'UPDATE: idea since November 2015 is to copy directly to illustrator and do Overlay there
'So, export on one graph, and make sure you use RGB color mode for proper "Multiply+50 % opacity"
'method to work!
'Juraj Ahel, 2015-04-24
'Last update 2015-11-09
'====================================================================================================

Dim SeparatePositiveAndNegativeByColor As Boolean

SeparatePositiveAndNegativeByColor = True

Dim InputCell As Object, OutputRange As Range
Dim OutputTable() As Double
Dim WindowSizeList(), ScaleList()

Dim GraphNumber As Integer
Dim i As Long, SequenceLength As Long

Dim InputSequence As String, tempResult As String
Dim FoldIndexValues() As String

Dim MaxWindow As Long, MinWindow As Long, NumberOfWindows As Long
Dim SeppFactor As Double

Dim CtrlVar

MinWindow = 50
MaxWindow = 250
NumberOfWindows = 15

ReDim WindowSizeList(0 To NumberOfWindows - 1)

Set InputCell = Selection
'Set InputCell = Application.InputBox("Select cell containing input sequence:","Input selection",Type:=8)
                                    
InputSequence = CStr(InputCell.Value)
SequenceLength = Len(InputSequence)

If MaxWindow > SequenceLength \ 10 Then MaxWindow = SequenceLength \ 10

'Classic windows size list (first successful Mys1a overlay):
'WindowSizeList = Array(5, 25, 51, 75, 101, 151, 201)

'Generate equally log-spaced windows:
If NumberOfWindows > 1 Then
    SeppFactor = Log(MaxWindow / MinWindow) / (NumberOfWindows - 1)
End If

For i = 0 To NumberOfWindows - 1
    WindowSizeList(i) = Round(Exp(Log(MinWindow) + SeppFactor * i), 0)
Next i

'Essentially equal to NumberOfWindows
GraphNumber = UBound(WindowSizeList) - LBound(WindowSizeList) + 1

ReDim ScaleList(0 To GraphNumber - 1)

'Set scales proportional to window width
For i = 1 To GraphNumber
    ScaleList(GraphNumber - i) = WindowSizeList(GraphNumber - i) / WindowSizeList(GraphNumber - 1)
Next i

ReDim OutputTable(1 To SequenceLength + 1, 1 To 1 + 2 * GraphNumber)

'First column is used just to generate the last graph (the axes without the profile)
'For i = 1 To SequenceLength: OutputTable(i, 1) = i: Next i
For i = 1 To SequenceLength: OutputTable(i, 1) = 0: Next i

'Other columns are scaled FoldIndex profiles
For i = 1 To GraphNumber
    
    tempResult = FoldIndex(InputSequence, CLng(WindowSizeList(i - 1)), vbTab)
    FoldIndexValues = Split(tempResult, vbTab)
    
    OutputTable(1, 2 * i) = WindowSizeList(i - 1)
    OutputTable(1, 2 * i + 1) = WindowSizeList(i - 1)
    For j = 1 To SequenceLength
        TempNumber = CDbl(FoldIndexValues(j - 1)) * CDbl(ScaleList(i - 1))
        
        OutputTable(j + 1, 2 * i) = TempNumber
        OutputTable(j + 1, 2 * i + 1) = TempNumber
        'Positive and negative ones on separate series - data in separate columns!
        If SeparatePositiveAndNegativeByColor And TempNumber < 0 Then
            OutputTable(j + 1, 2 * i) = 0
        Else
            OutputTable(j + 1, 2 * i + 1) = 0
        End If
        
    Next j
    
    
Next i
    
'Output calculated data
Set OutputRange = InputCell.Offset(1, 0).Resize(SequenceLength, 1 + 2 * GraphNumber)
OutputRange.Value = OutputTable

Dim PlotRange As Range
Dim GraphMaximum As Double, GraphMinimum As Double
Dim DrawMode As Integer


'To get meaningfully scaled visuals, graphs are drawn between 110 % global minimum and 110 % global maximum
GraphMaximum = WorksheetFunction.Max(OutputRange.Offset(1, 1).Resize(SequenceLength, 2 * GraphNumber))
GraphMaximum = RoundToNearestX(1.1 * GraphMaximum, 0.01)
GraphMinimum = WorksheetFunction.Min(OutputRange.Offset(1, 1).Resize(SequenceLength, 2 * GraphNumber))
GraphMinimum = RoundToNearestX(1.1 * GraphMinimum, 0.01)

'The graphs are drawn one below the other using a separate drawing Sub
'graphs are drawn without axes for ease of overlaying, axes are available as a separate graph
'Unless option "together" is on, in which case everything is on one graph (for export to illustrator rather than photoshop)



For i = 1 To GraphNumber
    
    Set PlotRange = OutputRange.Offset(1, 2 * i - 1).Resize(SequenceLength, 1)
    
    'Where the graph is drawn and how big it is'
    LeftOffset = 0
    GraphWidth = 2000
    TopOffset = 0 + 275 * (i - 1)
    GraphHeight = 250
    
    'Spacing between major markers and labels. There is a minor marker between 2 major ones
    TickSpace = 500
    LabelSpace = 500
    
    CtrlVar = MsgBox("Separate graphs? (Yes = separate image each graph No = all graphs on same image)", _
                        vbYesNo, "Output format")
    If CtrlVar = vbYes Then
        If SeparatePositiveAndNegativeByColor Then DrawMode = 2 Else DrawMode = 1
    Else
        DrawMode = 3
    End If
    
    Call FoldIndexDraw(CLng(WindowSizeList(i - 1)), _
                        PlotRange, _
                        LeftOffset, GraphWidth, TopOffset, GraphHeight, _
                        GraphMaximum, GraphMinimum, _
                        TickSpace, LabelSpace, DisplayGrid:=False, Mode:=DrawMode, Series:=2 * GraphNumber)
    If DrawMode = 3 Then GoTo FinalGraph
Next i

'In the end, also draw a separate plot with the axes
FinalGraph:
Set PlotRange = OutputRange.Offset(1, 0).Resize(SequenceLength, 1)
    
    LeftOffset = 0
    GraphWidth = 2000
    TopOffset = 0 + 275 * GraphNumber
    GraphHeight = 250
    
    TickSpace = 500
    LabelSpace = 500
    
    Call FoldIndexDraw(0, _
                        PlotRange, _
                        LeftOffset, GraphWidth, TopOffset, GraphHeight, _
                        GraphMaximum, GraphMinimum, _
                        TickSpace, LabelSpace, DisplayGrid:=True, Mode:=1, Series:=1)


End Sub

****************************************************************************************
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
Function StringJoin(RangeToJoin As Range, Optional Separator As String = "", Optional Direction As Integer) As String

'====================================================================================================
'Joins all the cell values in an array as strings
'Juraj Ahel, 2015-02-16, for general purposes
'Last update 2015-04-13
'====================================================================================================

Dim tempString As String
Dim cell As Range

For Each cell In RangeToJoin
    tempString = tempString & cell.Value & Separator
Next cell

StringJoin = tempString

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

'****************************************************************************************************
Function SequenceRangeSelect( _
                             InputString As String, _
                             IndexRange As String, _
                             Optional DNA As Boolean = False, _
                             Optional Separator As String = "-" _
                            ) As String

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
Function StringCompare(a As String, b As String, Optional Limit As Integer = 25, Optional Mode As String = "Verbose") As String

'====================================================================================================
'Compares two strings and lists their differences, very raw so far
'Juraj Ahel, 2015-02-12, for comparing of protein sequences to find point mutations
'Last update 2015-02-12
'====================================================================================================

Dim i As Integer, j As Integer
Dim Result As String, S As String
Dim LA As Integer, LB As Integer
Dim counter As Integer: counter = 0
Dim cA As String, cB As String

LA = Len(a): LB = Len(b)

S = "; "

Select Case UCase(Mode)

Case "SHORT", "S"

Do
    i = i + 1
    cA = Mid(a, i, 1)
    cB = Mid(b, i, 1)
    
    If cA <> cB Then
        counter = counter + 1
        Result = Result & S & i
    End If
Loop Until i = LA Or i = LB Or ((counter > Limit) And (Limit > 0))


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
Loop Until i = LA Or i = LB Or ((counter > Limit) And (Limit > 0))

If counter = 0 And LA = LB Then
    Result = "Exact Copy!"
    GoTo 99
End If

End Select

If LA <> LB Then Result = Result & S & "LenDiff=" & LA - LB

If Len(Result) > 0 Then Result = Right(Result, Len(Result) - Len(S))

If counter > Limit And Limit > 0 Then Result = "Threshold (" & Limit & ") reached!"

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
Function DTT(x, Optional y = 0, Optional DateFormat As String = "YYMMDDhhmm", _
             Optional RoundingMode As Integer = -1, Optional Output As String = "d")

'====================================================================================================
'Converts YYMMDDhhmm date/time format to excel's date format
'if there is y, then calculates difference x-y instead of absolute date
'other date formats possibly to be added
'Juraj Ahel, 2014-06-08, for Master's thesis
'Last update 2014-12-31
'====================================================================================================


Dim yearx As Integer, monthx As Integer, dayx As Integer, hourx As Integer, minutex As Integer, secondx As Single
Dim timex As Date, timey As Date

Select Case DateFormat
    Case "YYMMDDhhmm"
                    
        yearx = 2000 + Mid(x, 1, 2)
        monthx = Mid(x, 3, 2)
        dayx = Mid(x, 5, 2)
        hourx = Mid(x, 7, 2)
        minutex = Mid(x, 9, 2)
        secondx = 0
        
    Case "YYMMDD"
                
        yearx = 2000 + Mid(x, 1, 2)
        monthx = Mid(x, 3, 2)
        dayx = Mid(x, 5, 2)
        hourx = 0
        minutex = 0
        secondx = 0
        
    Case "YYMMDDhhmmss"
                
        yearx = 2000 + Mid(x, 1, 2)
        monthx = Mid(x, 3, 2)
        dayx = Mid(x, 5, 2)
        hourx = Mid(x, 7, 2)
        minutex = Mid(x, 9, 2)
        secondx = Mid(x, 11, 2)
        
    Case "YYYYMMDDhhmm"
                    
        yearx = Mid(x, 1, 4)
        monthx = Mid(x, 5, 2)
        dayx = Mid(x, 7, 2)
        hourx = Mid(x, 9, 2)
        minutex = Mid(x, 11, 2)
        secondx = 0
        
    Case "YYYYMMDDhhmmss"
    
        yearx = Mid(x, 1, 4)
        monthx = Mid(x, 5, 2)
        dayx = Mid(x, 7, 2)
        hourx = Mid(x, 9, 2)
        minutex = Mid(x, 11, 2)
        secondx = Mid(x, 13, 2)
        
    Case Else
End Select

timex = DateSerial(yearx, monthx, dayx) + hourx / 24 + minutex / 1440 + secondx / 86400
timey = 0 'if there is no y it will stay 0

If Not (y = 0) Then
    
    Dim yeary As Integer, monthy As Integer, dayy As Integer, houry As Integer, minutey As Integer, secondy As Single
    
        Select Case DateFormat
        Case "YYMMDDhhmm"
                        
            yeary = 2000 + Mid(y, 1, 2)
            monthy = Mid(y, 3, 2)
            dayy = Mid(y, 5, 2)
            houry = Mid(y, 7, 2)
            minutey = Mid(y, 9, 2)
            secondy = 0
            
        Case "YYMMDD"
                    
            yeary = 2000 + Mid(y, 1, 2)
            monthy = Mid(y, 3, 2)
            dayy = Mid(y, 5, 2)
            houry = 0
            minutey = 0
            secondy = 0
            
        Case "YYMMDDhhmmss"
                    
            yeary = 2000 + Mid(y, 1, 2)
            monthy = Mid(y, 3, 2)
            dayy = Mid(y, 5, 2)
            houry = Mid(y, 7, 2)
            minutey = Mid(y, 9, 2)
            secondy = Mid(y, 11, 2)
            
        Case "YYYYMMDDhhmm"
                        
            yeary = Mid(y, 1, 4)
            monthy = Mid(y, 5, 2)
            dayy = Mid(y, 7, 2)
            houry = Mid(y, 9, 2)
            minutey = Mid(y, 11, 2)
            secondy = 0
            
        Case "YYYYMMDDhhmmss"
        
            yeary = Mid(y, 1, 4)
            monthy = Mid(y, 5, 2)
            dayy = Mid(y, 7, 2)
            houry = Mid(y, 9, 2)
            minutey = Mid(y, 11, 2)
            secondy = Mid(y, 13, 2)
            
        Case Else
    End Select

    timey = DateSerial(yeary, monthy, dayy) + houry / 24 + minutey / 1440 + secondy / 86400

End If

Dim Result As Single
Result = timex - timey

Select Case Output
Case "d"
Case "h"
    Result = 24 * Result
    RoundingMode = 1
Case "m"
    Result = 60 * 24 * Result
    RoundingMode = 1
End Select


    Select Case RoundingMode
        Case -2
        Case -1
        Case Else
        Result = Round(Result, RoundingMode)
    End Select
    
DTT = Result
    
End Function







'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------     Section       -----------------------------------------------
'----------------------------------------------------------------------------------------------------
'                     DNA                and                   PROTEINS
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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
    tempResult = Abs(CtrlSum) & "#!:"
    tempResult = tempResult & " for: " & Join(ErrMF, ", ")
    tempResult = tempResult & " rev: " & Join(ErrMR, ", ")
    GoTo 999
End If

OverhangF = Left(ForwardPrimer, Len(ForwardPrimer) - Len(OverlapF))
If Len(OverlapR) < Len(ReversePrimer) Then
    OverhangR = DNAReverseComplement(Left(ReversePrimer, Len(ReversePrimer) - Len(OverlapR)))
End If

If Not Details Then
    tempResult = PCRSimulate(Template, OverlapF, DNAReverseComplement(OverlapR), Circular, Perfect)
    tempResult = OverhangF & tempResult & OverhangR
Else
    tempResult = "F:" & OligoTm(OverlapF) & " °C, " & Len(OverlapF)
    tempResult = tempResult & " R:" & OligoTm(OverlapR) & " °C, " & Len(OverlapR)
End If

999 PCRWithOverhangs = tempResult

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
    ElseIf PrimerFCount = 0 Then
        Result = "No binding site found for Forward primer."
    ElseIf PrimerRCount = 0 Then
        Result = "No binding site found for Reverse primer."
    ElseIf PrimerFCount = 0 And PrimerRCount = 0 Then
        Result = "No binding site found for either primer!"
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
Dim R As Double, LogDNA As Double

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
    counter = counter + PairCount
Next i
    
R = 1.98717

'### Florian's version
'LogDNA = r * Ln(DNAc / 4)
    
'### Version from PrecisePrimer (different assumptions, focusing on the initial state where [primer]>>[template]
'### and also additional effect of terminal nucleotides (from SantaLucia et al.)
LogDNA = R * Ln(DNAc)
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
Function PrimerTm(PrimerSequence As String, Optional mismatch As Byte = 0) As Double
'====================================================================================================
'Calculates the melting temperature of a given primer, optionally also giving number of mismatches
'for mutagenesis
'Juraj Ahel, 2015-02-04, for checking primers
'Last update 2015-02-04
'====================================================================================================

Dim PrimerLength As Byte, GCNumber As Byte, i As Byte
Dim Tm As Double

PrimerLength = Len(PrimerSequence)
GCNumber = StringCharCount(PrimerSequence, "G", "C")

Tm = 81.5 + (41 * GCNumber - 675 - 100 * mismatch) / PrimerLength

PrimerTm = Round(Tm, 1)


End Function

'****************************************************************************************************
Function DNAReverseComplement(InputSequence As String) As String

'====================================================================================================
'Outputs a DNA reverse complement of a given input sequence
'Juraj Ahel, 2015-02-04, for checking primers
'Last update 2015-02-16
'====================================================================================================
'So far, always UPPERCASE output. Non-ACGT are preserved.

If InputSequence = "" Then
    DNAReverseComplement = ""
    GoTo 999
End If

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

999 End Function

'****************************************************************************************************
'****************************************************************************************************
Function DNATranslate(ByVal InputSequence As String) As String

'====================================================================================================
'Translates a DNA sequence to a protein sequence, using standard code
'sequence is given as a single-line string, output is also a single-line string
'other date formats possibly to be added
'Juraj Ahel, 2015-01-25, for general purposes
'Last update 2015-09-28
'====================================================================================================

Dim i As Integer, SequenceLength As Integer
Dim Aminoacid As String, OutputSequence() As String, Codon As String

SequenceLength = Len(InputSequence)

InputSequence = Replace(UCase(InputSequence), "U", "T")

ReDim OutputSequence(1 To SequenceLength \ 3)

For i = 1 To SequenceLength \ 3

    Codon = Mid(InputSequence, 3 * i - 2, 3)
    
    Select Case Codon
        Case "GCA", "GCC", "GCG", "GCT"
        Aminoacid = "A"
        Case "AGA", "AGG", "CGA", "CGC", "CGG", "CGT"
        Aminoacid = "R"
        Case "AAC", "AAT"
        Aminoacid = "N"
        Case "GAC", "GAT"
        Aminoacid = "D"
        Case "TGC", "TGT"
        Aminoacid = "C"
        Case "CAA", "CAG"
        Aminoacid = "Q"
        Case "GAA", "GAG"
        Aminoacid = "E"
        Case "GGA", "GGC", "GGG", "GGT"
        Aminoacid = "G"
        Case "CAC", "CAT"
        Aminoacid = "H"
        Case "ATA", "ATC", "ATT"
        Aminoacid = "I"
        Case "CTA", "CTC", "CTG", "CTT", "TTA", "TTG"
        Aminoacid = "L"
        Case "AAA", "AAG"
        Aminoacid = "K"
        Case "ATG"
        Aminoacid = "M"
        Case "TTC", "TTT"
        Aminoacid = "F"
        Case "CCA", "CCC", "CCG", "CCT"
        Aminoacid = "P"
        Case "AGC", "AGT", "TCA", "TCC", "TCG", "TCT"
        Aminoacid = "S"
        Case "ACA", "ACC", "ACG", "ACT"
        Aminoacid = "T"
        Case "TGG"
        Aminoacid = "W"
        Case "TAC", "TAT"
        Aminoacid = "Y"
        Case "GTA", "GTC", "GTG", "GTT"
        Aminoacid = "V"
        Case "TAA", "TAG", "TGA"
        Aminoacid = "*"
        Case Else
        Aminoacid = "X"
    End Select
    
    OutputSequence(i) = Aminoacid
    
Next i

DNATranslate = Join(OutputSequence, "")

End Function

'****************************************************************************************************
Private Function AACharge(pKa As Double, pH As Double, Species As Variant) As Double

'====================================================================================================
'Calculates the charge of a particular acidic or basic residue, needed for theoretical pI calculation
'Juraj Ahel, 2015-02-02, for theoretical pI calculation
'Last update 2015-02-02
'====================================================================================================
'Species 0 = acid, 1 = base

Dim ChargeSign As Integer
Dim Charge As Double

If Species = 0 Then ChargeSign = -1 Else ChargeSign = 1

If pH = pKa Then Charge = 0 Else Charge = ChargeSign / (1 + 10 ^ (ChargeSign * (pH - pKa)))

AACharge = Charge

End Function
'****************************************************************************************************
Function Theoretical_pI(ProteinSequence As String) As Double

'====================================================================================================
'Calculates theoretical pI from a given protein sequence
'Juraj Ahel, 2015-02-02, for the table of constructs so I don't ever have to go to ProtParam
'Last update 2015-02-02
'====================================================================================================
'Requires AACharge() and StringCharCount()
'Need to add optional selection of source of pKas

'aminoacid representations: 1=D; 2=E; 3=R; 4=K; 5=H; 6=C; 7=Y; 8=Cterm; 9=Nterm

AminoAcids = Array("D", "E", "R", "K", "H", "C", "Y")
AASpecies = Array(0, 0, 1, 1, 1, 0, 0, 0, 1) '0 is acid, 1 is base

Dim pKa(1 To 9) As Double
pKa(1) = 3.9: pKa(2) = 4.3: pKa(3) = 12.01: pKa(4) = 10.5: pKa(5) = 6.08: pKa(6) = 8.28: pKa(7) = 10.1: pKa(8) = 3.7: pKa(9) = 8.2

Dim AACounts(1 To 9) As Integer
Dim PartialCharges(1 To 9) As Double
Dim pH As Double, TotalCharge As Double
Dim i As Integer

For i = 1 To 7
    AACounts(i) = StringCharCount(ProteinSequence, AminoAcids(i - 1))
Next i

AACounts(8) = 1: AACounts(9) = 1

pH = 0
pHl = 0
pHh = 14

Do
    TotalCharge = 0
    
    For i = 1 To 9
        PartialCharges(i) = AACharge(pKa(i), pH, AASpecies(i - 1))
        TotalCharge = TotalCharge + AACounts(i) * PartialCharges(i)
    Next i
    
    If Abs(TotalCharge) < 0.01 Then GoTo 10
    
    If TotalCharge > 0 Then
        pHl = pH
        pH = (pH + pHh) / 2
    Else
        pHh = pH
        pH = (pH + pHl) / 2
    End If
    
Loop

10 Theoretical_pI = Round(pH, 1)
    
End Function

'****************************************************************************************************
Sub ChemicalFormulaFormat()

'====================================================================================================
'Resets subscripts, and then sets all numbers in selected cells to subscripts
'Juraj Ahel, 2015-02-05, for KnowledgeBase empirical formula formatting
'Last update 2015-02-05
'====================================================================================================

Dim List As Range, cell As Range
Dim Formula As String
Dim i As Integer

Set List = Application.Selection
Dim tmp As String

'List.Font.Subscript = True
'List.Font.Subscript = False

For Each cell In List
    cell.Font.Subscript = False
    For i = 1 To Len(cell)
        tmp = Mid(cell, i, 1)
        If tmp Like "#" Then
            cell.Characters(i, 1).Font.Subscript = True
        Else
            cell.Characters(i, 1).Font.Subscript = False
        End If
    Next i
Next cell


End Sub

'****************************************************************************************************
'****************************************************************************************************
'GIBSON ASSEMBLY
'****************************************************************************************************
'****************************************************************************************************

'****************************************************************************************************
Function ExtractParameter(Source As String, ParameterName As String, Optional MarkerType As String = "[]") As String

'====================================================================================================
'Finds marker-enclosed pieces of data. By default, the data are hugged by [Marker] and [\Marker],
'with option of picking different ways of doing it.
'Input is a string, and the function extracts the first such piece of data from a string.
'Juraj Ahel, 2015-02-11, for extracting values given by Florian Weissman's secondary structure script
'Last update 2015-02-11
'====================================================================================================

Dim S As String, e As String
Dim StartIndex As Integer, EndIndex As Integer

Dim Locs As Integer, Loce As Integer
Dim Data As String

Select Case UCase(MarkerType)
    Case "[]", "[", "]", "SQUARE"
        S = "[" & ParameterName & "]"
        e = "[\" & ParameterName & "]"
        Off = Len(S)
    Case Else
        Data = "Not yet supported, sorry. Use ""[]"" for MarkerType"
        GoTo 90
End Select

Locs = InStr(1, Source, S, vbTextCompare)
Loce = InStr(1, Source, e, vbTextCompare)
StartIndex = Locs + Len(S)
EndIndex = Loce - 1
Data = Mid(Source, StartIndex, EndIndex - StartIndex + 1)

90 ExtractParameter = Data

End Function




'****************************************************************************************************
Sub GibsonTest()

'====================================================================================================
'A huge procedure that generates the final result of Gibson overlap analysis by Florian's script
'It takes a range with prepared inputs, and directly outputs the results to 9 cells to the right
'This one should be made more modular / cleaned up, when I get the time
'Juraj Ahel, 2015-02-11, for Gibson assembly
'Last update 2015-02-11
'====================================================================================================

'Just for printing running time in the end
Dim StartTime As Single
StartTime = Timer

'Output variables
Dim OverlapSequence As String, PrimerSequence(1 To 2) As String, PrimerName(1 To 2) As String
Dim OverlapEnergy As Double, OverlapTm As Double, PrimerTmNN(1 To 2) As Double
Dim OutputData(1 To 1, 1 To 9) As Variant
Dim OutputRange As Range

'Input and internal variables
Dim myRange As Range, cell As Range
Dim FullPythonInputFilename As String, FullPythonOutputFilename As String
Dim RunDir As String
Dim textline As String
Dim TempName As String, tempExtension As String: tempExtension = ".jatmp"

'Where the program will run and do its internal stuff, like leaving temp files
RunDir = "c:\Excel_outputs\"

'Input
Set myRange = Selection
If Selection.Columns.Count > 1 Then
    MsgBox ("No. NO. NO. Just one Column! And to the right it must be free!")
    GoTo 999
End If



For Each cell In myRange
    If cell.Value <> "" And cell.Value <> 0 Then '##################################################################1
        
        'Temporary name to be used for temp files
        TempName = TempTimeStampName & "_Row" & cell.Row & "_Column" & cell.Column
                        
        'The python script needs an external file, at least I don't know how to pipe it in directly well
        Call ExportToTXT(cell, RunDir, TempName, tempExtension)
        
        'Temporary name for the python script input and output
        FullPythonInputFilename = RunDir & TempName & tempExtension
        FullPythonOutputFilename = RunDir & TempName & "_python" & tempExtension
        
        'This actually gives the result. The Python path is defined in the Subprocedure
        Call CallPythonScript(FullPythonInputFilename, RunDir, FullPythonOutputFilename)
                     
        'From here on, it's just reading the python output and putting it into excel
        Open FullPythonOutputFilename For Input As #1
        
        Do Until EOF(1)
            Line Input #1, textline
            
            If Left(textline, 1) = "[" Then
                Loc0t = InStr(1, textline, "]", vbTextCompare)
                DataType = Mid(textline, 2, Loc0t - 2)
                
                Select Case UCase(DataType)
                
                    Case "OVERLAP"
                        OverlapSequence = ExtractParameter(textline, "OverlapSequence", "[]")
                        OverlapEnergy = Val(ExtractParameter(textline, "dG", "[]"))
                        OverlapTm = Val(ExtractParameter(textline, "Tm", "[]"))
                    
                    Case "PRIMER1"
                        PrimerName(1) = ExtractParameter(textline, "PrimerName", "[]")
                        PrimerSequence(1) = ExtractParameter(textline, "Sequence", "[]")
                        PrimerTmNN(1) = Val(ExtractParameter(textline, "Tm", "[]"))
                    Case "PRIMER2"
                        PrimerName(2) = ExtractParameter(textline, "PrimerName", "[]")
                        PrimerSequence(2) = ExtractParameter(textline, "Sequence", "[]")
                        PrimerTmNN(2) = Val(ExtractParameter(textline, "Tm", "[]"))
                        
                End Select
                   
                
            End If
                
        Loop
        
        'Everything is outputed in the sheet, to the area just to the right of the input cells
        Set OutputRange = cell.Offset(0, 1).Resize(1, 9)
        
            OutputData(1, 1) = OverlapSequence
            OutputData(1, 2) = OverlapEnergy
            OutputData(1, 3) = OverlapTm
            OutputData(1, 4) = PrimerName(1)
            OutputData(1, 5) = PrimerSequence(1)
            OutputData(1, 6) = PrimerTmNN(1)
            OutputData(1, 7) = PrimerName(2)
            OutputData(1, 8) = PrimerSequence(2)
            OutputData(1, 9) = PrimerTmNN(2)
        
        OutputRange.Value = OutputData
        
        Close #1
        
        'If Deleting is true, temp files are deleted. I might add an inputbox to choose whether to do it
        Dim Deleting As Boolean, ExistenceTest As String
        Deleting = False
        
        If Deleting Then
            
            ExistenceTest = Dir(FullPythonInputFilename)
            If ExistenceTest <> "" Then Kill (FullPythonInputFilename)
            ExistenceTest = Dir(FullPythonOutputFilename)
            If ExistenceTest <> "" Then Kill (FullPythonOutputFilename)
            
        End If
                

    End If '#############################################################################################################1
Next cell

MsgBox ("Done! Runtime: " & Round((Timer - StartTime), 2) & " seconds")

999 'Goto

End Sub

Sub CallPythonScript(InputFile As String, RunDir As String, OutputFile As String)

Dim prog As String, path As String, argum As String
Dim wait As Boolean

prog = "python27.exe"
path = "C:\Python27\"
argum = """E:\PhD\Tools Alati\Gibson overlap by Florian Weissman\JA_overlap_local1.py"" " & """" & InputFile & """"

wait = True

Call CallProgram(prog, path, argum, wait, 0, RunDir, True, OutputFile)

End Sub

Sub CallPython1(ExeName As String, ExePath As String, argum As String, RunDir As String, Optional wait As Boolean = True, Optional OutputFile As String = "")


Call CallProgram(ExeName, ExePath, argum, wait, 1, RunDir, True, OutputFile)

End Sub
'****************************************************************************************************
Sub CallProgram( _
                ProgramCommand As String, _
                Optional ProgramPath As String = "", _
                Optional ArgList As String = "", _
                Optional WaitUntilFinished As Boolean = True, _
                Optional WindowMode As String = "1", _
                Optional RunDirectory As String = "", _
                Optional RunAsRawCmd As Boolean = True, _
                Optional OutputFile As String = "" _
               )

'====================================================================================================
'Calls an external program under the windows environment, using windows scripting host (Wsh)
'Takes more intuitive inputs and does all the complicated mimbo-jimbo so the code calling it is clean
'Juraj Ahel, 2015-02-11, for Gibson assembly and general purposes
'Last update 2015-02-11
'====================================================================================================
'Made for Excel Professional Plus 2013 under Windows 8.1

Dim wsh As Object
Dim WaitOnReturn As Boolean: WaitOnReturn = WaitUntilFinished
Dim WindowVisibilityType As Integer
Dim RunCommand As String, ProgramFullPath As String, ParsedArguments As String
Dim ProgramCommandTemp As String, ProgramPathTemp As String

ProgramCommandTemp = ProgramCommand
ProgramPathTemp = ProgramPath

'Parse program path if it's used, so it is formatted as a folder
If ProgramPathTemp <> "" Then
    Select Case Right(ProgramPathTemp, 1)
        Case "/", "\"
            ProgramPathTemp = Left(ProgramPathTemp, Len(ProgramPathTemp) - 1)
    End Select
    ProgramPathTemp = ProgramPathTemp & "\"
End If
            
ParsedArguments = ArgList

'Parse the run command so it actually works
RunCommand = ProgramCommandTemp
ProgramFullPath = ProgramPathTemp & RunCommand

RunCommand = ProgramFullPath & " " & ParsedArguments

'Parse the visibility options
Select Case UCase(WindowMode)
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
        WindowVisibilityType = CInt(WindowMode)
    Case "HIDDEN", "HIDE", "BACKGROUND"
        WindowVisibilityType = 0
    Case Else
        WindowVisibilityType = 1
End Select

'The object that does all the work
Set wsh = VBA.CreateObject("WSCript.Shell")

ParsedRunDirectory = RunDirectory
wsh.CurrentDirectory = ParsedRunDirectory

If RunAsRawCmd Then RunCommand = "%comspec% /c " & RunCommand

'2>&1 at the end ensures that the error log will be appended to the result! Cool!
If OutputFile <> "" Then RunCommand = RunCommand & " >""" & OutputFile & """ 2>&1"

a = wsh.Run(RunCommand, WindowVisibilityType, WaitOnReturn)

End Sub
Sub ExportToTXTMacro()

Dim SourceData As Range
Dim FilePath As String, FilenameBase As String

FilePath = "C:\Excel_outputs\"
FilenameBase = "Fragment "

Set SourceData = Selection


Call ExportToTXTSequence(SourceData, FilePath, FilenameBase, ".txt")

End Sub

'****************************************************************************************************
Sub ExportToTXT( _
                SourceData As Range, _
                Optional FilePath As String = "C:\Excel_outputs\", _
                Optional FilenameBase As String = "ExcelOutput ", _
                Optional Extension As String = ".txt" _
               )

'====================================================================================================
'Exports a separate text file for each cell in selection
'Still needs to be added modular file naming, now it's always "Fragment #.txt"
'
'Juraj Ahel, 2015-02-10, for Gibson assembly and general purposes
'Last update 2015-02-11
'====================================================================================================

Dim OutputFile As String
Dim DataSource As Range



    OutputFile = FilePath & FilenameBase & Extension
    Call ExportDataToTextFile(SourceData(1, 1).Value, OutputFile)


End Sub



'****************************************************************************************************
Sub ExportToTXTSequence(SourceData As Range, Optional FilePath As String = "C:\Excel_outputs\", Optional FilenameBase As String = "ExcelOutput ", Optional Extension As String = ".txt")

'====================================================================================================
'Exports a separate text file for each cell in selection
'Still needs to be added modular file naming, now it's always "Fragment #.txt"
'
'Juraj Ahel, 2015-02-10, for Gibson assembly and general purposes
'Last update 2015-02-11
'====================================================================================================

Dim OutputFile As String
Dim DataSource As Range

Dim i As Integer

Set DataSource = SourceData

For i = 1 To DataSource.Rows.Count
    OutputFile = FilePath & FilenameBase & i & Extension
    Call ExportDataToTextFile(DataSource(i, 1).Value, OutputFile)
Next i

End Sub



