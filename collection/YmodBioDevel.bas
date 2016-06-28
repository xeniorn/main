Attribute VB_Name = "YmodBioDevel"
Option Explicit

'****************************************************************************************************
Function mFastaToArrayOfFasta(FASTASequence As String, _
                                Optional SequenceCase As String = "UPPER", _
                                Optional Alignment As Boolean = False) As Variant

'====================================================================================================
'Takes a FASTA or mFASTA text as input and extracts the headers and the sequences to a 2D array
'Juraj Ahel, 2016-04-12, for general purposes
'Last update 2016-04-12
'====================================================================================================

    Dim SStart As Long, SEnd As Long, HStart As Long, HEnd As Long
    Dim SequenceNumber As Long
    Dim i As Long, j As Long
    
    Dim LineTerminator As String: LineTerminator = Chr(10)
    Dim DoubleLineTerminator As String: DoubleLineTerminator = LineTerminator & LineTerminator
    
    Dim ForbiddenSymbols As Variant
    
    Dim ErrorMessage As String
    
    Dim tempOutput() As String
    
    'Symbols that will be removed from sequence data - they are ok in headers'
    
    'Normal Fasta
    'ForbiddenSymbols = Array( _
    '                         Chr(9), Chr(124), LineTerminator, " ", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
    '                         "-", "*", ":", ";", "'", """", "#", "@", "&", "/", "\", "^", "_", "+", "?", "!", "$", _
    '                         "%", "=", "[", "]", "(", ")", "{", "}" _
    '                        )
    
    'AlignedFasta - do not remove "-"
    
    If Alignment Then
        ForbiddenSymbols = Array( _
                                 Chr(9), Chr(124), LineTerminator, " ", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                                 "*", ":", ";", "'", """", "#", "@", "&", "/", "\", "^", "_", "+", "?", "!", "$", _
                                 "%", "=", "[", "]", "(", ")", "{", "}" _
                                )
    Else
        ForbiddenSymbols = Array( _
                                 Chr(9), Chr(124), LineTerminator, " ", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                                 "*", ":", ";", "'", """", "#", "@", "&", "/", "\", "^", "_", "+", "?", "!", "$", _
                                 "%", "=", "[", "]", "(", ")", "{", "}", _
                                 "-")
    End If
    
                                
    SequenceNumber = StringCharCount(FASTASequence, ">")
    
    ReDim tempOutput(1 To SequenceNumber, 1 To 2)
    
    FASTASequence = Replace(FASTASequence, Chr(13), LineTerminator)
    FASTASequence = Replace(FASTASequence, Chr(10), LineTerminator)
    FASTASequence = FASTASequence & LineTerminator & ">" 'to allow termination for the final sequence
    
    Do While StringCharCount(FASTASequence, DoubleLineTerminator) > 0
        FASTASequence = Replace(FASTASequence, DoubleLineTerminator, LineTerminator)
    Loop
    
    HStart = 1: HEnd = 1
    SStart = 1: SEnd = 1
        
    For i = 1 To SequenceNumber
        
        HStart = InStr(SEnd, FASTASequence, ">", vbBinaryCompare) + 1
        HEnd = InStr(HStart, FASTASequence, LineTerminator, vbBinaryCompare) - 1
        SStart = HEnd + 2
        SEnd = InStr(SStart, FASTASequence, ">", vbBinaryCompare) - 2
        
        If HEnd > HStart Then tempOutput(i, 1) = Mid(FASTASequence, HStart, HEnd - HStart + 1) Else tempOutput(i, 1) = "[EMPTY_HEADER]"
        If SEnd > SStart Then tempOutput(i, 2) = Mid(FASTASequence, SStart, SEnd - SStart + 1) Else tempOutput(i, 2) = ""
        
        For j = LBound(ForbiddenSymbols) To UBound(ForbiddenSymbols)
            tempOutput(i, 2) = Replace(tempOutput(i, 2), ForbiddenSymbols(j), "")
        Next j
        
    Next i
        
    'Change case, as per settings. UPPERCASE is the default.
    
    Select Case SequenceCase
        Case "UPPER"
            For i = 1 To SequenceNumber: tempOutput(i, 2) = UCase(tempOutput(i, 2)): Next i
        Case "lower"
            For i = 1 To SequenceNumber: tempOutput(i, 2) = LCase(tempOutput(i, 2)): Next i
        Case "Preserve"
        Case Else
            For i = 1 To SequenceNumber: tempOutput(i, 2) = UCase(tempOutput(i, 2)): Next i
    End Select
    
    'in alignments, total length of each sequence must be the same
    If Alignment Then
        
        j = Len(tempOutput(1, 2))
        
        For i = 2 To SequenceNumber
            
            If Len(tempOutput(i, 2)) <> j Then
                ErrorMessage = "Sequence #" _
                                & i _
                                & " not equal in length to #1 (" _
                                & Len(tempOutput(i, 2)) _
                                & " vs " _
                                & j _
                                & "). Check input file!"
                Call Err.Raise(13, "mFastaToArrayOfFasta", ErrorMessage)
            End If
                        
        Next i
        
    End If
    
   
    mFastaToArrayOfFasta = tempOutput
    
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
Dim NumberOfVariants As Long, ProteinSequenceLength As Long
Dim Variants() As String
Dim AminoAcidIndex() As Long
Dim Multiplicity() As Long
Dim counter As Long, CumulativeIndex As Long
Dim Codon As String
Dim CodonIndex As Long

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
Function GCRich(InputSequence As String, GCType, HalfWindowSize As Long) As String

'====================================================================================================
'
'Juraj Ahel, 2015-07-09
'Last update 2015-07-09
'====================================================================================================

Dim SequenceLength As Long
Dim StartIndex As Long, EndIndex As Long
Dim GCRichness() As Double
Dim GCRichnessIndex() As String
Dim TempGCRich As Long

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


Function CalculateConservationScore(FASTAArray As Variant, _
                                    Optional ReferenceSequence As Long = 1, _
                                    Optional Smoothing As Long = 1) _
                                    As Variant
    
'this calculation is too basic. Try to find documentation about AACon from JalView, which uses
'physicochemical properties
    
    Const conSpacer = "-"
    
    Dim SequenceNumber As Long
    Dim SequenceLength As Long
    Dim ReferenceLength As Long
    Dim tempString As String
    
    Dim i As Long, j As Long
    
    Dim ConservationArray() As Long
    Dim ReferenceSeqArray() As String
    Dim OutputArray() As Long
    Dim tempSmooth
        
    SequenceNumber = 1 + UBound(FASTAArray, 1) - LBound(FASTAArray, 1)
    SequenceLength = Len(FASTAArray(1, 2))
    ReferenceLength = Len(Replace(FASTAArray(ReferenceSequence, 2), conSpacer, ""))
    
    ReDim ConservationArray(1 To SequenceLength)
    ReDim ReferenceSeqArray(1 To SequenceLength)
    
    For i = 1 To SequenceLength
        
        ReferenceSeqArray(i) = Mid(FASTAArray(ReferenceSequence, 2), i, 1)
        tempString = ""
        
        For j = 1 To SequenceNumber
            tempString = tempString & Mid(FASTAArray(j, 2), i, 1)
        Next j
        
        ConservationArray(i) = GetMaxLetterCount(tempString)
        
    Next i
    
    ReDim OutputArray(1 To ReferenceLength)
    
    If Smoothing > 1 Then
        
        tempSmooth = SmoothData(ConservationArray, Smoothing)
    
    Else
    
        tempSmooth = ConservationArray
    
    End If
        
    j = 0
    For i = 1 To SequenceLength
        If ReferenceSeqArray(i) <> conSpacer Then
            j = j + 1
            OutputArray(j) = tempSmooth(i)
        End If
    Next i
            
           
    CalculateConservationScore = OutputArray
    
    
End Function

'****************************************************************************************************
Function ReindexFromAlignment(inp1, Inp2)

    Const AllowedLetters = "[ACDEFGHIKLMNPQRSTVWY]"
    
    Dim len1, len2
    Dim i
    Dim a1(), a2(), a3()
    len1 = Len(inp1)
    
    ReDim a1(1 To len1)
    ReDim a2(1 To len1)
    
    c1 = 1
    c2 = 1
    
    For i = 1 To len1
    
        If Mid(inp1, i, 1) Like AllowedLetters Then
            a1(i) = c1
            c1 = c1 + 1
        Else
            a1(i) = ""
        End If
        
        If Mid(Inp2, i, 1) Like AllowedLetters Then
            a2(i) = c2
            c2 = c2 + 1
        Else
            a2(i) = ""
        End If
        
    Next i
    
    For i = 1 To len1
        If a1(i) Like "[1-9][0-9]*" Then
            OP = OP & vbCrLf & a1(i) & vbTab
            If a2(i) Like "##*" Then
                OP = OP & a2(i)
            Else
                OP = OP & "%"
            End If
        End If
    Next i
    
    aaa = Len(OP)
    ReindexFromAlignment = OP
    Call ExportStringToTXT(OP)


End Function

