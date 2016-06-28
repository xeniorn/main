Attribute VB_Name = "modGenerateCodingVariants"
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
Dim Counter As Long, CumulativeIndex As Long
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

    For Counter = 1 To NumberOfVariants
        'CodonIndex = 1 + Counter Mod Multiplicity(i)
        CodonIndex = 1 + ((Counter - 1) \ CumulativeIndex) Mod Multiplicity(i)
        Codon = Codons(AminoAcidIndex(i), CodonIndex)
        Variants(Counter, 1) = Variants(Counter, 1) & Codon
    Next Counter
    
    CumulativeIndex = Multiplicity(i) * CumulativeIndex

Next i

Selection.Offset(1, 0).Resize(NumberOfVariants, 1).Value = Variants

End Sub

