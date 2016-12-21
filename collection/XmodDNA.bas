Attribute VB_Name = "XmodDNA"
Option Explicit

'****************************************************************************************************
Function DNAEqual( _
    ByVal Input1 As String, _
    ByVal Input2 As String, _
    Optional ByVal Circular As Boolean = True, _
    Optional ByVal CheckReverseComplement As Boolean = True _
    ) As Boolean
   
'====================================================================================================
'Compares 2 DNA sequences to see if they are the same, allowing circular sequences
'Juraj Ahel, 2016-07-21
'
'====================================================================================================

    Dim i As Long
    Dim FoundForward As Boolean, FoundReverse As Boolean
    Dim tempIndex As Long
    
    If Len(Input1) <> Len(Input2) Then
        DNAEqual = False
        Exit Function
    End If
    
    Input1 = DNAParseTextInput(Input1, ExtendedSymbolSet:=True)
    Input2 = DNAParseTextInput(Input2, ExtendedSymbolSet:=True)
    
    If Circular Then
        tempIndex = InStr(1, Input2 & Input2, Input1)
        If tempIndex > 0 Then
            If Mid(Input2 & Input2, tempIndex, Len(Input1)) = Input1 Then
                FoundForward = True
            Else
                FoundForward = False
            End If
        Else
            FoundForward = False
        End If
    Else
        If Input1 = Input2 Then
            FoundForward = True
        Else
            FoundForward = False
        End If
    End If
        
    If CheckReverseComplement Then
        FoundReverse = DNAEqual(Input1, DNAReverseComplement(Input2), Circular, False)
    End If
    
    DNAEqual = FoundForward Or FoundReverse
    
End Function

'****************************************************************************************************
Function DNAParseTextInput( _
    ByVal InputString As String, _
    Optional ByVal Uppercase As Boolean = True, _
    Optional ByVal ExtendedSymbolSet As Boolean = False _
    ) As String
'====================================================================================================
'Finds the longest ORF in a DNA sequence, read in forward direction, assuming it's circular by default
'Juraj Ahel, 2016-06-28
'
'====================================================================================================
    
    Const AllowedSet As String = "[ATGC]"
    Const ExtendedAllowedSet As String = "[ACGTRYWSMKBDHVUN]"
    
    Dim i As Long
    Dim InputArray() As String, OutputArray() As String
    Dim InputLength As Long
    Dim tChar As String
    Dim Filter As String
    Dim Result As String
    
    If Len(InputString) = 0 Then
        DNAParseTextInput = ""
        Exit Function
    End If
    
    If ExtendedSymbolSet = True Then
        Filter = ExtendedAllowedSet
    Else
        Filter = AllowedSet
    End If
    
    InputLength = Len(InputString)
    
    ReDim InputArray(1 To InputLength)
    ReDim OutputArray(1 To InputLength)
    
    For i = 1 To InputLength
    
        tChar = UCase(Mid(InputString, i, 1))
        If tChar Like Filter Then OutputArray(i) = tChar
        
    Next i
    
    Result = Join(OutputArray, "")
    
    If Uppercase Then Result = UCase(Result)
    
    DNAParseTextInput = Result

End Function

'****************************************************************************************************
Private Function DNAFindORFsRecursion( _
    ByVal Sequence As String, _
    Optional ByVal MinimumORFLength As Long = 50 _
    ) As Collection
'====================================================================================================
'Finds the longest ORF in a DNA sequence, read in forward direction, assuming it's circular by default
'Juraj Ahel, 2016-06-28
'
'====================================================================================================
            
    Dim ORFs As VBA.Collection
    Dim tColl As VBA.Collection
    
    Dim tempORF As String
    
    Dim i As Long
    Dim j As Long
    
    Dim iStart As Long
    Dim iLen As Long
    
    Dim LeftSeq As String
    Dim RightSeq As String
    Dim SubSeq As String
    
    Dim Inputs As VBA.Collection
    
    
    tempORF = DNALongestORF(Sequence, False, 1, MinimumORFLength, False)
    
    If tempORF <> "" Then
        
        Set ORFs = New VBA.Collection
        ORFs.Add tempORF
        
        iStart = InStr(1, Sequence, tempORF)
        iLen = Len(tempORF)
        
        LeftSeq = Left(Sequence, iStart - 1)
        RightSeq = Right(Sequence, 1 + Len(Sequence) - iStart - iLen)
                
        'if the source is circular, then there is only 1 contigous DNA segment, not 2
        'otherwise, the fragments are recursively split around the
        
        Set Inputs = New VBA.Collection
            Inputs.Add LeftSeq
            Inputs.Add RightSeq
        
        For j = 1 To Inputs.Count
            SubSeq = Inputs.Item(j)
            If SubSeq <> "" Then
                Set tColl = DNAFindORFsRecursion(SubSeq, MinimumORFLength)
                If Not tColl Is Nothing Then
                    Call CollectionAppend(ORFs, tColl)
                End If
            End If
        Next j
                
    End If
    
    Dim tempCount As Long
    
    If ORFs Is Nothing Then
        tempCount = 0
    Else
        tempCount = ORFs.Count
    End If
    
    Debug.Print ("SeqLen: " & Len(Sequence) & " | ORFs: " & tempCount)
    
    Set DNAFindORFsRecursion = ORFs
    
    Set ORFs = Nothing
    Set tColl = Nothing

End Function

'****************************************************************************************************
Function DNAFindORFs( _
    ByVal Sequence As String, _
    Optional ByVal Circular As Boolean = True, _
    Optional ByVal MinimumORFLength As Long = 50, _
    Optional ByVal AllowORFOverlap As Boolean = False, _
    Optional ByVal AllowReverseStrand As Boolean = True _
    ) As Collection
'====================================================================================================
'Finds the longest ORF in a DNA sequence, read in forward direction, assuming it's circular by default
'Juraj Ahel, 2016-06-28
'2016-07-14 add quick hack to allow the case where the entire circular construct is an ORF, which does not start at 1
'====================================================================================================
'AllowORFOverlap not yet implemented - but it can be easily acquired by just running DNALongestORF

            
    Dim ORFs As VBA.Collection
    Dim tColl As VBA.Collection
    
    Dim tempORF As String
  
    Dim i As Long
    
    Dim iStart As Long
    Dim iLen As Long
    
    Dim LeftSeq As String
    Dim RightSeq As String
    
    If AllowReverseStrand Then
        'just call 2 recursions of the program without allowing strand reversal!
    
        Set ORFs = DNAFindORFs(Sequence, Circular, MinimumORFLength, AllowORFOverlap, False)
        Set tColl = DNAFindORFs(DNAReverseComplement(Sequence), Circular, MinimumORFLength, AllowORFOverlap, False)
        Call CollectionAppend(ORFs, tColl)
        
    Else
    
        If Circular Then
            'pick the first ORF and recursively treat the rest
        
            tempORF = DNALongestORF(Sequence, Circular, 1, MinimumORFLength, False)
            
            If tempORF <> "" Then
                
                Set ORFs = New VBA.Collection
                ORFs.Add tempORF
                
                iStart = InStr(1, Sequence, tempORF)
                iLen = Len(tempORF)
                
                'HACKY
                If iStart < 1 Then GoTo 999
                
                LeftSeq = Left(Sequence, iStart - 1)
                RightSeq = Right(Sequence, 1 + Len(Sequence) - iStart - iLen)
                        
                'if the source is circular, then there is only 1 contigous DNA segment, not 2
                'otherwise, the fragments are recursively split around the
                RightSeq = RightSeq & LeftSeq
                
                
            End If
            
        Else
            'go straight to recursion
            RightSeq = Sequence
        
        End If
        
        If RightSeq <> "" Then
            Set tColl = DNAFindORFsRecursion(RightSeq, MinimumORFLength)
            If Not tColl Is Nothing Then
                Call CollectionAppend(ORFs, tColl)
            End If
        End If
        
    End If
        
999 'hacky
        
    If Not ORFs Is Nothing Then
        If ORFs.Count > 1 Then
            Call SortStringCollectionByLength(ORFs)
        End If
    End If
        
    Set DNAFindORFs = ORFs
    
    Set ORFs = Nothing
    Set tColl = Nothing

End Function

'****************************************************************************************************
Function DNALongestORF( _
    ByVal Sequence As String, _
    Optional ByVal Circular As Boolean = True, _
    Optional ByVal GetNthORF As Long = 1, _
    Optional ByVal MinimumORFLength As Long = 50, _
    Optional ByVal AllowORFOverlap As Boolean = False _
    ) As String
'====================================================================================================
'Finds the longest ORF in a DNA sequence, read in forward direction, assuming it's circular by default
'Juraj Ahel, 2015-09-29
'Last update 2016-01-14
'2016-06-27 implement selecting Nth longest ORF
'2016-06-28 make overlapping ORFs optional
'2016-07-04 require a stop codon at the end of the CDS
'====================================================================================================
    
    Dim TempStart As Long, TempEnd As Long, BestStart As Long
    Dim SequenceLength As Long
    Dim MaxEnd As Long
    
    Dim ORFs As VBA.Collection
    Dim tColl As VBA.Collection
    Dim MakeCollection As Boolean
    
    Dim i As Long
    Dim j As Long
    Dim tIndex As Long
    Dim tCounter As Long
    
    Dim Codon As String
    Dim ScannedNucleotides As Long
    
    Dim BestLength As Long, CurrentLength As Long
    Dim BestSeq As String
    
    SequenceLength = Len(Sequence)
    
    'If Circular Then
        'Sequence = Right(Sequence, SequenceLength \ 2 + 1) & _
        '            Sequence & _
        '        Left(Sequence, SequenceLength \ 2 + 1)
    '    Sequence = Sequence & Left(Sequence, SequenceLength - 1)
    'End If
    
    TempStart = 0
    BestStart = 0
    
    'create a class to record all the longest ORFs
    If GetNthORF > 1 Then
        MakeCollection = True
        Set ORFs = New VBA.Collection
    Else
        MakeCollection = False
        Set ORFs = Nothing
    End If
    
    ':::main:::
    ScannedNucleotides = 0
    
    Do
    
        TempStart = InStr(TempStart + 1, Sequence, "ATG") 'beginning of start codon
        
        'if no more start codons found, exit the loop
        If TempStart = 0 Then
            Exit Do
        End If
        
        'for circular to work well, it is sufficient to make the start codon always at position 1
        If Circular Then
            Sequence = Right(Sequence, 1 + Len(Sequence) - TempStart) & Left(Sequence, TempStart - 1)
            ScannedNucleotides = ScannedNucleotides + TempStart - 1
            TempStart = 1
            'bit clumsy, but I must somehow avoid hitting the same start codon twice in circular mode
            'might redo this in another way in the future
            If ScannedNucleotides > SequenceLength Then
                Exit Do
            End If
        End If
        
        TempEnd = TempStart
        MaxEnd = TempStart + SequenceLength - 3 'the start of end codon must
        
        If MaxEnd > Len(Sequence) Then
            MaxEnd = Len(Sequence) - 2
        End If
        
        'just jump forward by 3 until a stop codon is reached (I'm sure this could be more efficient
        'but I don't care)
        j = 0
        Do
            TempEnd = TempEnd + 3
            Codon = Mid(Sequence, TempEnd, 3)
        Loop Until Codon = "TGA" Or Codon = "TAA" Or Codon = "TAG" Or TempEnd > MaxEnd
        
        Select Case Codon
            Case "TGA", "TAA", "TAG"
                'all good
            Case Else
                GoTo NextLoop
        End Select
        
        CurrentLength = TempEnd - TempStart
        
        'if we are in the tolerated range
        If CurrentLength > MinimumORFLength Then
                    
            'populate the collection
            If MakeCollection Then
                    
                Set tColl = New VBA.Collection
                    tColl.Add Mid(Sequence, TempStart, CurrentLength + 3)
                    tColl.Add CurrentLength
                    
                ORFs.Add tColl
                Set tColl = Nothing
                
            End If
            
            'update biggest result if necessary
            If CurrentLength > BestLength And CurrentLength <= SequenceLength Then
                
                BestLength = CurrentLength
                BestStart = TempStart
                BestSeq = Mid(Sequence, BestStart, BestLength + 3)
            
            End If
        
        End If
    
NextLoop:
    
    Loop Until TempStart = 0 Or TempStart > (SequenceLength - MinimumORFLength)
    
    
    'decide on output value
    If MakeCollection Then
        
        tCounter = 1
        
        'If AllowORFOverlap Then
        
        'remove N-1 largest ORFs from the collection
        Do While tCounter < GetNthORF
            
            CurrentLength = 0
            tIndex = 0
            For i = 1 To ORFs.Count
                If ORFs.Item(i).Item(2) >= CurrentLength Then
                    tIndex = i
                    CurrentLength = ORFs.Item(i).Item(2)
                End If
            Next i
            
            If tIndex = 0 Then
                Err.Raise 1, , "Tried to get ORF #" & GetNthORF & "but there are not enough ORFs above " _
                & "the threshold size (" & MinimumORFLength & ")!"
            End If
            
            ORFs.Remove (tIndex)
            tCounter = tCounter + 1
            
        Loop
        
        'get Nth ORF!
            CurrentLength = 0
            tIndex = 0
            For i = 1 To ORFs.Count
                If ORFs.Item(i).Item(2) > CurrentLength Then
                    tIndex = i
                    CurrentLength = ORFs.Item(i).Item(2)
                End If
            Next i
            
            If tIndex = 0 Then
                Err.Raise 1, , "Tried to get ORF #" & GetNthORF & "but there are not enough ORFs above " _
                & "the threshold size (" & MinimumORFLength & ")!"
            End If
            
            Set tColl = ORFs.Item(tIndex)
            BestStart = tColl.Item(1)
            BestLength = tColl.Item(2)
            Set tColl = Nothing
        
    End If
    
    
    DNALongestORF = BestSeq
    
    'cleanup
    If MakeCollection Then
        For i = 1 To ORFs.Count
            Call ORFs.Remove(1)
        Next i
    End If
    
    Set tColl = Nothing
    Set ORFs = Nothing

End Function


'****************************************************************************************************
Function OligoTm( _
                 Sequence As String, _
                 Optional EffectiveMonovalentCation_mM As Double = 50, _
                 Optional OligoConcentration_nM As Double = 500, _
                 Optional mode As String = "DNA", _
                 Optional TargetSequence As String = "" _
                ) As Double

'====================================================================================================
'Returns DNA melting temperature using Nearest Neighbour thermodynamics (NN)
'Works as EMBOSS dan, except it doesn't implement % formamide / DMSO and mismatches yet
'Replicated from Florian Weissman's script for Gibson assembly | originaly by Sebastina Bassi
'Juraj Ahel, 2015-02-11, for more proper oligo Tm calculations than with the older naive algorithm
'Last update 2015-03-24
'2016-06-28 explicit variable declaration
'====================================================================================================
'still lacks additional energy by terminal GC or AT on either side (can take also from PrecisePrimer manual)
'for this, I would first implement the possibility of selecting the subsequence that actually anneals, + mismatches
'Also, I would like to implement the effect of Magnesium (and other divalent) ions, and possibly DMSO

    Dim Pairs() As Variant, dHTable() As Variant, dSTable() As Variant
    Pairs = Array("AA", "TT", "AT", "TA", "CA", "TG", "GT", "AC", "CT", "AG", "GA", "TC", "CG", "GC", "GG", "CC")
    dHTable = Array(7.9, 7.9, 7.2, 7.2, 8.5, 8.5, 8.4, 8.4, 7.8, 7.8, 8.2, 8.2, 10.6, 9.8, 8, 8)
    dSTable = Array(22.2, 22.2, 20.4, 21.3, 22.7, 22.7, 22.4, 22.4, 21, 21, 22.2, 22.2, 27.2, 24.4, 19.9, 19.9)
    
    Sequence = UCase(Sequence)
    
    Dim i As Long
    Dim Seq() As String, Seqp() As String
    Dim N As Long
    Dim salt As Double, DNAc As Double
    Dim R As Double, LogDNA As Double
    
    Dim dG As Double
    
    Dim dH As Double, dS As Double
    Dim Pair As String, PairCount As Long
    
    Dim counter As Long
    
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
Function DNAReverseComplement(InputSequence As String) As String

'====================================================================================================
'Outputs a DNA reverse complement of a given input sequence
'Juraj Ahel, 2015-02-04, for checking primers
'Last update 2015-02-04
'2016-06-28 explicit variable declaration
'2016-07-14 exit check for zero length input
'====================================================================================================
'So far, always UPPERCASE output. Non-ACGT are preserved.
    
    Dim i As Long
    Dim j As Long
    
    Dim StringLength As Long
    Dim OutputSequence() As String
    
    If Len(InputSequence) = 0 Then
        DNAReverseComplement = ""
        Exit Function
    End If
    
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
    
    Dim i As Long, SequenceLength As Long, ProteinLength As Long
    Dim Aminoacid As String, OutputSequence As String, Codon As String
    Dim AminoAcids() As String
    
    SequenceLength = Len(InputSequence)
    OutputSequence = ""
    Aminoacid = ""
    i = 0
    
    InputSequence = Replace(UCase(InputSequence), "U", "T")
    
    ProteinLength = SequenceLength \ 3
    
    If SequenceLength Mod 3 <> 0 Then
        OutputSequence = "Input is not a valid coding sequence (len = 3k, k�N)"
        GoTo 99
    End If
    
    ReDim AminoAcids(1 To ProteinLength)
    
    For i = 1 To ProteinLength
    
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
        
        OutputSequence = OutputSequence & Aminoacid
        AminoAcids(i) = Aminoacid
        
    Next i
    
'99     DNATranslate = OutputSequence
99     DNATranslate = Join(AminoAcids, "")

End Function






'****************************************************************************************************
Function DNAGibsonLigation(ParamArray DNAList() As Variant) As String

'====================================================================================================
'Ligates a number of DNA sequences, requiring the final product to be circular
'Juraj Ahel, 2015-09-27
'Last update 2015-09-28
'====================================================================================================
'demonstrated to work 2015-09-28 on pJA1K and PLS46 (Mys1b in pFastBAC1 from 1-2, 3-5, 6-7, DF14)
'2016-07-04 add ability to pass an array variable as a parameter, in a dirty way
'           rewrite code a bit

    Const MinOverlap = 10          'overlap should be at least this
    Const MaxOverlapCheck = 250     'max meaningful to check, could be arbitrarily long code-wise, but no reason
    Const MinTm = 48                'Tm should be at least this
    
    Dim FragmentCount As Long
    Dim OverlapLength As Long
    Dim tempResult As String
    Dim i As Long, j As Long
    Dim NextFragment As Long
    Dim Tm As Double
    Dim ParsedDNAList() As Variant
            
    If VarType(DNAList(LBound(DNAList))) >= vbArray Then
        ParsedDNAList = DNAList(LBound(DNAList))
    Else
        ParsedDNAList = DNAList
    End If
    
    FragmentCount = 1 + UBound(ParsedDNAList) - LBound(ParsedDNAList)
    
    tempResult = ParsedDNAList(LBound(ParsedDNAList))
    
    
    For i = LBound(ParsedDNAList) To UBound(ParsedDNAList)
        
        If i = UBound(ParsedDNAList) Then
            NextFragment = LBound(ParsedDNAList)
        Else
            NextFragment = i + 1
        End If
        
        j = MaxOverlapCheck
        Do While (Right(ParsedDNAList(i), j) <> Left(ParsedDNAList(NextFragment), j))
            j = j - 1
        Loop
        
        OverlapLength = j
        Tm = OligoTm(Right(ParsedDNAList(i), j))
        
        If (OverlapLength < MinOverlap) Then
            tempResult = "#ERROR! Overlap " & (1 + i) & "-" & (1 + ((i + 1) Mod FragmentCount)) & " too low Tm!"
            GoTo 999
        End If
        If (Tm < MinTm) Then
            tempResult = "#ERROR! Overlap " & (1 + i) & "-" & (1 + ((i + 1) Mod FragmentCount)) & " too short!"
            GoTo 999
        End If
        
        ParsedDNAList(i) = Left(ParsedDNAList(i), Len(ParsedDNAList(i)) - OverlapLength)
        
    Next i
    
    tempResult = Join(ParsedDNAList, "")
    
999     DNAGibsonLigation = tempResult

End Function

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
'2016-06-27 put a condition Fsite > 1 - it was crashing when Primer would anneal at position 1!!!
'====================================================================================================

    Dim ErrorPrefix As String
    ErrorPrefix = "#! "
    
    Dim PrimerFCount As Long, PrimerRCount As Long
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
    
    Dim FSite As Long, RSite As Long, FLen As Long, RLen As Long
    Dim Reverse As Boolean
    
    Reverse = False
    FSite = InStr(1, Template, ForwardPrimer)
    RSite = InStr(1, Template, DNAReverseComplement(ReversePrimer))
    
    'If circular, pretend it's linear that starts exactly where F primer starts
    'and remap the indexing
    If Circular Then
        If FSite > 1 Then
            Template = SubSequenceSelect(Template, FSite, Len(Template)) & _
                        SubSequenceSelect(Template, 1, FSite - 1)
            RSite = RSite - FSite + 1
            FSite = 1
            If RSite < 1 Then RSite = Len(Template) + RSite
        End If
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
    
999     PCRSimulate = Result

End Function

'****************************************************************************************************
Function PCRWithOverhangs(Template As String, _
                    ForwardPrimer As String, ReversePrimer As String, _
                    Optional Circular = False, _
                    Optional Perfect = True, _
                    Optional IgnoreBestMatch = True, _
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
'2016-06-28 added explicit variable declaration + indentation
'2016-07-21 if the entire primer was annealing it would give horribly wrong result (N-1 !)
'====================================================================================================

    Dim OverhangF As String, OverhangR As String
    Dim OverlapF As String, OverlapR As String
    Dim ReversePrimerRC As String, TempFrag As String
    
    Dim i As Long
    
    Dim tempResult As String
    
    Dim NCheck As Long: NCheck = 3
    Dim ErrorMsg() As String
    Dim CtrlF() As Boolean, CtrlR() As Boolean
    Dim ErrMF() As String, ErrMR() As String
    ReDim ErrMF(1 To NCheck)
    ReDim ErrMR(1 To NCheck)
    ReDim ErrorMsg(1 To NCheck)
    ReDim CtrlF(1 To NCheck)
    ReDim CtrlR(1 To NCheck)
    Dim CtrlSum As Long
    
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
            
            If InStr(1, Template, TempFrag) = 0 Then
                i = i - 1
                Exit Do
            End If
            
            If i = Len(ForwardPrimer) Then
                Exit Do
            End If
        Loop
            
        OverlapF = Right(ForwardPrimer, i)
        
        i = 0
        Do
            i = i + 1
            TempFrag = Left(ReversePrimerRC, i)
            
            If InStr(1, Template, TempFrag) = 0 Then
                i = i - 1
                Exit Do
            End If
            
            If i = Len(ReversePrimer) Then
                Exit Do
            End If
        Loop
        
        OverlapR = Left(ReversePrimerRC, i)
        
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
        If Left(tempResult, 2) <> "#!" Then
            tempResult = OverhangF & tempResult & OverhangR
        End If
    Else
        tempResult = "F:" & OligoTm(OverlapF) & " �C, " & Len(OverlapF)
        tempResult = tempResult & " R:" & OligoTm(OverlapR) & " �C, " & Len(OverlapR)
    End If
    
999     PCRWithOverhangs = tempResult

End Function

'****************************************************************************************************
Function PCRGetFragmentFromTemplate( _
    ByVal TargetSequence As String, _
    ByVal Template As String, _
    Optional ByVal TargetPrimerTm As Double = 60, _
    Optional ByVal MinPrimerLength As Long = 15, _
    Optional ByVal TemplateCircular As Boolean = False, _
    Optional ByVal TryReverseComplement As Boolean = False, _
    Optional ByVal MaxExtension As Long = 50 _
    ) As VBA.Collection

'====================================================================================================
'Calculates optimal primers for getting a target sequence from a chosen template, allowing for extension
'Juraj Ahel, 2016-12-21, general purposes and Gibson

'====================================================================================================
'TODO: add check whether the primers anneal elsewhere in the sequence
'TODO: add check whether primers anneal to each other

    Dim OverlappingSequence As String
    Dim TempSeq As String
    Dim RCTemplate As String
    Dim TargetLength As Long, TemplateLength As Long, OverlapLength As Long
    Dim i As Long
    Dim RCIsBetter As Boolean
    
    Dim LeftExtensionLength As Long
    Dim RightExtensionLength As Long
        
    'find best overlap between template and target sequence - anything outside it is overlap
        If TemplateCircular Then
            For i = 1 To TemplateLength
                TempSeq = StringFindOverlap(TargetSequence, DNAReindex(Template, i))
                If Len(TempSeq) = Len(OverlappingSequence) Then
                    OverlappingSequence = TempSeq
                End If
            Next i
        Else
            OverlappingSequence = StringFindOverlap(TargetSequence, Template)
        End If
        
        RCIsBetter = False
        
    'check if a better overlap can be achieved with reverse complement
        If TryReverseComplement Then
            RCTemplate = DNAReverseComplement(Template)
            If TemplateCircular Then
                For i = 1 To TemplateLength
                    TempSeq = StringFindOverlap(TargetSequence, DNAReindex(RCTemplate, i))
                    If Len(TempSeq) > Len(OverlappingSequence) Then
                        OverlappingSequence = TempSeq
                        RCIsBetter = True
                    End If
                Next i
            Else
                TempSeq = StringFindOverlap(TargetSequence, RCTemplate)
                If Len(TempSeq) > Len(OverlappingSequence) Then
                    OverlappingSequence = TempSeq
                    RCIsBetter = True
                End If
            End If
        End If
        
    'If RC was better, simply pretend RC is the actual template during the calculation
    If RCIsBetter Then
        Template = RCTemplate
        Debug.Print ("Using Reverse Complement")
    End If
            
    TargetLength = Len(TargetSequence)
    TemplateLength = Len(Template)
    OverlapLength = Len(OverlappingSequence)
    
    'check if a PCR is feasible
        If TargetLength > OverlapLength + 2 * MaxExtension Then
            Call Err.Raise(1, , "Target sequence too long for template with given primer extension limit (" & MaxExtension & "). Total " & _
            (TargetLength - TemplateLength) & "-residue extension required.")
        End If
        
    LeftExtensionLength = InStr(1, TargetSequence, OverlappingSequence) - 1
    RightExtensionLength = TargetLength - OverlapLength - LeftExtensionLength
    
    If LeftExtensionLength > MaxExtension Then
        Call Err.Raise(1, , "Required 5' extension too long (" & LeftExtensionLength & ", allowed " & MaxExtension & ").")
    End If
    
    If RightExtensionLength > MaxExtension Then
        Call Err.Raise(1, , "Required 3' extension too long (" & RightExtensionLength & ", allowed " & MaxExtension & ").")
    End If
    
    Dim FAnneal As String, RAnneal As String
    Dim LeftExtension As String, RightExtension As String
    Dim FPrimer As String, RPrimer As String
        
    LeftExtension = Left(TargetSequence, LeftExtensionLength)
    RightExtension = Right(TargetSequence, RightExtensionLength)
        
    FAnneal = PCROptimizePrimer(OverlappingSequence, TargetPrimerTm, MinPrimerLength)
    RAnneal = PCROptimizePrimer(DNAReverseComplement(OverlappingSequence), TargetPrimerTm, MinPrimerLength)
    
    FPrimer = LeftExtension & FAnneal
    RPrimer = DNAReverseComplement(RightExtension) & RAnneal
    
    'confirm
        If PCRWithOverhangs(Template, FPrimer, RPrimer, TemplateCircular, True, True, False, MinPrimerLength) = TargetSequence Then
            Debug.Print ("Sequence ok!")
        Else
            Call Err.Raise(1, , "Simulated PCR result doesn't match target sequence!")
        End If
            
    Set PCRGetFragmentFromTemplate = New VBA.Collection
        With PCRGetFragmentFromTemplate
            .Add FPrimer, "F"
            .Add RPrimer, "R"
            .Add OligoTm(FAnneal), "TmF"
            .Add OligoTm(RAnneal), "TmR"
            .Add FAnneal, "AnnealF"
            .Add RAnneal, "AnnealR"
        End With
           
    Debug.Print (TargetSequence)
    Debug.Print (Template)
    For i = 1 To PCRGetFragmentFromTemplate.Count
        Debug.Print (PCRGetFragmentFromTemplate.Item(i))
    Next i
        

End Function

Sub testx()

Dim a As VBA.Collection

Set a = PCRGetFragmentFromTemplate("AAAAAAAAAAACATCATCATCATCATCATCATCATCATCATCCCCCCCCCCCCCCCCCCCCCCCCCCCCCGGGGGGGGGGG", "TTTTTTTTTTTTTTTTTTTTTTTTTTACATCATCATCATCATCATCATCATCATCATCCCCCCCCCCCTTTTTTTTTTTTTTTTTTTTTTTTTT")


End Sub

'****************************************************************************************************
Function PCROptimizePrimer(TargetSequence As String, Optional TargetTm As Double = 60, Optional MinLength As Long = 15) As String

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
Dim Length As Long
Dim Score() As Double, MaxScore As Long
Dim Variants() As String
Dim i As Long, j As Long
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

PCROptimizePrimer = Variants(j)

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
