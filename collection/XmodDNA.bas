Attribute VB_Name = "XmodDNA"
Option Explicit

Private Const DNAAllowedSet As String = "[ATGC]"
Private Const DNAExtendedAllowedSet As String = "[ACGTRYWSMKBDHVUN]"

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
'TODO: add proper support for upercase in input / output...
    
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
        Filter = DNAExtendedAllowedSet
    Else
        Filter = DNAAllowedSet
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
    
    'Debug.Print ("SeqLen: " & Len(Sequence) & " | ORFs: " & tempCount)
    
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
    ) As VBA.Collection
'====================================================================================================
'Finds the longest ORF in a DNA sequence, read in forward direction, assuming it's circular by default
'Juraj Ahel, 2016-06-28
'2016-07-14 add quick hack to allow the case where the entire circular construct is an ORF, which does not start at 1
'====================================================================================================
'AllowORFOverlap not yet implemented - but it can be easily acquired by just running DNALongestORF
'2017-01-03 implement so that it returns empty collection even if nothing is found - never return "nothing" unless it's an error state
            
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
    Else
        Set ORFs = New VBA.Collection
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
Function IsDNA(ByVal Sequence As String, Optional ByVal Extended As Boolean = False) As Boolean
'====================================================================================================
'checks if target is a valid DNA sequence
'Juraj Ahel, 2016-12-22
'====================================================================================================
    
    If Extended Then
        IsDNA = (UCase(Sequence) = DNAParseTextInput(Sequence, Uppercase:=True, ExtendedSymbolSet:=True))
    Else
        IsDNA = (UCase(Sequence) = DNAParseTextInput(Sequence, Uppercase:=True, ExtendedSymbolSet:=False))
    End If

End Function


'****************************************************************************************************
Function OligoTm( _
                 ByVal Sequence As String, _
                 Optional ByVal EffectiveMonovalentCation_mM As Double = 50, _
                 Optional ByVal OligoConcentration_nM As Double = 500, _
                 Optional ByVal mode As String = "DNA", _
                 Optional ByVal Template As String = "" _
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
'
'2016-12-22 make it byval
'           add checks for sequence and template
'           add support for calculating Tm vs an actual template (max overlap region)

    Dim Pairs() As Variant, dHTable() As Variant, dSTable() As Variant
    Pairs = Array("AA", "TT", "AT", "TA", "CA", "TG", "GT", "AC", "CT", "AG", "GA", "TC", "CG", "GC", "GG", "CC")
    dHTable = Array(7.9, 7.9, 7.2, 7.2, 8.5, 8.5, 8.4, 8.4, 7.8, 7.8, 8.2, 8.2, 10.6, 9.8, 8, 8)
    dSTable = Array(22.2, 22.2, 20.4, 21.3, 22.7, 22.7, 22.4, 22.4, 21, 21, 22.2, 22.2, 27.2, 24.4, 19.9, 19.9)
    
    If IsDNA(Sequence) Then Sequence = UCase$(Sequence)
    If IsDNA(Template) Then Template = UCase$(Template)
    
    If Len(Template) > 0 Then
        Debug.Print ("Calculating anneal to template...")
        Sequence = StringFindOverlap(Sequence, Template, False)
    End If
    
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
Function DNAReverseComplement(ByVal InputSequence As String) As String

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
Function DNATranslate( _
    ByVal InputSequence As String, _
    Optional IncludeStopCodon As Boolean = False _
    ) As String

'====================================================================================================
'Translates a DNA sequence to a protein sequence, using standard code
'sequence is given as a single-line string, output is also a single-line string
'other date formats possibly to be added
'Juraj Ahel, 2015-01-25, for general purposes
'Last update 2016-01-15
'====================================================================================================
'2017-01-03 make Stop Codon optional, default false
    
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
        OutputSequence = "Input is not a valid coding sequence (len = 3k, k€N)"
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
            If IncludeStopCodon Then
                Aminoacid = "*"
            Else
                Aminoacid = ""
            End If
            Case Else
            Aminoacid = "X"
        End Select
        
        'OutputSequence = OutputSequence & Aminoacid
        AminoAcids(i) = Aminoacid
        
    Next i
    
'99     DNATranslate = OutputSequence
99  DNATranslate = Join(AminoAcids, "")
    

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

Function PCRSimulate(ByVal Template As String, _
                    ByVal ForwardPrimer As String, ByVal ReversePrimer As String, _
                    Optional ByVal Circular = False, _
                    Optional ByVal Perfect = True _
                    ) As String

'====================================================================================================
'Simulates a PCR using selected primers. So far supports only perfect primers (no overhang, no mismatch)
'Can simulate PCR of circular templates
'Juraj Ahel, 2015-03-24, for Gibson assembly and general purposes
'Last update 2015-05-05
'2016-06-27 put a condition Fsite > 1 - it was crashing when Primer would anneal at position 1!!!
'====================================================================================================
'2016-12-22 make byval
'           make sure the correct result is given even when the sequence is so short the primers actually overlap!
'           add DNAReIndex instead of manual reindexing

    Dim ErrorPrefix As String
    ErrorPrefix = "#! "
    
    Dim PrimerFCount As Long, PrimerRCount As Long
    Dim Result As String
    
    If Len(Template) = 0 Then
        PCRSimulate = "#! empty template"
        Exit Function
    End If
    
    If Len(Template) < Len(ForwardPrimer) Or Len(Template) < Len(ReversePrimer) Then
        PCRSimulate = "#! primers too long for template"
        Exit Function
    End If
    
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
            Template = DNAReindex(Template, FSite)
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
        
    If (FSite + FLen > RSite + RLen) Or (RSite < FSite) Then
        Result = ErrorPrefix & "Primers extend over each other, check sequences."
    Else
        Result = SubSequenceSelect(Template, FSite, RSite + RLen - 1)
    End If
        
999 PCRSimulate = Result

End Function

'****************************************************************************************************
Function PCRWithOverhangs(ByVal Template As String, _
                    ByVal ForwardPrimer As String, ByVal ReversePrimer As String, _
                    Optional ByVal Circular = False, _
                    Optional ByVal Perfect = True, _
                    Optional ByVal IgnoreBestMatch = True, _
                    Optional ByVal Details = False, _
                    Optional ByVal MinimalOverlap = 15, _
                    Optional ByVal TryReverseComplement As Boolean = True _
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
'2016-12-22 add support for also PCRing from the reverse complement!
'2016-12-22 add byval

    Dim OverhangF As String, OverhangR As String
    Dim OverlapF As String, OverlapR As String
    Dim ReversePrimerRC As String, TempFrag As String
    
    Dim i As Long
    
    Dim tempResult As String
    Dim tempRC As String
    
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
        tempResult = "F:" & OligoTm(OverlapF) & " °C, " & Len(OverlapF)
        tempResult = tempResult & " R:" & OligoTm(OverlapR) & " °C, " & Len(OverlapR)
    End If
    
999 If TryReverseComplement Then
        tempRC = PCRWithOverhangs(DNAReverseComplement(Template), ForwardPrimer, ReversePrimer, Circular, Perfect, IgnoreBestMatch, Details, MinimalOverlap, False)
        If InStr(1, tempRC, "#!") = 0 Then
            If InStr(1, tempResult, "#!") = 0 Then
                tempResult = "#! Primers anneal both to forward and reverse strand!"
            Else
                tempResult = tempRC
            End If
        End If
    End If

PCRWithOverhangs = tempResult

End Function


'****************************************************************************************************
Function PCRGetFragmentFromTemplate( _
         ByVal TargetSequence As String, _
         ByVal Template As String, _
         Optional ByVal TargetPrimerTm As Double = 62, _
         Optional ByVal MinPrimerLength As Long = 15, _
         Optional ByVal TemplateCircular As Boolean = False, _
         Optional ByVal TryReverseComplement As Boolean = True, _
         Optional ByVal MaxExtension As Long = 50 _
         ) As VBA.Collection

'====================================================================================================
'Calculates optimal primers for getting a target sequence from a chosen template, allowing for extension
'Juraj Ahel, 2016-12-21, general purposes and Gibson

'====================================================================================================
'TODO: add check whether the primers anneal elsewhere in the sequence
'TODO: add check whether primers anneal to each other
'2016-12-21 add input parsing / error handling
'2016-12-22 change default primer Tm to 62
'           add check for Tm of non-anneal region

    Dim OverlappingSequence As String
    Dim TempSeq As String
    Dim RCTemplate As String
    Dim TargetLength As Long, TemplateLength As Long, OverlapLength As Long
    Dim i As Long
    Dim RCIsBetter As Boolean
    
    Dim LeftExtensionLength As Long
    Dim RightExtensionLength As Long
    
    Dim TmF As Double
    Dim TmR As Double
        
    If Len(TargetSequence) = 0 Then Call Err.Raise(1, , "Empty target sequence")
    If Len(Template) = 0 Then Call Err.Raise(1, , "Empty template")
        
    RCTemplate = DNAReverseComplement(Template)
        
    'find best overlap between template and target sequence - anything outside it is overlap
        If TemplateCircular Then
            For i = 1 To TemplateLength
                TempSeq = StringFindOverlap(TargetSequence, DNAReindex(Template, i), False)
                If Len(TempSeq) = Len(OverlappingSequence) Then
                    OverlappingSequence = TempSeq
                End If
            Next i
        Else
            OverlappingSequence = StringFindOverlap(TargetSequence, Template, False)
        End If
        
    RCIsBetter = False
        
    'check if a better overlap can be achieved with reverse complement
        If TryReverseComplement Then
            If TemplateCircular Then
                For i = 1 To TemplateLength
                    TempSeq = StringFindOverlap(TargetSequence, DNAReindex(RCTemplate, i), False)
                    If Len(TempSeq) > Len(OverlappingSequence) Then
                        OverlappingSequence = TempSeq
                        RCIsBetter = True
                    End If
                Next i
            Else
                TempSeq = StringFindOverlap(TargetSequence, RCTemplate, False)
                If Len(TempSeq) > Len(OverlappingSequence) Then
                    OverlappingSequence = TempSeq
                    RCIsBetter = True
                End If
            End If
        End If
        
    'If RC was better, simply pretend RC is the actual template during the calculation
        If RCIsBetter Then
            Call SwapValue(Template, RCTemplate)
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
    
    'check if extensions are well-calcualted / feasible
        Debug.Assert (LeftExtensionLength >= 0 And RightExtensionLength >= 0)
            
        If LeftExtensionLength > MaxExtension Then
            Call Err.Raise(1, , "Required 5' extension too long (" & LeftExtensionLength & ", allowed " & MaxExtension & ").")
        End If
            
        If RightExtensionLength > MaxExtension Then
            Call Err.Raise(1, , "Required 3' extension too long (" & RightExtensionLength & ", allowed " & MaxExtension & ").")
        End If
        
    'calculate annealing sequences and extension sequences
        Dim FAnneal As String, RAnneal As String
        Dim LeftExtension As String, RightExtension As String
        Dim FPrimer As String, RPrimer As String
                
        LeftExtension = Left(TargetSequence, LeftExtensionLength)
        RightExtension = Right(TargetSequence, RightExtensionLength)
                
        FAnneal = PCROptimizePrimer(OverlappingSequence, TargetPrimerTm, MinPrimerLength)
        RAnneal = PCROptimizePrimer(DNAReverseComplement(OverlappingSequence), TargetPrimerTm, MinPrimerLength)
    
    'Confirm annealing sites indeed correspond to termini
        Debug.Assert (Left(OverlappingSequence, Len(FAnneal)) = FAnneal)
        Debug.Assert (Right(OverlappingSequence, Len(RAnneal)) = DNAReverseComplement(RAnneal))
        
    
    'check if annealing sites are unique
        
        Dim tempc(1 To 4) As Long
            
        tempc(1) = StringCharCount_IncludeOverlap(Template, FAnneal)
        If TryReverseComplement Then
            tempc(2) = StringCharCount_IncludeOverlap(RCTemplate, FAnneal)
            tempc(3) = StringCharCount_IncludeOverlap(Template, RAnneal)
        End If
        tempc(4) = StringCharCount_IncludeOverlap(RCTemplate, RAnneal)
            
        If RCIsBetter Then
            Call SwapValue(tempc(1), tempc(2))
            Call SwapValue(tempc(3), tempc(4))
        End If
                
            
        Select Case True
            Case tempc(1) + tempc(2) > 1
                Call Err.Raise(1, , "forward anneal site (" & FAnneal & ") anneals at multiple sites - main strand: " & tempc(1) & " reverse strand: " & tempc(2) & "! Check sequence.")
            Case tempc(3) + tempc(4) > 1
                Call Err.Raise(1, , "reverse anneal site (" & RAnneal & ") anneals at multiple sites - main strand: " & tempc(3) & " reverse strand: " & tempc(4) & "! Check sequence.")
        End Select
            
            
        FPrimer = LeftExtension & FAnneal
        RPrimer = DNAReverseComplement(RightExtension) & RAnneal
        
        TmF = OligoTm(FAnneal)
        TmR = OligoTm(RAnneal)
        
    'confirm primers don't anneal better elsewhere
        If LeftExtensionLength > 0 Then
            If DNAAnnealToTemplate(Left(FPrimer, Len(FPrimer) - 2), Template) >= TmF - 2 Then
                Call Err.Raise(1, , "forward primer anneals to alternative site too well!")
            End If
            If DNAAnnealToTemplate(FPrimer, DNAReverseComplement(Template)) >= TmF - 5 Then
                Call Err.Raise(1, , "forward primer anneals to alternative site on RC strand too well!")
            End If
        End If
            
        If RightExtensionLength > 0 Then
            If DNAAnnealToTemplate(Left(RPrimer, Len(RPrimer) - 2), RCTemplate) >= TmR - 2 Then
                Call Err.Raise(1, , "reverse primer anneals to alternative site on RC strand too well!")
            End If
            If DNAAnnealToTemplate(RPrimer, Template) >= TmR - 5 Then
                Call Err.Raise(1, , "forward primer anneals to alternative site on primary too well!")
            End If
        End If
            
    
    
    'confirm in silico PCR
        If PCRWithOverhangs(Template, FPrimer, RPrimer, TemplateCircular, True, True, False, MinPrimerLength, True) = TargetSequence Then
            Debug.Print ("Simulated PCR successful!")
        Else
            Call Err.Raise(1, , "Simulated PCR result doesn't match target sequence!")
        End If
            
    'output
        Set PCRGetFragmentFromTemplate = New VBA.Collection
        With PCRGetFragmentFromTemplate
            .Add FPrimer, "F"
            .Add RPrimer, "R"
            .Add OligoTm(FAnneal), "TmF"
            .Add OligoTm(RAnneal), "TmR"
            .Add FAnneal, "AnnealF"
            .Add RAnneal, "AnnealR"
        End With
               
    'Debug output
    'Debug.Print (TargetSequence)
    'Debug.Print (Template)
    For i = 1 To PCRGetFragmentFromTemplate.Count
        Debug.Print (PCRGetFragmentFromTemplate.Item(i))
    Next i
        

End Function

'************************************************************************************
Private Function DNAFindProteinInTemplate( _
    ByVal ProteinSequence As String, _
    ByVal DNASource As String, _
    Optional ByVal Circular As Boolean = True, _
    Optional ByVal CheckReverseComplement As Boolean = True, _
    Optional ByVal Interactive As Boolean = True _
    ) As VBA.Collection

'====================================================================================================
'Takes in a protein sequence, DNA source and finds whether and where the protein is encoded in this DNA
'
'Juraj Ahel, 2017-01-03
'====================================================================================================

    Const MyName As String = "DNAFindProteinInTemplate"
        
'TODO: parse protein / DNA sequences / Trunc list
        
    Dim Translations As VBA.Collection
    Dim tempORF As String
    Dim i As Long
    
    Dim ProteinLength As Long
    Dim DNALength As Long
    
    Dim ORFLength As Long
    Dim ORFFound As Boolean
    Dim ORFLocus As Long
    
    Dim RCDNA As String
    Dim tempRCORFLocus As Long
    
    Dim IsReverse As Boolean
        
    ProteinLength = Len(ProteinSequence)
    DNALength = Len(DNASource)
    ORFLength = 3 * ProteinLength + 3
    
    If ProteinLength = 0 Or DNALength = 0 Then
        Call ApplyNewError(jaErr + 18, MyName, "Empty input")
        If Interactive Then
            ErrReraise
        Else
            Set DNAFindProteinInTemplate = Nothing
            Exit Function
        End If
    End If
        
    Set Translations = DNAFindORFs(DNASource, Circular, 24, True, CheckReverseComplement)
    
    For i = 1 To Translations.Count
        
        tempORF = Translations.Item(i)
        If Len(tempORF) < ORFLength Then
            Exit For
        Else
            If Len(tempORF) = ORFLength Then
                If DNATranslate(tempORF) = ProteinSequence Then
                    ORFFound = True
                    Exit For
                End If
            End If
        End If
        
    Next i
    
    If Not (ORFFound) Then
    
        Call ApplyNewError(jaErr + 1, MyName, "DNA does not encode for protein, check inputs")
        If Interactive Then
            ErrReraise
        Else
            Set DNAFindProteinInTemplate = Nothing
            Exit Function
        End If
        
    End If
    
    'if RC is to be checked, do it
    If CheckReverseComplement Then
        RCDNA = DNAReverseComplement(DNASource)
        tempRCORFLocus = DNAFindInsertInTemplate(tempORF, RCDNA, Circular, Interactive:=False)
        
        'handle errors
        If Err.Number <> 0 Then
            If Err.Source = "DNAFindInsertInTemplate" Then
                If Not Interactive Then
                    Set DNAFindProteinInTemplate = Nothing
                    Exit Function
                Else
                    ErrReraise
                End If
            Else
                ErrReraise
            End If
        End If
    End If
                
    ORFLocus = DNAFindInsertInTemplate(tempORF, DNASource, Circular, Interactive:=False)
    
    'handle errors
    If Err.Number <> 0 Then
        If Err.Source = "DNAFindInsertInTemplate" Then
            If Not Interactive Then
                Set DNAFindProteinInTemplate = Nothing
                Exit Function
            Else
                ErrReraise
            End If
        Else
            ErrReraise
        End If
    End If
    
    'confirm a locus could indeed be found
    If (ORFLocus <= 0) And Not (CheckReverseComplement And tempRCORFLocus > 0) Then
        Call ApplyNewError(jaErr + 2, MyName, "A strange error has occured - cannot confirm ORF was found even though it was found before")
        If Interactive Then
            ErrReraise
        Else
            Set DNAFindProteinInTemplate = Nothing
            Exit Function
        End If
    End If
    
    'check if the locus was unique
    If tempRCORFLocus > 0 Then
        If ORFLocus > 0 Then
            Call ApplyNewError(jaErr + 3, MyName, "Multiple encoding ORFs found for target protein in template DNA - cannot proceed")
            If Interactive Then
                ErrReraise
            Else
                Set DNAFindProteinInTemplate = Nothing
                Exit Function
            End If
        Else
            IsReverse = True
            DNASource = RCDNA
            ORFLocus = tempRCORFLocus
        End If
    Else
        If ORFLocus > 0 Then
            IsReverse = False
        End If
    End If
            
    Dim tempResult As VBA.Collection
    
    Set tempResult = New VBA.Collection
    
    tempResult.Add ORFLocus, "LOCUS"
    tempResult.Add IsReverse, "ISREVERSE"
    
    Set DNAFindProteinInTemplate = tempResult
    
    Set tempResult = Nothing
    
End Function

'****************************************************************************************************
Function DNAFindInsertInTemplate( _
         ByVal Probe As String, _
         ByVal Template As String, _
         Optional ByVal Circular As Boolean = True, _
         Optional Interactive As Boolean = True _
         ) As Long

'====================================================================================================
'finds where a certain DNA is inside another DNA, allowing also circular
'Juraj Ahel, 2017-01-03, for checking whether a plasmid encodes a protein

'====================================================================================================
    
    Dim ProbeLength As Long
    Dim TemplateLength As Long
    
    Dim tempLocus As Long
    Dim Overlap As String
        
    Dim OverlapCount As Long
        
    ProbeLength = Len(Probe)
    TemplateLength = Len(Template)
   
    'parse inputs
    If (ProbeLength > TemplateLength) Or (ProbeLength = 0) Then
    
        DNAFindInsertInTemplate = 0
        Exit Function
        
    End If
    
    If Circular Then Template = Template & Left(Template, TemplateLength - 1)

    Overlap = StringFindOverlap(Probe, Template, False)
    
    'if there is an error with finding an overlap
    If Err.Number <> 0 Then
    
        'if the error doesn't come from my function, but is a general error, raise it
        If Err.Source <> "StringFindOverlap" Then
            ErrReraise
        Else
            'otherwise, try to handle it
            Select Case Err.Number
            
                Case jaErr + 1
                    If Len(Overlap) <> ProbeLength Then
                        'this is not an error - insert simply doesn't exist in template
                        Err.Clear
                        DNAFindInsertInTemplate = 0
                        Exit Function
                    Else
                        Call ApplyNewError(jaErr + 1, "DNAFindInsertInTemplate", "Multiple overlaps of same length exist in template")
                    End If
                    
                Case Else
                    Call ApplyNewError(Err.Number, "DNAFindInsertInTemplate", "Unhandled error:" & Err.Source & ":" & Err.Number & ":" & Err.Description)
                    
            End Select
        End If
        
        If Interactive Then
            'throw the error
            Call ErrReraise
        Else
            'give null result and bubble error upwards if anyone cares (this would be handled as an _EVENT_ in a more modern language, e.g. VB.NET
            DNAFindInsertInTemplate = 0
            Exit Function
        End If
    
    End If
    
    If Len(Overlap) <> ProbeLength Then
    
        DNAFindInsertInTemplate = 0
        Exit Function
        
    Else

        tempLocus = InStr(1, Template, Probe)
    
        tempLocus = ((tempLocus - 1) Mod TemplateLength) + 1
    
        'this is logically approx:
        '
        'If tempLocus > TemplateLength Then
        '    tempLocus = tempLocus - TemplateLength
        'End If
    
    End If

        DNAFindInsertInTemplate = tempLocus

    End Function

'****************************************************************************************************
Function DNAAnnealToTemplate( _
    ByVal Probe As String, _
    ByVal Template As String, _
    Optional ByVal MaxLengthToTest As Long = 100, _
    Optional ByVal MinLengthToTest = 10 _
    ) As Double

'====================================================================================================
'Calculates the highest anneal temperature for a given oligo sequence binding to a given template
'Juraj Ahel, 2016-12-22, general purposes

'====================================================================================================

    Dim i As Long, j As Long
    Dim TestColl As VBA.Collection
    Dim tempString As String
    Dim tempTm As Double
    Dim bestTm As Double
    
    'Input parsing
        If MinLengthToTest < 1 Then MinLengthToTest = 1
        
        If Len(Probe) > MaxLengthToTest Then
            Call Err.Raise(1, , "Probe larger than largest allowed probe length (" & MaxLengthToTest & "). Increase the limit if you want to continue.")
        End If
        
        If Not (IsDNA(Probe) And IsDNA(Template)) Then
            Call Err.Raise(1, , "Inputs need to be proper DNA sequences")
        End If
        
    Set TestColl = New VBA.Collection
    
    'for all different substrings of Probe of target minimal length
        For i = 1 To Len(Probe) - MinLengthToTest + 1
            For j = i + MinLengthToTest - 1 To Len(Probe)
            
                tempString = SubSequenceSelect(Probe, i, j)
                
                'test if they have already been tested
                    If Not (IsElementOf(tempString, TestColl)) Then
                        TestColl.Add tempString, tempString
                        'if not, see if they exist in template
                            If InStr(1, Template, tempString) > 0 Then
                            
                                tempTm = OligoTm(tempString)
                                'and if their Tm is good, take it as the result
                                If tempTm > bestTm Then bestTm = tempTm
                                
                            End If
                    End If
                
            Next j
        Next i
    
    DNAAnnealToTemplate = bestTm
    
    'cleanup
        Set TestColl = Nothing

End Function




'****************************************************************************************************
Function PCROptimizePrimer(ByVal TargetSequence As String, Optional ByVal TargetTm As Double = 60, Optional ByVal MinLength As Long = 15) As String

'====================================================================================================
'Designs a simple primer for regular PCR amplification, trying to optimize the Tm and trying to
'keep the termini either G or C
'Always does a forward primer - do DNAReverseTranslate to Target to get the reverse. Might implement
'it as an option later
'In the future might be made more robust
'Juraj Ahel, 2015-03-24, general purposes
'Last update 2015-03-24
'====================================================================================================
'2016-12-22 make byval

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
Function DNAGCContent(ByVal Sequence As String) As Double
'====================================================================================================
'Calculates GC % as sum(G+C) / total length
'Juraj Ahel, 2015-09-28, for general purposes
'Last update 2015-09-28
'====================================================================================================
'2016-12-22 make byval

    DNAGCContent = StringCharCount(UCase(Sequence), "G", "C", "S") / Len(Sequence)

End Function

'****************************************************************************************************
Function DNAReindex(ByVal DNASequence As String, ByVal NewStartBase As Long) As String

'====================================================================================================
'Reindexes a circular DNA sequence
'Juraj Ahel, 2015-09-27
'Last update 2015-09-28
'====================================================================================================
'2016-12-22 make byval

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
Private Sub test_PCRSimulate()

    Const TestNumber As Long = 7
    Const FunctionName As String = "PCRSimulate"
    
    Dim N As Long
    Dim TestResults(1 To TestNumber) As Long
    Dim Test As String
    Dim Input1 As String, Input2 As String, Input3 As String
    
    On Error Resume Next
    
    '1 empty inputs
        N = 1
        Test = PCRSimulate("AAAAA", vbNullString, vbNullString)
        If Err.Number = 0 Then
            If Left(Test, 3) = "#! " Then TestResults(N) = 1
        End If
        Err.Clear
        
    '2 empty inputs
        N = 2
        Test = PCRSimulate("AAAAAAAAAAAAAAAAAAAAA", "AAAAA", vbNullString)
        If Err.Number = 0 Then
            If Left(Test, 3) = "#! " Then TestResults(N) = 1
        End If
        Err.Clear
        
    '3 empty inputs
        N = 3
        Test = PCRSimulate("AAAAAAAAAAAAAAAAAAAAA", vbNullString, "AAAAA")
        If Err.Number = 0 Then
            If Left(Test, 3) = "#! " Then TestResults(N) = 1
        End If
        Err.Clear
        
    '4 multiple anneal sites (repetitive sequence)
        N = 4
        Input3 = "TGAGGGGGAAAGTGGTGAGTG"
        Input1 = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" & DNAReverseComplement(Input3)
        Input2 = "AAAAAAAAAAAAAAAAAAAAAA"
        Test = PCRSimulate(Input1, Input2, Input3)
        If Err.Number = 0 Then
            If Left(Test, 3) = "#! " Then TestResults(N) = 1
        End If
        Err.Clear
        
    '5 positive control no overhangs
        N = 5
        Input2 = "AAAAAAAAAAAAAAAA"
        Input3 = "GGGGGGGGGGGGGGGG"
        Input1 = Input2 & "TTTTTTTTTTTTTTTTTTTTT" & DNAReverseComplement(Input3)
        
        Test = PCRSimulate(Input1, Input2, Input3)
        If Err.Number = 0 Then
            If Left(Test, 3) = "#! " Then TestResults(N) = 1
        End If
        Err.Clear
        
    '6 positive control with overhangs
        N = 6
        Input1 = "TTTTT" & "AAAAAGGGGGTTTTTCCCCC" & "TTTTTTTTTTTTTTTTT" & "AGTCAGTCAGTCAGTCAGTC" & "TTTTT"
        Input2 = "AAAAAGGGGGTTTTTCCCCC"
        Input3 = DNAReverseComplement("AGTCAGTCAGTCAGTCAGTC")
        Test = PCRSimulate(Input1, Input2, Input3)
        If Err.Number = 0 Then
            If Test = "AAAAAGGGGGTTTTTCCCCC" & "TTTTTTTTTTTTTTTTT" & "AGTCAGTCAGTCAGTCAGTC" Then TestResults(N) = 1
        End If
        Err.Clear
        
    '7 primers overlap on sequence
        N = 7
        Input1 = "AAAAAGGGGGTTTTTCCCCC" & "AGTCAGTCAGTCAGTCAGTC"
        Input2 = "AAAAAGGGGGTTTTTCCCCCAGT"
        Input3 = DNAReverseComplement("CCC" & "AGTCAGTCAGTCAGTCAGTC")
        Test = PCRSimulate(Input1, Input2, Input3)
        If Err.Number = 0 Then
            If Test = Input1 Then TestResults(N) = 1
        End If
        Err.Clear
    
    
    
    On Error GoTo 0
    
    Dim i As Long
    Dim j As Long
    
    For i = 1 To TestNumber
        If TestResults(i) <> 1 Then
            Debug.Print ("Test #" & i & " failed!")
            j = j + 1
        End If
    Next i
    
    If j = 0 Then
        Debug.Print (FunctionName & ": All tests successfully passed!")
        If JA_InteractiveTesting Then MsgBox (FunctionName & ": All tests successfully passed!")
    Else
        Debug.Print (FunctionName & ": " & j & " tests failed!")
        If JA_InteractiveTesting Then MsgBox (FunctionName & ": " & j & " tests failed!")
    End If

End Sub


'****************************************************************************************************
Private Sub test_PCRGetFragmentFromTemplate()

    Const TestNumber As Long = 9
    Const FunctionName As String = "PCRGetFragmentFromTemplate"
    
    Dim TestResults(1 To TestNumber) As Long
    Dim Test As VBA.Collection
    Dim Input1 As String, Input2 As String
    
    On Error Resume Next
    
    '1 empty inputs
        Set Test = PCRGetFragmentFromTemplate(vbNullString, vbNullString)
        If Err.Number = 1 Then TestResults(1) = 1
        Err.Clear
        
    '2 empty inputs
        Set Test = PCRGetFragmentFromTemplate(vbNullString, "A")
        If Err.Number = 1 Then TestResults(2) = 1
        Err.Clear
        
    '3 empty inputs
        Set Test = PCRGetFragmentFromTemplate("A", vbNullString)
        If Err.Number = 1 Then TestResults(3) = 1
        Err.Clear
        
    '4 multiple anneal sites (repetitive sequence)
        Input1 = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
        Set Test = PCRGetFragmentFromTemplate(Input1, Input1)
        If Err.Number = 1 Then TestResults(4) = 1
        Err.Clear
        
    '5 positive control no overhangs
        Input1 = "AAAAAGGGGGTTTTTCCCCC" & "AGTCAGTCAGTCAGTCAGTC"
        Set Test = PCRGetFragmentFromTemplate(Input1, Input1)
        If Err.Number = 0 Then
            If PCRWithOverhangs(Input1, Test.Item(1), Test.Item(2)) = Input1 Then
                TestResults(5) = 1
            End If
        End If
        Err.Clear
        
    '6 positive control with overhangs
        Input1 = "TTTTT" & "AAAAAGGGGGTTTTTCCCCC" & "AGTCAGTCAGTCAGTCAGTC" & "TTTTT"
        Input2 = "AAAAAGGGGGTTTTTCCCCC" & "AGTCAGTCAGTCAGTCAGTC"
        Set Test = PCRGetFragmentFromTemplate(Input1, Input2)
        If Err.Number = 0 Then
            If PCRWithOverhangs(Input2, Test.Item(1), Test.Item(2)) = Input1 Then
                TestResults(6) = 1
            End If
        End If
        Err.Clear
        
    '7 template has repeat sequence
        Input1 = "AAAAAGGGGGTTTTTCCCCC" & "AGTCAGTCAGTCAGTCAGTC"
        Input2 = "AAAAAGGGGGTTTTTCCCCC" & "AGTCAGTCAGTCAGTCAGTC" & "AGTCAGTCAGTCAGTCAGTC"
        Set Test = PCRGetFragmentFromTemplate(Input1, Input2)
        If Err.Number = 1 Then TestResults(7) = 1
        Err.Clear
        
    '8 positive control with overhangs and extra stuff in template
        Input1 = "TTTTT" & "AAAAAGGGGGTTTTTCCCCC" & "AGTCAGTCAGTCAGTCAGTC" & "TTTTT"
        Input2 = "TAGGGATTAGGGATTAGGGAT" & "AAAAAGGGGGTTTTTCCCCC" & "AGTCAGTCAGTCAGTCAGTC" & "CCTTCCTTCTCTCCTTCCTCTCT"
        Set Test = PCRGetFragmentFromTemplate(Input1, Input2)
        If Err.Number = 0 Then
            If PCRWithOverhangs(Input2, Test.Item(1), Test.Item(2)) = Input1 Then
                TestResults(8) = 1
            End If
        End If
        Err.Clear
        
    'test 9 - positive control with overhangs and extra stuff in template
        Input1 = "TTTTT" & "AAAAAGGGGGTTTTTCCCCC" & "AGTCAGTCAGTCAGTCAGTC" & "TTTTT"
        Input2 = DNAReverseComplement("TAGGGATTAGGGATTAGGGAT" & "AAAAAGGGGGTTTTTCCCCC" & "AGTCAGTCAGTCAGTCAGTC" & "CCTTCCTTCTCTCCTTCCTCTCT")
        Set Test = PCRGetFragmentFromTemplate(Input1, Input2)
        If Err.Number = 0 Then
            If PCRWithOverhangs(Input2, Test.Item(1), Test.Item(2)) = Input1 Then
                TestResults(9) = 1
            End If
        End If
        Err.Clear
    
    'test 10 - primer extension annealing too well - better than normal anneal
        
    
    
    
    
    On Error GoTo 0
    
    Dim i As Long
    Dim j As Long
    
    For i = 1 To TestNumber
        If TestResults(i) <> 1 Then
            Debug.Print ("Test #" & i & " failed!")
            j = j + 1
        End If
    Next i
    
    If j = 0 Then
        Debug.Print (FunctionName & ": All tests successfully passed!")
        If JA_InteractiveTesting Then MsgBox (FunctionName & ": All tests successfully passed!")
    Else
        Debug.Print (FunctionName & ": " & j & " tests failed!")
        If JA_InteractiveTesting Then MsgBox (FunctionName & ": " & j & " tests failed!")
    End If

End Sub



'****************************************************************************************************
Private Sub test_DNAFindInsertInTemplate()

    Const TestNumber As Long = 11
    Const FunctionName As String = "DNAFindInsertInTemplate"
    
    Dim N As Long
    Dim TestResults(1 To TestNumber) As Long
    Dim Test As String
    Dim Input1 As String, Input2 As String, Input3 As String
    
    On Error Resume Next
    
    '1 empty inputs
        N = 1
        Test = DNAFindInsertInTemplate("", "", True)
        If Err.Number = 0 Then
            If Test = 0 Then TestResults(N) = 1
        End If
        Err.Clear
        
    '2 empty inputs
        N = 2
        Test = DNAFindInsertInTemplate("", "", False)
        If Err.Number = 0 Then
            If Test = 0 Then TestResults(N) = 1
        End If
        Err.Clear
        
    '3 simple input
        N = 3
        Test = DNAFindInsertInTemplate("A", "A", True)
        If Err.Number = 0 Then
            If Test = 1 Then TestResults(N) = 1
        End If
        Err.Clear
        
    '4 repetitive sequence - interactive
        N = 4
        Input1 = "A"
        Input2 = "AA"
        Test = DNAFindInsertInTemplate(Input1, Input2, False, True)
        If Err.Number = jaErr + 1 Then
            If Test = 1 Then TestResults(N) = 1
        End If
        Err.Clear
    
    'repetitive sequence - batch
        N = 5
        Input1 = "A"
        Input2 = "AA"
        Test = DNAFindInsertInTemplate(Input1, Input2, False, False)
        If Err.Number = jaErr + 1 Then
            If Test = 1 Then TestResults(N) = 1
        End If
        Err.Clear
        
    'positive control - simple
        N = 5
        Input1 = "ATGGTA"
        Input2 = "ATGGTA"
        Test = DNAFindInsertInTemplate(Input1, Input2, False, False)
        If Err.Number = 0 Then
            If Test = 1 Then TestResults(N) = 1
        End If
        Err.Clear
        
    'positive control - added
        N = 6
        Input1 = "ATGGTA"
        Input2 = "ATGGTACCCCC"
        Test = DNAFindInsertInTemplate(Input1, Input2, False, False)
        If Err.Number = 0 Then
            If Test = 1 Then TestResults(N) = 1
        End If
        Err.Clear
        
    'positive control - added
        N = 7
        Input1 = "ATGGTA"
        Input2 = "CCCCCATGGTA"
        Test = DNAFindInsertInTemplate(Input1, Input2, False, False)
        If Err.Number = 0 Then
            If Test = 6 Then TestResults(N) = 1
        End If
        Err.Clear
        
    'positive control - circular, trivial
        N = 8
        Input1 = "ATG"
        Input2 = "ATG"
        Test = DNAFindInsertInTemplate(Input1, Input2, True, False)
        If Err.Number = 0 Then
            If Test = 1 Then TestResults(N) = 1
        End If
        Err.Clear
        
    'positive control - circular, simple
        N = 9
        Input1 = "ATG"
        Input2 = "GAT"
        Test = DNAFindInsertInTemplate(Input1, Input2, True, False)
        If Err.Number = 0 Then
            If Test = 2 Then TestResults(N) = 1
        End If
        Err.Clear
        
    'positive control - circular, full
        N = 10
        Input1 = "ATGCGA" & "GGTAT"
        Input2 = "GGTAT" & "CCCCCCCCCCCCCCCC" & "ATGCGA"
        Test = DNAFindInsertInTemplate(Input1, Input2, True, False)
        If Err.Number = 0 Then
            If Test = 22 Then TestResults(N) = 1
        End If
        Err.Clear
        
    'negative control - circular, simple
        N = 11
        Input1 = "AG"
        Input2 = "GA"
        Test = DNAFindInsertInTemplate(Input1, Input2, False, False)
        If Err.Number = 0 Then
            If Test = 0 Then TestResults(N) = 1
        End If
        Err.Clear
    
    
    On Error GoTo 0
    
    Dim i As Long
    Dim j As Long
    
    For i = 1 To TestNumber
        If TestResults(i) <> 1 Then
            Debug.Print ("Test #" & i & " failed!")
            j = j + 1
        End If
    Next i
    
    If j = 0 Then
        Debug.Print (FunctionName & ": All tests successfully passed!")
        If JA_InteractiveTesting Then MsgBox (FunctionName & ": All tests successfully passed!")
    Else
        Debug.Print (FunctionName & ": " & j & " tests failed!")
        If JA_InteractiveTesting Then MsgBox (FunctionName & ": " & j & " tests failed!")
    End If

End Sub


Private Sub test_DNAFindProteinInTemplate()

    Const TestNumber As Long = 5
    Const FunctionName As String = "DNAFindProteinInTemplate"
    
    Dim N As Long
    Dim TestResults(1 To TestNumber) As Long
    Dim Test As VBA.Collection
    Dim Input1 As String, Input2 As String, Input3 As String
    Dim t As String
    
    On Error Resume Next
    
    '1 empty inputs
        N = 1
        Input1 = ""
        Input2 = ""
        Set Test = DNAFindProteinInTemplate(Input1, Input2, False, False, False)
        If Err.Number = jaErr + 18 Then
            If Test Is Nothing Then TestResults(N) = 1
        End If
        Err.Clear
        
    '2 orf exists on both strands
        N = 2
        Input1 = "MKKKKKKKKKKKKKKKKKKKKKKKKK"
        t = "ATGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAATAG"
        Input2 = "CCC" & t & "CCCCCCCC" & DNAReverseComplement(t)
       
        Set Test = DNAFindProteinInTemplate(Input1, Input2, False, True, False)
        If Err.Number = jaErr + 3 Then
            If Err.Source = FunctionName Then
                If Test Is Nothing Then TestResults(N) = 1
            End If
        End If
        Err.Clear
        
    '3 orf exists on both strands, but only one is checked
        N = 3
        Input1 = "MKKKKKKKKKKKKKKKKKKKKKKKKK"
        t = "ATGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAATAG"
        Input2 = "CCC" & t & "CCCCCCCC" & DNAReverseComplement(t)

        Set Test = DNAFindProteinInTemplate(Input1, Input2, False, False, False)
        If Err.Number = 0 Then
            If Not Test Is Nothing Then
                If Test.Item(1) = 4 And Test.Item(2) = False Then TestResults(N) = 1
            End If
        End If
        Err.Clear
        
    '4 negative - orf exists in a circular context, but not linear one
        N = 4
        Input1 = "MKKKKKKKKKKKKKKKKKKKKKKKKK"
        t = "ATGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAATAG"
        Input2 = DNAReindex("CCC" & t & "CCCCCCCC", 15)
      
        Set Test = DNAFindProteinInTemplate(Input1, Input2, False, False, False)
        If Err.Number = jaErr + 1 Then
            If Err.Source = FunctionName Then
                If Test Is Nothing Then TestResults(N) = 1
            End If
        End If
        
        Err.Clear
    
    'positive - orf exists in a circular context, but not linear one
        N = 5
        Input1 = "MKKKKKKKKKKKKKKKKKKKKKKKKK"
        t = "ATGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAATAG"
        Input2 = DNAReindex("CCC" & t & "CCCCCCCC", 15)
        
        Set Test = DNAFindProteinInTemplate(Input1, Input2, True, False, False)
        If Err.Number = 0 Then
            If Not Test Is Nothing Then
                If Test.Item(1) = 82 And Test.Item(2) = False Then TestResults(N) = 1
            End If
        End If
        Err.Clear
        
       
        
        
        
    
    On Error GoTo 0
    
    Dim i As Long
    Dim j As Long
    
    For i = 1 To TestNumber
        If TestResults(i) <> 1 Then
            Debug.Print ("Test #" & i & " failed!")
            j = j + 1
        End If
    Next i
    
    If j = 0 Then
        Debug.Print (FunctionName & ": All tests successfully passed!")
        If JA_InteractiveTesting Then MsgBox (FunctionName & ": All tests successfully passed!")
    Else
        Debug.Print (FunctionName & ": " & j & " tests failed!")
        If JA_InteractiveTesting Then MsgBox (FunctionName & ": " & j & " tests failed!")
    End If

End Sub
