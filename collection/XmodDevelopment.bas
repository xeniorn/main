Attribute VB_Name = "XmodDevelopment"
Option Explicit

Sub testaaaa()

    Dim a
    
    Set a = CloningMakeConstructs_ParseTruncations("N50;C20;50-70+C20", 100)
    

End Sub

Sub testGibson()

    Dim a As Range
    
    Dim Gibson As clsGibsonAssembly
    
    
    
    If TypeOf Selection Is Excel.Range Then
        
        Set a = Selection
        
        Set Gibson = New clsGibsonAssembly
        
        Call Gibson.ImportRange(a)
        
        Call Gibson.Yes
        
        If Gibson.FragmentNumber > 0 Then
        
            Call OutputGibson(Gibson)
        
        Else
        
            MsgBox ("No outputs")
                
        End If
        
    End If
        
    Set Gibson = Nothing

End Sub

Sub OutputGibson(ByRef Gibson As clsGibsonAssembly)
    
    Const conOutputs As Long = 15
    
    Dim OutputSheet As Excel.Worksheet
    
    Dim G As clsGibsonSingleLigation
    
    Dim OutputRange As Excel.Range
    Dim AnnotRange As Excel.Range
    Dim OutputArray() As Variant
    Dim AnnotArray() As Variant
    Dim SheetName As String
    
    Dim i As Long
    
    SheetName = CreateSheetFromName("Gibson")
    
    Set OutputSheet = ActiveWorkbook.Worksheets(SheetName)
    
    Set OutputRange = OutputSheet.Cells(3, 3).Resize(conOutputs, Gibson.FragmentNumber)
    OutputArray = OutputRange.Value2
    
    Set AnnotRange = OutputRange.Offset(0, -1).Resize(conOutputs, 1)
    AnnotArray = AnnotRange.Value2
    
    For i = 1 To Gibson.FragmentNumber
        
        AnnotArray(1, 1) = "index"
        AnnotArray(2, 1) = "fragment name"
        AnnotArray(3, 1) = "fragment"
        AnnotArray(4, 1) = "primer_f name"
        AnnotArray(5, 1) = "primer_f"
        AnnotArray(6, 1) = "primer_f len"
        AnnotArray(7, 1) = "primer_f Tm"
        AnnotArray(8, 1) = "primer_r name"
        AnnotArray(9, 1) = "primer_r"
        AnnotArray(10, 1) = "primer_r len"
        AnnotArray(11, 1) = "primer_r Tm"
        AnnotArray(12, 1) = "overlap"
        AnnotArray(13, 1) = "overlap len"
        AnnotArray(14, 1) = "overlap Tm"
        AnnotArray(15, 1) = "overlap dG"
        
        With Gibson
            
            OutputArray(1, i) = Format(i, "00")
            With .PCR(i).FinalDNA
                OutputArray(2, i) = .Name
                OutputArray(3, i) = .Sequence
            End With
            With .PCR(i).ForwardPrimer
                OutputArray(4, i) = .Name
                OutputArray(5, i) = .Sequence
                OutputArray(6, i) = Len(.Sequence)
                OutputArray(7, i) = DNAAnnealToTemplate(.Sequence, Gibson.PCR(i).SourceDNA.Sequence)
            End With
            With .PCR(i).ReversePrimer
                OutputArray(8, i) = .Name
                OutputArray(9, i) = .Sequence
                OutputArray(10, i) = Len(.Sequence)
                OutputArray(11, i) = DNAAnnealToTemplate(DNAReverseComplement(.Sequence), Gibson.PCR(i).SourceDNA.Sequence)
            End With
            Set G = .Ligations.Item(i)
            With G
                OutputArray(12, i) = .Overlap
                OutputArray(13, i) = Len(.Overlap)
                OutputArray(14, i) = .Tm
                OutputArray(15, i) = .dG
            End With
            
        End With
        
    Next i
        
    OutputRange.Value2 = OutputArray
    AnnotRange.Value2 = AnnotArray
    
    AnnotRange.Columns.AutoFit
    AnnotRange.HorizontalAlignment = xlRight
    
    OutputRange.Borders.LineStyle = xlContinuous
    OutputRange.Borders.Weight = xlMedium
    
    OutputRange.WrapText = True
    
    OutputRange.ColumnWidth = 50
    OutputRange.RowHeight = 15
    
    OutputRange.HorizontalAlignment = xlLeft
    OutputRange.VerticalAlignment = xlTop
            
    Set G = Nothing
    Set OutputRange = Nothing
    Set OutputSheet = Nothing

End Sub

Sub testCMC()

    Dim tColl As VBA.Collection
    Dim pSeq As String
    Dim DNASeq As String
    Dim TruncList As String
    
    DNASeq = Range("H42").Value
    pSeq = Range("H43").Value
    'TruncList = "C50;C38;C27;C14;C4"
    TruncList = "C50;C4"
    
    Set tColl = CloningMakeConstructs(pSeq, DNASeq, TruncList)
    


End Sub




Private Function CloningMakeConstructs_ParseTruncations(ByVal TruncationList As String, ByVal ProteinLength As Long) As VBA.Collection

    Dim ConstructsArray() As String
    Dim MutationsArray() As String
    
    Dim i As Long
    Dim j As Long
    
    Dim tColl As VBA.Collection
    Dim tMutationObject As VBA.Collection
    Dim tConstructObject As VBA.Collection
    
    Dim ConstructsNumber As Long
    Dim MutationNumber As Long
    
    Dim Mutation As String
    
    Dim tempArray() As String
    
    Dim RegEx As RegExp
    
    Set RegEx = New RegExp
    
    ConstructsArray = VBA.Split(TruncationList, ";")
    
    ConstructsNumber = UBound(ConstructsArray) - LBound(ConstructsArray) + 1
    
    Set tColl = New VBA.Collection
    
    For i = 0 To ConstructsNumber - 1
    
        MutationsArray = Split(ConstructsArray(i), "+")
        MutationNumber = UBound(MutationsArray) - LBound(MutationsArray) + 1
        
        Set tConstructObject = New VBA.Collection
        
        For j = 0 To MutationNumber - 1
            
            Set tMutationObject = New VBA.Collection
            
            tMutationObject.Add "DEL", "TYPE"
                    
            Mutation = MutationsArray(j)
            
            '***************** identify and parse mutation
            
            RegEx.Pattern = "N[1-9]\d*"
            
            If RegEx.Test(Mutation) Then
            
                tMutationObject.Add 1, "START"
                tMutationObject.Add Int(Right(Mutation, Len(Mutation) - 1)), "END"
                
            Else
                    
                RegEx.Pattern = "C[1-9]\d*"
                
                If RegEx.Test(Mutation) Then
                
                    tMutationObject.Add ProteinLength + 1 - Int(Right(Mutation, Len(Mutation) - 1)), "START"
                    tMutationObject.Add ProteinLength, "END"
                    
                Else
                
                    RegEx.Pattern = "[1-9]\d*-[1-9]\d*"
                
                    If RegEx.Test(Mutation) Then
                        
                        tempArray = Split(Mutation, "-")
                        
                        If tempArray(0) > tempArray(1) Or tempArray(1) > ProteinLength Then
                            Call Err.Raise(vbError + 1, "CloningMakeConstructs_ParseTruncations", "Invalid input in truncation list")
                        End If
                        
                        tMutationObject.Add CLng(val(tempArray(0))), "START"
                        tMutationObject.Add CLng(val(tempArray(1))), "END"
                        
                    End If
                    
                End If
            
            End If
            
            tConstructObject.Add tMutationObject, Str(j)
            
        Next j
        
        tColl.Add tConstructObject, Str(i)
                
    Next i
    
    
    Set CloningMakeConstructs_ParseTruncations = tColl
    
    Set tColl = Nothing
    Set RegEx = Nothing
    Set tMutationObject = Nothing
    Set tConstructObject = Nothing
    
End Function

'************************************************************************************
Function CloningMakeConstructs( _
         ByVal ProteinSequence As String, _
         ByVal DNASource As String, _
         ByVal TruncationList As String, _
         Optional ByVal Circular As Boolean = True, _
         Optional ByVal CheckReverseComplement As Boolean = True, _
         Optional ByVal Interactive As Boolean = True _
         ) As VBA.Collection

'====================================================================================================
'Takes in a protein sequence, DNA source, and list of truncations to introduce
'Formulates a cloning strategy - fragments to clone out + primers to get these fragments from the soruce
'Juraj Ahel, 2017-01-03
'====================================================================================================
'works only for deltaN and deltaC constructs so far!
'2017-01-24 fix multiple bugs
    
'result is a collection: 1: DNA seq 2: protein seq 3: Gibson Assembly Object
    
    Const MyName As String = "CloningMakeConstructs"
        
'TODO: parse protein / DNA sequences / Trunc list
    
    Dim i As Long
    Dim j As Long
    
    Dim ProteinLength As Long
    Dim DNALength As Long
    Dim ORFLength As Long
    
    Dim ORFLocus As Long
    Dim IsReverse As Boolean
    
    Dim ORF As String
    Dim Base As String
    
    Dim ConstructNumber As Long
    
    Dim tColl As VBA.Collection
    Dim tTruncations As VBA.Collection
    Dim tResults As VBA.Collection
    
    Dim Gibson As clsGibsonAssembly
    Dim Ligation As clsGibsonSingleLigation
    Dim GI As clsGibsonInput
    Dim SourcePlasmid As clsDNA
    
    Dim tDNA As clsDNA
    
    'If Not Circular Then
    '    Err.Raise vbError + 2, "CloningMakeConstructs", "Non-circular inputs not yet supported"
    '
    
    ProteinLength = Len(ProteinSequence)
    DNALength = Len(DNASource)
    ORFLength = 3 * ProteinLength + 3
    
    If ProteinLength = 0 Or DNALength = 0 Then
        Call ApplyNewError(jaErr + 18, MyName, "Empty input")
        If Interactive Then
            ErrReraise
        Else
            Set CloningMakeConstructs = Nothing
            Exit Function
        End If
    End If
        
    '1 confirm DNA encodes for full protein
    
    Set tColl = DNAFindProteinInTemplate(ProteinSequence, DNASource, Circular, CheckReverseComplement, False)
    
    If Err.Number <> 0 Then
        If Err.Source = "DNAFindProteinInTemplate" Then
            Err.Source = MyName
        End If
        If Interactive Then
            ErrReraise
        Else
            Set CloningMakeConstructs = Nothing
            Exit Function
        End If
    Else
        If Not tColl Is Nothing Then
            If tColl.Count = 2 Then
                ORFLocus = tColl.Item(1)
                IsReverse = tColl.Item(2)
            End If
        End If
        If IsReverse Then DNASource = DNAReverseComplement(DNASource)
        ORF = Left(DNAReindex(DNASource, ORFLocus), 3 * ProteinLength)
        Base = Right(DNAReindex(DNASource, ORFLocus), DNALength - 3 * ProteinLength)
    End If
        
    Debug.Print ("DNA encodes for protein at locus: " & ORFLocus & " Reverse Strand: " & IsReverse)
       
       
    '2 formulate truncated sequences
    '3 in silico truncate DNA
    
    Set tTruncations = CloningMakeConstructs_ParseTruncations(TruncationList, ProteinLength)
       
    If Not tTruncations Is Nothing Then
    
        Set tResults = New VBA.Collection
        
        '1: DNA seq
        '2: protein seq
        
        ConstructNumber = tTruncations.Count
        
        For i = 1 To ConstructNumber
            Set tColl = New VBA.Collection
            tColl.Add CloningMakeConstructs_ApplyTruncations(tTruncations.Item(i), ORF)
            tColl.Add DNATranslate(tColl.Item(1))
            tResults.Add tColl
        Next i
       
    End If
    
    
    
    
    Set SourcePlasmid = New clsDNA
    Call SourcePlasmid.Define(Name:="SourcePlasmid", Sequence:=DNASource, Circular:=True)
        
    For i = 1 To ConstructNumber
            
        '4 design Gibson assembly
                        
        'Gibson inputs:
        
            Set tColl = New VBA.Collection
                
            'insert
            Set tDNA = New clsDNA
            With tDNA
                .Name = "Insert_" & Format(i, "00")
                .Sequence = tResults.Item(i).Item(1)
            End With
            
            Set GI = New clsGibsonInput
            With GI
                Set .Source = SourcePlasmid
                Set .InsertBefore = tDNA.DefineNew()
                Set .InsertAfter = tDNA.DefineNew()
                Set .Fragment = tDNA
                .ForbiddenRegions = "1"
            End With
            
            tColl.Add GI
            
            
            'backbone
            Set tDNA = New clsDNA
            With tDNA
                .Name = "Backbone_" & Format(i, "00")
                .Sequence = Base
            End With
            
            Set GI = New clsGibsonInput
            With GI
                Set .Source = SourcePlasmid
                Set .InsertBefore = tDNA.DefineNew()
                Set .InsertAfter = tDNA.DefineNew()
                Set .Fragment = tDNA
                .ForbiddenRegions = "3"
            End With
            
            tColl.Add GI
            
        
        Set Gibson = New clsGibsonAssembly
        
        Call Gibson.ImportCollection(tColl)
            
        '5 confirm PCR / gibson / translation of assembly
            
        Gibson.Yes
        
        tResults.Item(i).Add Gibson
        
    Next i
        
            
    'result is a collection: 1: DNA seq 2: protein seq 3: Gibson Assembly Object
    
    
    Dim tOutput As VBA.Collection
    Dim PrimColl As clsDNAs
    Dim NewPrim As clsDNAs
    
    
    Dim tIndex As Long
    Dim tName As String
    
    Set tOutput = New VBA.Collection
    Set PrimColl = New clsDNAs
    Set NewPrim = New clsDNAs
    
    
    
    'figure out the primers I will need
    
    For j = 1 To ConstructNumber
        
        'Forward
        Set tColl = New VBA.Collection
        tColl.Add j
        tColl.Add tResults.Item(j).Item(2)
        Set Gibson = tResults.Item(j).Item(3)
        
        For i = 1 To Gibson.FragmentNumber
            tIndex = tempCheckPrimer1(Gibson.PCR(i).ForwardPrimer, PrimColl)
            If tIndex > 0 Then
                Set Gibson.PCR(i).ForwardPrimer = PrimColl.DNA(tIndex)
            Else
                tName = tempCheckPrimer(Gibson.PCR(i).ForwardPrimer)
                If Len(tName) > 0 Then
                    Set tDNA = New clsDNA
                    With tDNA
                        .Sequence = Gibson.PCR(i).ForwardPrimer.Sequence
                        .Name = tName
                    End With
                    Call PrimColl.AddDNA(tDNA)
                Else
                    Set tDNA = New clsDNA
                    With tDNA
                        .Sequence = Gibson.PCR(i).ForwardPrimer.Sequence
                        .Name = "JA" & GetLastID + NewPrim.Count + 1
                    End With
                    Call PrimColl.AddDNA(tDNA)
                    Call NewPrim.AddDNA(tDNA)
                    Set Gibson.PCR(i).ForwardPrimer = tDNA
                End If
            End If
        Next i
        
        'reverse
        For i = 1 To Gibson.FragmentNumber
            tIndex = tempCheckPrimer1(Gibson.PCR(i).ReversePrimer, PrimColl)
            If tIndex > 0 Then
                Set Gibson.PCR(i).ReversePrimer = PrimColl.DNA(tIndex)
            Else
                tName = tempCheckPrimer(Gibson.PCR(i).ReversePrimer)
                If Len(tName) > 0 Then
                    Set tDNA = New clsDNA
                    With tDNA
                        .Sequence = Gibson.PCR(i).ReversePrimer.Sequence
                        .Name = tName
                    End With
                    Call PrimColl.AddDNA(tDNA)
                Else
                    Set tDNA = New clsDNA
                    With tDNA
                        .Sequence = Gibson.PCR(i).ReversePrimer.Sequence
                        .Name = "JA" & GetLastID + NewPrim.Count + 1
                    End With
                    Call PrimColl.AddDNA(tDNA)
                    Call NewPrim.AddDNA(tDNA)
                    Set Gibson.PCR(i).ReversePrimer = tDNA
                End If
            End If
        Next i
        
               
    Next j
                
    
    
    
    Set CloningMakeConstructs = tResults
    
    
    Set Gibson = Nothing
    Set tResults = Nothing
    Set tTruncations = Nothing
    Set tColl = Nothing
    Set Ligation = Nothing
    Set GI = Nothing
    Set SourcePlasmid = Nothing
    
    
    Set tOutput = Nothing
    Set PrimColl = Nothing
    Set NewPrim = Nothing
    

    End Function


Private Function tempCheckPrimer1(DNA As clsDNA, PrimColl As clsDNAs) As Long
    
    Dim i As Long
    
    If Not PrimColl Is Nothing Then
    
        For i = 1 To PrimColl.Count
            If DNA.Sequence = PrimColl.DNA(i).Sequence Then
                tempCheckPrimer1 = i
                Exit For
            End If
        Next i
        
    End If

End Function

Private Function tempCheckPrimer(DNA As clsDNA) As String
    
    Const conPrimersName As String = "tempPrimers"
    Const conMax As Long = 1000
    
    Dim Primers As Excel.Range
    Dim PrimersName As String
    Dim PrimersArray() As Variant
    
    Dim i As Long
            
    PrimersName = conPrimersName
    
    Set Primers = ActiveWorkbook.Worksheets(PrimersName).Cells(1, 1).Resize(conMax, 3)
    PrimersArray = Primers.Value2
    
    For i = 1 To conMax
        If DNA.Sequence = PrimersArray(i, 3) Then
            tempCheckPrimer = PrimersArray(i, 1)
            Exit For
        End If
    Next i
    
    
End Function

Private Function GetLastID() As Long
    
    Const conMax As Long = 1000
    Const conPrimersName As String = "tempPrimers"
    
    Dim RegEx As RegExp
    Dim tempIndex As Long
    Dim maxIndex As Long
    Dim PrimersName As String
    
    Dim i As Long
    
    Dim IDs As Variant
    
    PrimersName = conPrimersName
    IDs = ActiveWorkbook.Worksheets(PrimersName).Cells(1, 1).Resize(conMax, 1).Value2
        
    Set RegEx = New RegExp
    RegEx.Pattern = "^JA(\d{3,4})$"
    
        
    For i = LBound(IDs, 1) To UBound(IDs, 1)
    
        If Len(IDs(i, 1)) = 0 Then Exit For
        
        maxIndex = 0
        
        If RegEx.Test(IDs(i, 1)) Then
            tempIndex = RegEx.Replace(IDs(i, 1), "$1")
            If tempIndex > maxIndex Then maxIndex = tempIndex
        End If
    
    Next i
    
    GetLastID = maxIndex


End Function

Private Function CloningMakeConstructs_ApplyTruncations(ByVal TruncCollection As VBA.Collection, ByVal DNASequence As String) As String

    Dim i As Long
    Dim SeqArray() As String
    Dim tColl As VBA.Collection
    
    ReDim SeqArray(1 To Len(DNASequence))
    
    'put the DNA seq into an array
    For i = 1 To Len(DNASequence)
        SeqArray(i) = Mid(DNASequence, i, 1)
    Next i
    
    'for each protein seq range to truncate
    For Each tColl In TruncCollection
        
        'delete bases associated with given protein truncation ranges
        For i = (-2 + 3 * tColl.Item("START")) To (3 * tColl.Item("END"))
            SeqArray(i) = ""
        Next i
    
    Next tColl
        
    CloningMakeConstructs_ApplyTruncations = Join(SeqArray, "")

End Function


'************************************************************************************
Function SmoothData(DataArray As Variant, WindowSize As Long) As Variant

    Dim DataLength As Long
    Dim TempOutput() As Variant
    Dim tempIndex As Long
    Dim i As Long, j As Long
    Dim tempsum As Double
        
        
    DataLength = 1 + UBound(DataArray) - LBound(DataArray)
    
    ReDim TempOutput(1 To DataLength)
        
    
            
    For i = 1 To DataLength - WindowSize
        
        tempIndex = i + WindowSize \ 2
        tempsum = 0
        
        For j = i To i + WindowSize - 1
            tempsum = tempsum + DataArray(j)
        Next j
        
        TempOutput(tempIndex) = tempsum / WindowSize
        
    Next i
        
        
    For i = 1 To WindowSize \ 2
        TempOutput(i) = 0
        TempOutput(DataLength - i + 1) = 0
    Next i
    
    SmoothData = TempOutput

End Function

'************************************************************************************
Function GetMaxLetterCount(InputString As String) As Long

    Dim i As Byte
    Dim Char As String
    Dim tout As String
    Dim tempCount As Long
    Dim CharCount As Long
    
    tempCount = 0
    
    For i = 65 To 90
    
        Char = Chr(i)
        CharCount = StringCharCount(InputString, Char)
        If CharCount > tempCount Then tempCount = CharCount
        
    Next i
    
    GetMaxLetterCount = tempCount
        

End Function




