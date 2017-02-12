Attribute VB_Name = "XmodDevelopment"
Option Explicit

Public Sub testaaaa()

    Dim a
    
    Set a = CloningMakeConstructs_ParseTruncations("N50;C20;50-70+C20", 100)
    

End Sub

Public Sub testGibson()

    Dim a As Range
    
    Dim Gibson As clsGibsonAssembly
    
    
    
    If TypeOf Selection Is Excel.Range Then
        
        Set a = Selection
        
        Set Gibson = New clsGibsonAssembly
        
        Call Gibson.ImportRange(a)
        
        Call Gibson.Yes
        
        If Gibson.FragmentNumber > 0 Then
        
            Call CMC_OutputGibson(Gibson)
        
        Else
        
            MsgBox ("No outputs")
                
        End If
        
    End If
        
    Set Gibson = Nothing

End Sub

Private Sub CMC_OutputGibson(ByVal Gibson As clsGibsonAssembly)
    
    Const conOutputs As Long = 19
    Const conSummaryOutputs = 6
    
    Dim OutputSheet As Excel.Worksheet
    
    Dim G As clsGibsonSingleLigation
    
    Dim OutputRange As Excel.Range
    Dim AnnotRange As Excel.Range
    Dim SummaryRange As Excel.Range
    Dim SummaryAnnotRange As Excel.Range
    
    Dim OutputArray() As Variant
    Dim AnnotArray() As Variant
    Dim SummaryArray() As Variant
    Dim SummaryAnnot() As Variant
    
    Dim SheetName As String
    
    Dim i As Long
    
    SheetName = CreateSheetFromName(Gibson.Name)
    
    Set OutputSheet = ActiveWorkbook.Worksheets(SheetName)
    
    Set OutputRange = OutputSheet.Cells(3, 3).Resize(conOutputs, Gibson.FragmentNumber)
    OutputArray = OutputRange.Value2
    
    Set AnnotRange = OutputRange.Offset(0, -1).Resize(conOutputs, 1)
    AnnotArray = AnnotRange.Value2
    
    Set SummaryRange = OutputRange.Offset(conOutputs + 1, 0).Resize(conSummaryOutputs, 1)
    SummaryArray = SummaryRange.Value2
    
    Set SummaryAnnotRange = SummaryRange.Offset(0, -1).Resize(conSummaryOutputs, 1)
    SummaryAnnot = SummaryAnnotRange.Value2
    
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
        AnnotArray(16, 1) = "source name"
        AnnotArray(17, 1) = "source"
        AnnotArray(18, 1) = "primerF anneal"
        AnnotArray(19, 1) = "primerR anneal"
        
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
                OutputArray(7, i) = DNAAnnealToTemplate(.Sequence, Gibson.PCR(i).SourceDNA.Sequence).Item("TM")
                OutputArray(18, i) = DNAAnnealToTemplate(.Sequence, Gibson.PCR(i).SourceDNA.Sequence).Item("SEQ")
            End With
            With .PCR(i).ReversePrimer
                OutputArray(8, i) = .Name
                OutputArray(9, i) = .Sequence
                OutputArray(10, i) = Len(.Sequence)
                OutputArray(11, i) = DNAAnnealToTemplate(DNAReverseComplement(.Sequence), Gibson.PCR(i).SourceDNA.Sequence).Item("TM")
                OutputArray(19, i) = DNAAnnealToTemplate(DNAReverseComplement(.Sequence), Gibson.PCR(i).SourceDNA.Sequence).Item("SEQ")
            End With
            Set G = .Ligations.Item(i)
            With G
                OutputArray(12, i) = .Overlap
                OutputArray(13, i) = Len(.Overlap)
                OutputArray(14, i) = .Tm
                OutputArray(15, i) = .dG
            End With
            
            OutputArray(16, i) = .SourceDNA(i).Name
            OutputArray(17, i) = .SourceDNA(i).Sequence
            
        End With
        
    Next i
    
    SummaryAnnot(1, 1) = "Name"
    SummaryAnnot(2, 1) = "Final sequence"
    SummaryAnnot(3, 1) = "DNA length"
    SummaryAnnot(4, 1) = "Longest ORF"
    SummaryAnnot(5, 1) = "Translation"
    SummaryAnnot(6, 1) = "Protein Length"
    
    With Gibson
        SummaryArray(1, 1) = .Name
        SummaryArray(2, 1) = .FinalAssembly.Sequence
        SummaryArray(3, 1) = Len(.FinalAssembly.Sequence)
        SummaryArray(4, 1) = DNALongestORF(.FinalAssembly.Sequence)
        SummaryArray(5, 1) = DNATranslate(SummaryArray(4, 1))
        SummaryArray(6, 1) = Len(SummaryArray(5, 1))
    End With
            
        
    OutputRange.Value2 = OutputArray
    AnnotRange.Value2 = AnnotArray
    SummaryRange.Value2 = SummaryArray
    SummaryAnnotRange.Value2 = SummaryAnnot
    
    With AnnotRange
        .Columns.AutoFit
        .HorizontalAlignment = xlRight
    End With
    
    With SummaryAnnotRange
        .HorizontalAlignment = xlRight
    End With
    
    With SummaryRange
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        .WrapText = True
        .ColumnWidth = 50
        .RowHeight = 15
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With
    
    
    Dim FormatSelect As String
    Dim FormatArray() As String
    FormatSelect = "3 5 9 12 17 18 19 22 24 25"
    FormatArray = Split(FormatSelect, " ")
    
    With OutputRange
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        .WrapText = True
        .ColumnWidth = 50
        .RowHeight = 15
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        For i = LBound(FormatArray) To UBound(FormatArray)
            .Rows(FormatArray(i)).Font.Name = "Consolas"
        Next i
    End With
    
    
            
    Set G = Nothing
    Set OutputRange = Nothing
    Set OutputSheet = Nothing

End Sub

Private Sub CMC_OutputPrimers(ByVal Primers As clsDNAs)
    
    Const conOutputs As Long = 4
    
    Dim OutputSheet As Excel.Worksheet
    
    
    Dim OutputRange As Excel.Range
    Dim AnnotRange As Excel.Range
    Dim OutputArray() As Variant
    Dim AnnotArray() As Variant
    Dim SheetName As String
    
    Dim i As Long
    
    If Primers.Count = 0 Then Exit Sub
    
    SheetName = CreateSheetFromName("NewPrimers")
    
    Set OutputSheet = ActiveWorkbook.Worksheets(SheetName)
    
    Set OutputRange = OutputSheet.Cells(3, 3).Resize(Primers.Count, conOutputs)
    OutputArray = OutputRange.Value2
    
    Set AnnotRange = OutputRange.Offset(-1, 0).Resize(1, conOutputs)
    AnnotArray = AnnotRange.Value2
    
    AnnotArray(1, 1) = "ID"
    AnnotArray(1, 2) = "description"
    AnnotArray(1, 3) = "sequence"
    AnnotArray(1, 4) = "length"
    
    With Primers
        For i = 1 To .Count
            With .DNA(i)
                OutputArray(i, 1) = Split(.Name, "#")(0)
                OutputArray(i, 2) = Split(.Name, "#")(1)
                OutputArray(i, 3) = .Sequence
                OutputArray(i, 4) = Len(.Sequence)
            End With
        Next i
    End With
        
    OutputRange.Value2 = OutputArray
    AnnotRange.Value2 = AnnotArray
    
    OutputRange.Offset(-1).Resize(AnnotRange.Rows.Count + 1).Columns.AutoFit
    AnnotRange.HorizontalAlignment = xlCenter
    
    OutputRange.Borders.LineStyle = xlContinuous
    OutputRange.Borders.Weight = xlMedium
    
    OutputRange.WrapText = True
    
    OutputRange.Columns(3).ColumnWidth = 50
    OutputRange.Columns(3).Font.Name = "Consolas"
    OutputRange.RowHeight = 15
    
    OutputRange.HorizontalAlignment = xlLeft
    OutputRange.VerticalAlignment = xlTop
            
    Set OutputRange = Nothing
    Set OutputSheet = Nothing

End Sub

Private Sub CMC_OutputProteins(ByVal Proteins As VBA.Collection)
    
    Const conOutputs As Long = 3
    
    Dim OutputSheet As Excel.Worksheet
    
    
    Dim OutputRange As Excel.Range
    Dim AnnotRange As Excel.Range
    Dim OutputArray() As Variant
    Dim AnnotArray() As Variant
    Dim SheetName As String
    
    Dim i As Long
    
    SheetName = CreateSheetFromName("NewProteins")
    
    Set OutputSheet = ActiveWorkbook.Worksheets(SheetName)
    
    Set OutputRange = OutputSheet.Cells(3, 3).Resize(Proteins.Count, conOutputs)
    OutputArray = OutputRange.Value2
    
    Set AnnotRange = OutputRange.Offset(-1, 0).Resize(1, conOutputs)
    AnnotArray = AnnotRange.Value2
    
    AnnotArray(1, 1) = "ID"
    AnnotArray(1, 2) = "sequence"
    AnnotArray(1, 3) = "length"
    
    With Proteins
        For i = 1 To .Count
            With .Item(i)
                OutputArray(i, 1) = .Item(1)
                OutputArray(i, 2) = .Item(2)
                OutputArray(i, 3) = Len(.Item(2))
            End With
        Next i
    End With
        
    OutputRange.Value2 = OutputArray
    AnnotRange.Value2 = AnnotArray
    
    OutputRange.Offset(-1).Resize(AnnotRange.Rows.Count + 1).Columns.AutoFit
    AnnotRange.HorizontalAlignment = xlCenter
    
    OutputRange.Borders.LineStyle = xlContinuous
    OutputRange.Borders.Weight = xlMedium
    
    OutputRange.WrapText = True
    
    OutputRange.Columns(2).Font.Name = "Consolas"
    OutputRange.Columns(2).ColumnWidth = 50
    OutputRange.RowHeight = 15
    
    OutputRange.HorizontalAlignment = xlLeft
    OutputRange.VerticalAlignment = xlTop
            
    Set OutputRange = Nothing
    Set OutputSheet = Nothing

End Sub


Private Sub CMC_OutputAssemblies(ByVal Assemblies As VBA.Collection)
    
    'TODO: give automatic pJAxxxx names to assemblies!
    
    Const conOutputs As Long = 6
        
    Dim OutputSheet As Excel.Worksheet
    
    
    Dim OutputRange As Excel.Range
    Dim AnnotRange As Excel.Range
    Dim OutputArray() As Variant
    Dim AnnotArray() As Variant
    Dim SheetName As String
    Dim DataCount As Long
    
    Dim Gibson As clsGibsonAssembly
    
    Dim i As Long
    
    DataCount = Assemblies.Count
    SheetName = CreateSheetFromName("NewAssemblies")
    
    Set OutputSheet = ActiveWorkbook.Worksheets(SheetName)
    
    Set OutputRange = OutputSheet.Cells(3, 3).Resize(DataCount, conOutputs)
    OutputArray = OutputRange.Value2
    
    Set AnnotRange = OutputRange.Offset(-1, 0).Resize(1, conOutputs)
    AnnotArray = AnnotRange.Value2
    
    AnnotArray(1, 1) = "ID"
    AnnotArray(1, 2) = "sequence"
    AnnotArray(1, 3) = "length"
    AnnotArray(1, 4) = "insert fragment"
    AnnotArray(1, 5) = "backbone fragment"
    AnnotArray(1, 6) = "protein sequence"
        
    With Assemblies
        For i = 1 To .Count
            
            Set Gibson = .Item(i)
            
            With Gibson
                OutputArray(i, 1) = .Name
                OutputArray(i, 2) = .FinalAssembly.Sequence
                OutputArray(i, 3) = .FinalAssembly.Length
                OutputArray(i, 4) = .FinalDNA(1).Name
                OutputArray(i, 5) = .FinalDNA(2).Name
                OutputArray(i, 6) = DNATranslate(DNALongestORF(.FinalAssembly.Sequence))
            End With
            
        Next i
    End With
        
    OutputRange.Value2 = OutputArray
    AnnotRange.Value2 = AnnotArray
    
    OutputRange.Offset(-1).Resize(AnnotRange.Rows.Count + 1).Columns.AutoFit
    AnnotRange.HorizontalAlignment = xlCenter
    
    OutputRange.Borders.LineStyle = xlContinuous
    OutputRange.Borders.Weight = xlMedium
    
    OutputRange.WrapText = True
        
    OutputRange.Columns(2).ColumnWidth = 50
    OutputRange.Columns(2).Font.Name = "Consolas"
    OutputRange.Columns(6).ColumnWidth = 50
    OutputRange.Columns(6).Font.Name = "Consolas"
    OutputRange.RowHeight = 15
    
    OutputRange.HorizontalAlignment = xlLeft
    OutputRange.VerticalAlignment = xlTop
            
    Set OutputRange = Nothing
    Set OutputSheet = Nothing
    Set Gibson = Nothing

End Sub


Private Sub CMC_OutputFragments(ByVal Fragments As clsDNAs, ByVal PCRPrimers As VBA.Collection)
    
    Const conOutputs As Long = 6
        
    Dim OutputSheet As Excel.Worksheet
    
    
    Dim OutputRange As Excel.Range
    Dim AnnotRange As Excel.Range
    Dim OutputArray() As Variant
    Dim AnnotArray() As Variant
    Dim SheetName As String
    
    Dim tColl As VBA.Collection
    
    Dim i As Long
    
    SheetName = CreateSheetFromName("NewFragments")
    
    Set OutputSheet = ActiveWorkbook.Worksheets(SheetName)
    
    Set OutputRange = OutputSheet.Cells(3, 3).Resize(Fragments.Count, conOutputs)
    OutputArray = OutputRange.Value2
    
    Set AnnotRange = OutputRange.Offset(-1, 0).Resize(1, conOutputs)
    AnnotArray = AnnotRange.Value2
    
    AnnotArray(1, 1) = "ID"
    AnnotArray(1, 2) = "sequence"
    AnnotArray(1, 3) = "length"
    AnnotArray(1, 4) = "source"
    AnnotArray(1, 5) = "fwd"
    AnnotArray(1, 6) = "rev"
    
    With Fragments
        For i = 1 To .Count
        
            With .DNA(i)
                OutputArray(i, 1) = Split(.Name, "#")(0)
                OutputArray(i, 2) = .Sequence
                OutputArray(i, 3) = Len(.Sequence)
            End With
            
            Set tColl = PCRPrimers.Item(i)
            OutputArray(i, 4) = tColl.Item(1)
            OutputArray(i, 5) = tColl.Item(2)
            OutputArray(i, 6) = tColl.Item(3)
            
        Next i
    End With
        
    OutputRange.Value2 = OutputArray
    AnnotRange.Value2 = AnnotArray
    
    OutputRange.Offset(-1).Resize(AnnotRange.Rows.Count + 1).Columns.AutoFit
    AnnotRange.HorizontalAlignment = xlCenter
    
    OutputRange.Borders.LineStyle = xlContinuous
    OutputRange.Borders.Weight = xlMedium
    
    OutputRange.WrapText = True
        
    OutputRange.Columns(2).ColumnWidth = 50
    OutputRange.Columns(2).Font.Name = "Consolas"
    OutputRange.RowHeight = 15
    
    OutputRange.HorizontalAlignment = xlLeft
    OutputRange.VerticalAlignment = xlTop
            
    Set OutputRange = Nothing
    Set OutputSheet = Nothing

End Sub


Public Sub CMC_testFromRange()
    
    Const conHeaderBlocks As Long = 3
    
    Dim inputRange As Excel.Range
    Dim DNARange As Excel.Range
    Dim ProtRange As Excel.Range
    Dim ConstructsRange As Excel.Range
    Dim DataBlock As Excel.Range
    
    Dim insSource As clsDNA
    Dim vecSource As clsDNA
    
    Dim i As Long
    
    Dim pSeq As String
    Dim DNASeq As String
    Dim VectorDNASeq As String
    
    Dim TruncList As String
    Dim NameList As String
    Dim ForbidList As String
    
    Dim SourceName As String
    Dim VectorSourceName As String
    
    Dim N As Long
    
    Dim dataArray() As Variant
    
    Dim tColl As VBA.Collection
    
    Dim Stopwatch As clsTimer
    
    Set Stopwatch = New clsTimer
    
    Stopwatch.Trigger
    
    'DNA source insert | NAME
    'DNA source vector | NAME
    'protein
    'name | trunc | forbid
    'name | trunc | forbid
    'name | trunc | forbid
    '...
    
    If TypeOf Selection Is Excel.Range Then
        Set inputRange = Selection
    Else
        Exit Sub
    End If
    
    N = inputRange.Rows.Count - conHeaderBlocks
    
    'HEADER
        Set DNARange = inputRange.Offset(0, 0).Resize(1, 1)
        DNASeq = DNARange.Value2
        SourceName = DNARange.Offset(0, 1).Value2
        
        Set DNARange = inputRange.Offset(1, 0).Resize(1, 1)
        VectorDNASeq = DNARange.Value2
        VectorSourceName = DNARange.Offset(0, 1).Value2
        
        Set ProtRange = inputRange.Offset(2, 0).Resize(1, 1)
        pSeq = ProtRange.Value2
        
        Set ConstructsRange = inputRange.Offset(conHeaderBlocks, 0).Resize(N, 3)
                    
    'DATA
        dataArray = ConstructsRange.Value2
                   
        Set DataBlock = ConstructsRange.Offset(0, 0).Resize(N, 1)
        NameList = RangeJoin(DataBlock, ";")
        Set DataBlock = ConstructsRange.Offset(0, 1).Resize(N, 1)
        TruncList = RangeJoin(DataBlock, ";")
        Set DataBlock = ConstructsRange.Offset(0, 2).Resize(N, 1)
        ForbidList = RangeJoin(DataBlock, ";")
        
        Set insSource = New clsDNA
        Call insSource.Define(SourceName, DNASeq, Circular:=True)
        Set vecSource = New clsDNA
        Call vecSource.Define(VectorSourceName, VectorDNASeq, Circular:=True)
    
    'CALCULATION & EXPORT
        Set tColl = CloningMakeConstructs(pSeq, insSource, vecSource, TruncList, NameList, ForbidList)
    
        Stopwatch.Trigger
        
        Debug.Print ("Sequences:" & N & " Runtime: " & Round(Stopwatch.Duration, 1) & " s")
    
    'CLEANUP
        Set insSource = Nothing
        Set vecSource = Nothing
        Set tColl = Nothing
        Set Stopwatch = Nothing
        
End Sub

Public Function RangeJoin(ByVal IR As Excel.Range, Optional ByVal Delimiter As String = "")

    Dim cell As Excel.Range
    Dim ar() As String
    Dim i As Long
    
    ReDim ar(1 To IR.Count)
    
    For Each cell In IR
        i = i + 1
        ar(i) = cell.Value2
    Next cell
    
    RangeJoin = Join(ar, Delimiter)

End Function

Public Sub CMC_test()

    Dim tColl As VBA.Collection
    Dim pSeq As String
    Dim DNASeq As String
    Dim TruncList As String
    Dim NameList As String
    Dim ForbidList As String
    
    DNASeq = Range("H42").Value
    pSeq = Range("H43").Value
    'TruncList = "C50;C38;C27;C14;C4"
    'TruncList = "C61;C49;C38;C25;C15;C11"
    'NameList = "5a;6a;7a;8a;9a;10a"
    
    TruncList = "N50"
    NameList = "5a"
    ForbidList = "1 "
    'ForbidList = ""
    
    'Set tColl = CloningMakeConstructs(pSeq, DNASeq, "PLSX", TruncList, NameList, ForbidList)
    


End Sub

Private Function CloningMakeConstructs_ParseNames(ByVal NameList As String) As VBA.Collection

    Dim i As Long
    Dim N As Long
    Dim NameArray() As String
    Dim tColl As VBA.Collection
    
    NameArray = Split(NameList, ";")
    
    Set tColl = New VBA.Collection
    
    For i = LBound(NameArray) To UBound(NameArray)
        tColl.Add NameArray(i)
    Next i
    
    
    Set CloningMakeConstructs_ParseNames = tColl
    Set tColl = Nothing

End Function

Private Function CloningMakeConstructs_ParseForbid(ByVal ForbidList As String) As VBA.Collection

    Dim i As Long
    Dim ForbidArray() As String
    Dim tColl As VBA.Collection
    Dim tColl1 As VBA.Collection
    
    If ForbidList = vbNullString Then
        ReDim ForbidArray(0 To 0)
        ForbidArray(0) = ""
    Else
        ForbidArray = Split(ForbidList, ";")
    End If
    
    Set tColl = New VBA.Collection
    
    For i = LBound(ForbidArray) To UBound(ForbidArray)
        Set tColl1 = New VBA.Collection
        With tColl1
            If ForbidArray(i) = vbNullString Then
                .Add ""
                .Add ""
            Else
                .Add Split(ForbidArray(i), " ")(0)
                .Add Split(ForbidArray(i), " ")(1)
            End If
        End With
        tColl.Add tColl1
    Next i
    
    
    Set CloningMakeConstructs_ParseForbid = tColl
    Set tColl = Nothing
    Set tColl1 = Nothing

End Function


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
                
                If Int(Right(Mutation, Len(Mutation) - 1)) <= 1 Then Call Err.Raise(jaErr + 1, "CMC_ParseTruncations", "unimplemented input")
            
                tMutationObject.Add 2, "START"
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
Public Function CloningMakeConstructs( _
         ByVal ProteinSequence As String, _
         ByRef InsertSourceDNA As clsDNA, _
         ByRef VectorSourceDNA As clsDNA, _
         ByVal TruncationList As String, _
         ByVal NameList As String, _
         ByVal ForbidList As String, _
         Optional ByVal Circular As Boolean = True, _
         Optional ByVal CheckReverseComplement As Boolean = True, _
         Optional ByVal Interactive As Boolean = True _
         ) As VBA.Collection

'====================================================================================================
'Takes in a protein sequence, DNA source, and list of truncations to introduce
'Formulates a cloning strategy - fragments to clone out + primers to get these fragments from the soruce
'Juraj Ahel, 2017-01-03
'result is a collection: 1: DNA seq 2: protein seq 3: Assembly Name 4: Gibson Assembly Object
'====================================================================================================
'works only for deltaN and deltaC constructs so far!
'2017-01-24 fix multiple bugs
'
'
'TODO: give automatic pJAxxxx names to assemblies!
    
    
    Const MyName As String = "CloningMakeConstructs"
    
    Dim i As Long
    Dim j As Long
    
    Dim ProteinLength As Long
    Dim ORFLength As Long
    
    Dim ORFLocus As Long
    Dim IsReverse As Boolean
    
    Dim ORF As String
    Dim Base As String
    
    Dim ConstructNumber As Long
    
    Dim tColl As VBA.Collection
    
    Dim tTruncations As VBA.Collection
    Dim tNames As VBA.Collection
    Dim tForbid As VBA.Collection
        
    Dim tResults As VBA.Collection
    
    
    Dim Gibson As clsGibsonAssembly
    Dim Ligation As clsGibsonSingleLigation
    Dim GI As clsGibsonInput
    
    Dim tDNA As clsDNA
    
    'If Not Circular Then
    '    Err.Raise vbError + 2, "CloningMakeConstructs", "Non-circular inputs not yet supported"
    '
    
    ProteinLength = Len(ProteinSequence)
    ORFLength = 3 * ProteinLength + 3
    
    If ProteinLength = 0 Or InsertSourceDNA.Length = 0 Then
        Call ApplyNewError(jaErr + 18, MyName, "Empty input")
        If Interactive Then
            ErrReraise
        Else
            Set CloningMakeConstructs = Nothing
            Exit Function
        End If
    End If
        
    '1 confirm DNA encodes for full protein
    
    Set tColl = DNAFindProteinInTemplate(ProteinSequence, InsertSourceDNA.Sequence, Circular, CheckReverseComplement, False)
    
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
        If IsReverse Then InsertSourceDNA.Sequence = DNAReverseComplement(InsertSourceDNA.Sequence)
        ORF = Left(DNAReindex(InsertSourceDNA.Sequence, ORFLocus), 3 * ProteinLength)
        Base = Right(DNAReindex(InsertSourceDNA.Sequence, ORFLocus), InsertSourceDNA.Length - 3 * ProteinLength)
    End If
        
    Debug.Print ("DNA encodes for protein at locus: " & ORFLocus & " Reverse Strand: " & IsReverse)
       
       
    '2 formulate truncated sequences
    '3 in silico truncate DNA
    
    '================= parse inputs
    
    Set tTruncations = CloningMakeConstructs_ParseTruncations(TruncationList, ProteinLength)
    Set tNames = CloningMakeConstructs_ParseNames(NameList)
    Set tForbid = CloningMakeConstructs_ParseForbid(ForbidList)
    
    'check if the input is fine
        If Not tTruncations Is Nothing And Not tNames Is Nothing And Not tForbid Is Nothing Then
            If tTruncations.Count <> tNames.Count Or tTruncations.Count <> tForbid.Count Then
                Call ApplyNewError(jaErr + 19, MyName, "Number of names doesn't match the number of truncations given")
                If Interactive Then
                    ErrReraise
                Else
                    Set CloningMakeConstructs = Nothing
                    Exit Function
                End If
            End If
            
            
        Else
            Call ApplyNewError(jaErr + 20, MyName, "Error parsing truncations or names, check inputs")
            If Interactive Then
                ErrReraise
            Else
                Set CloningMakeConstructs = Nothing
                Exit Function
            End If
        End If
       
        
    Set tResults = New VBA.Collection
    
    '1: DNA seq
    '2: protein seq
    
    ConstructNumber = tTruncations.Count
    
    For i = 1 To ConstructNumber
        Set tColl = New VBA.Collection
        With tColl
            .Add CloningMakeConstructs_ApplyTruncations(tTruncations.Item(i), ORF)
            .Add DNATranslate(tColl.Item(1))
            .Add tNames.Item(i)
        End With
        tResults.Add tColl
    Next i
    
    
    If InsertSourceDNA.Name = vbNullString Then InsertSourceDNA.Name = "SourceDNA"
        
    For i = 1 To ConstructNumber
            
        '4 design Gibson assembly
                        
        '===================== PREPARE Gibson inputs:
            
        Set tColl = New VBA.Collection
            
            'insert
            Set tDNA = New clsDNA
            With tDNA
                .Name = "insert" '"Insert_" & Format(i, "00")
                .Sequence = tResults.Item(i).Item(1)
            End With
            
            Set GI = New clsGibsonInput
            With GI
                '.Name = "insert"
                Set .Source = InsertSourceDNA
                Set .InsertBefore = tDNA.DefineNew()
                Set .InsertAfter = tDNA.DefineNew()
                Set .Fragment = tDNA
                .ForbiddenRegions = tForbid.Item(i).Item(1)
            End With
            
            tColl.Add GI
            
            
            'backbone
            Set tDNA = New clsDNA
            With tDNA
                .Name = "vector" '"Backbone_" & Format(i, "00")
                .Sequence = Base
            End With
            
            Set GI = New clsGibsonInput
            With GI
                Set .Source = VectorSourceDNA
                Set .InsertBefore = tDNA.DefineNew()
                Set .InsertAfter = tDNA.DefineNew()
                Set .Fragment = tDNA
                .ForbiddenRegions = tForbid.Item(i).Item(2)
            End With
            
            tColl.Add GI
            
        
        '==============================PERFORM GIBSON CALCULATION
        
        Set Gibson = New clsGibsonAssembly
        
        With Gibson
            
            Call .ImportCollection(tColl)
            .Name = tNames.Item(i)
            '5 confirm PCR / gibson / translation of assembly
            .Yes
            
        End With
            
        tResults.Item(i).Add Gibson
        
    Next i
        
            
    'result is a collection: 1: DNA seq 2: protein seq 3: Gibson Assembly Object
    
    
    Dim tOutput As VBA.Collection
    Dim PrimColl As clsDNAs
    Dim NewPrim As clsDNAs
    Dim Frags As clsDNAs
    Dim PCRPrimers As VBA.Collection
    Dim Proteins As VBA.Collection
    Dim Assemblies As VBA.Collection
    
    Dim tIndex As Long
    Dim tName As String
    
    Dim IsNewPrimer As Boolean
    
    Dim tFragName As String
    
    Set tOutput = New VBA.Collection
    Set PrimColl = New clsDNAs
    Set NewPrim = New clsDNAs
    Set Frags = New clsDNAs
    Set PCRPrimers = New VBA.Collection
    Set Proteins = New VBA.Collection
    Set Assemblies = New VBA.Collection
    
    
    '===========CALC OUTPUTS===========
    'figure out the new primers I will need and which are duplicated
    'figure out which fragments are duplicated
    'export primers table, fragments table, proteins table
    
    For j = 1 To ConstructNumber
        
        'Set tColl = New VBA.Collection
        'construct #
        'tColl.Add j
        'protein sequence
        'tColl.Add tResults.Item(j).Item(2)
        'Gibson object
        
        Set Gibson = tResults.Item(j).Item(4)
                
        With Gibson
                    
            'figure out the new primers I will need and which are duplicated
                    
            For i = 1 To .FragmentNumber
                With .PCR(i)
                    
                    'forward
                        IsNewPrimer = False
                        tIndex = CMC_tempCheckPrimer1(.ForwardPrimer, PrimColl)
                        
                        If tIndex > 0 Then
                        
                            Set .ForwardPrimer = PrimColl.DNA(tIndex)
                            
                        Else
                        
                            tName = CMC_tempCheckPrimer(.ForwardPrimer)
                            
                            If Len(tName) = 0 Then
                                tName = "JA" & GetLastID + NewPrim.Count + 1 & "#" & Gibson.Name & "_" & .FinalDNA.Name & "_f"
                                IsNewPrimer = True
                            End If
                            
                            .ForwardPrimer.Name = tName
                            
                            If IsNewPrimer Then Call NewPrim.AddDNA(.ForwardPrimer)
                            Call PrimColl.AddDNA(.ForwardPrimer)
                            
                        End If
                    
                    
                    'reverse
                        IsNewPrimer = False
                        tIndex = CMC_tempCheckPrimer1(.ReversePrimer, PrimColl)
                        
                        If tIndex > 0 Then
                        
                            Set .ReversePrimer = PrimColl.DNA(tIndex)
                            
                        Else
                        
                            tName = CMC_tempCheckPrimer(.ReversePrimer)
                            
                            If Len(tName) = 0 Then
                                tName = "JA" & GetLastID + NewPrim.Count + 1 & "#" & Gibson.Name & "_" & .FinalDNA.Name & "_r"
                                IsNewPrimer = True
                            End If
                            
                            .ReversePrimer.Name = tName
                            
                            If IsNewPrimer Then Call NewPrim.AddDNA(.ReversePrimer)
                            Call PrimColl.AddDNA(.ReversePrimer)
                            
                        End If
                        
                        
                    'figure out which fragments are duplicated
                        tFragName = Gibson.Name & "_" & .FinalDNA.Name
                        
                        If Frags.ContainsDNASeq(.FinalDNA) Then
                        
                            .FinalDNA.Name = Frags.GetDNABySeq(.FinalDNA.Sequence).Name
                            
                        Else
                            
                            'fill collection of new fragments
                            .FinalDNA.Name = tFragName
                            Call Frags.AddDNA(.FinalDNA)
                            
                            'also add associated primers
                            Set tColl = New VBA.Collection
                                tColl.Add .SourceDNA.Name
                                tColl.Add Split(.ForwardPrimer.Name, "#")(0)
                                tColl.Add Split(.ReversePrimer.Name, "#")(0)
                            PCRPrimers.Add tColl
                            
                        End If
                    
                End With
            Next i
          
        
        'export proteins
            Set tColl = New VBA.Collection
            tColl.Add .FinalAssembly.Name
            tColl.Add tResults.Item(j).Item(2)
            Proteins.Add tColl
            
        'export assemblies
            Assemblies.Add tResults.Item(j).Item(4)
        
        End With
               
    Next j
    
    '=====================export stuff to excel sheets
        Application.ScreenUpdating = False
            
            'output each Gibson reaction as a unit
            For i = 1 To ConstructNumber
                Set Gibson = tResults.Item(i).Item(4)
                Call CMC_OutputGibson(Gibson)
            Next i
            
            'output each object type as a group
            Call CMC_OutputPrimers(NewPrim)
            Call CMC_OutputFragments(Frags, PCRPrimers)
            Call CMC_OutputProteins(Proteins)
            Call CMC_OutputAssemblies(Assemblies)
        
        Application.ScreenUpdating = True
    
    
    Set CloningMakeConstructs = tResults
        
    'cleanup
        Set Gibson = Nothing
        Set tResults = Nothing
        Set tTruncations = Nothing
        Set tColl = Nothing
        Set Ligation = Nothing
        Set GI = Nothing
        Set InsertSourceDNA = Nothing
        
        
        Set tOutput = Nothing
        Set PrimColl = Nothing
        Set NewPrim = Nothing
        Set Frags = Nothing
        Set PCRPrimers = Nothing
        Set Proteins = Nothing
    
    End Function


Private Function CMC_tempCheckPrimer1(ByVal DNA As clsDNA, ByVal PrimColl As clsDNAs) As Long
    
    Dim i As Long
    
    If Not PrimColl Is Nothing Then
    
        For i = 1 To PrimColl.Count
            If DNA.Sequence = PrimColl.DNA(i).Sequence Then
                CMC_tempCheckPrimer1 = i
                Exit For
            End If
        Next i
        
    End If

End Function

Private Function CMC_tempCheckPrimer(ByVal DNA As clsDNA) As String
    
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
            CMC_tempCheckPrimer = PrimersArray(i, 1) & "#" & PrimersArray(i, 2)
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
Public Function SmoothData(ByVal dataArray As Variant, ByVal WindowSize As Long) As Variant

    Dim DataLength As Long
    Dim TempOutput() As Variant
    Dim tempIndex As Long
    Dim i As Long, j As Long
    Dim tempsum As Double
        
        
    DataLength = 1 + UBound(dataArray) - LBound(dataArray)
    
    ReDim TempOutput(1 To DataLength)
        
    
            
    For i = 1 To DataLength - WindowSize
        
        tempIndex = i + WindowSize \ 2
        tempsum = 0
        
        For j = i To i + WindowSize - 1
            tempsum = tempsum + dataArray(j)
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
Public Function GetMaxLetterCount(ByVal InputString As String) As Long

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




