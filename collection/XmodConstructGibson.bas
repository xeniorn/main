Attribute VB_Name = "XmodConstructGibson"
'Juraj Ahel, 2015-02-01
'Last Update 2016-02-17

Const DebugMode As Boolean = True

Const jaQuote As String = """"

'These need to be [Private] Dim and a final Wrapper function should ask
Const PythonScriptPath As String = "C:\Users\juraj.ahel\Documents\GitHub\main\python\Gibson overlap by Florian Weissman\JA_GibsonOverlapScript_v160622.py"
Const Python27ProgramName As String = "python.exe"
Const PathToPython27 As String = "C:\Python27"
Const ExcelExportFolder As String = "C:\ExcelExports\GibsonMacro"
Const RNAFoldPath As String = "C:\Program Files (x86)\ViennaRNA Package\RNAfold.exe"



'****************************************************************************************************
Private Function DNAGibsonConstructOverlaps( _
    ByVal Sequence1 As String, _
    ByVal Sequence2 As String, _
    ByVal Sequence3 As String, _
    ByVal ForbiddenRegions As String, _
    ByVal TargetLength As Long, _
    ByVal TargetTm As Double, _
    ByVal MaxDistanceFromCenter As Long, _
    ByVal MaxLength As Long _
    ) As VBA.Collection
    
'====================================================================================================
'Constructs the collection of overlaps to be used for Gibson
'Juraj Ahel, 2016-12-20, for Gibson assembly

'====================================================================================================

    Dim i As Long
    Dim MiddleIndex As Long
    
    Dim Len1 As Long, Len2 As Long, Len3 As Long
    Dim SearchStart As Long, SearchEnd As Long
    
    Dim CurrentLength As Long
    Dim CurrentStart As Long
    Dim CurrentEnd As Long
        
    Dim OverlapAllowed As Boolean
    Dim tOverlap As String
    Dim FinalSequence As String
    
    Dim PutativeOverlaps As VBA.Collection
    Dim tColl As VBA.Collection
    
    Len1 = Len(Sequence1)
    Len2 = Len(Sequence2)
    Len3 = Len(Sequence3)
    
    FinalSequence = Sequence1 & Sequence2 & Sequence3
    
    FinalLength = Len1 + Len2 + Len3
    
    MiddleIndex = 1 + Len1 + Len2 \ 2
    
    Set PutativeOverlaps = New VBA.Collection
    'Sequence, start index'
    
    FinalSequence = Sequence1 & Sequence2 & Sequence3
    
    SearchStart = MiddleIndex - MaxDistanceFromCenter
    If SearchStart < 1 Then SearchStart = 1
    
    'test all allowed lengths
    For CurrentLength = TargetLength To MaxLength
        
        
        SearchEnd = MiddleIndex + MaxDistanceFromCenter - CurrentLength
        If SearchEnd > FinalLength - CurrentLength Then SearchEnd = FinalLength - CurrentLength
        
        'all allowed positions for a given length
        For CurrentStart = SearchStart To SearchEnd
            CurrentEnd = CurrentStart + CurrentLength - 1
        
            OverlapAllowed = True
            
            'if it isn't in any of the forbidden regions
                If OverlapAllowed Then
                    If InStr(1, ForbiddenRegions, "1") <> 0 Then
                        If RangeOverlaps(CurrentStart, CurrentEnd, 1, Len1) Then
                            OverlapAllowed = False
                        End If
                    End If
                End If
                
                If OverlapAllowed Then
                    If InStr(1, ForbiddenRegions, "2") <> 0 Then
                        If RangeOverlaps(CurrentStart, CurrentEnd, Len1 + 1, Len1 + Len2) Then
                            OverlapAllowed = False
                        End If
                    End If
                End If
                
                If OverlapAllowed Then
                    If InStr(1, ForbiddenRegions, "3") <> 0 Then
                        If RangeOverlaps(CurrentStart, CurrentEnd, Len1 + Len2 + 1, Len1 + Len2 + Len3) Then
                            OverlapAllowed = False
                        End If
                    End If
                End If
            
            
            If OverlapAllowed Then
                tOverlap = SubSequenceSelect(FinalSequence, CurrentStart, CurrentEnd)
                ' if the Tm is allowed
                If OligoTm(tOverlap) >= TargetTm Then
                    ' and it doesn't have multiple annealing sites
                    If StringCharCount_IncludeOverlap(FinalSequence, tOverlap) > 1 Then
                        'Debug.Print ("Multiple annealing sites for " & tOverlap)
                    Else
                        Set tColl = New VBA.Collection
                        tColl.Add tOverlap
                        tColl.Add CurrentStart
                        PutativeOverlaps.Add tColl
                    End If
                End If
            End If
            
        Next CurrentStart
        
    Next CurrentLength
    
    Set DNAGibsonConstructOverlaps = PutativeOverlaps
    
    Set tColl = Nothing
    Set PutativeOverlaps = Nothing


End Function
    
    
Sub GibsonOptimizeFrag1()

    Dim InputR As Excel.Range
    Dim OutputR As Excel.Range
    Dim Inputs() As Variant
    Dim Outputs() As Variant
    
    Set InputR = Selection
    Set OutputR = InputR.Offset(InputR.Rows.Count + 2).Resize(5, InputR.Columns.Count)
    
    Dim Results As VBA.Collection
    
    Inputs = InputR.Value
    Outputs = OutputR.Value
    
    Dim pre As Long
    Dim nex As Long
    
    Dim i As Long
    Dim j As Long
    
    Dim obj As clsGibsonSingleLigation
    
    
    OutputR.ClearContents
    
    For j = 1 To OutputR.Columns.Count
        pre = j
        nex = j + 1
        If nex > OutputR.Columns.Count Then nex = 1
    
        
        Set obj = New clsGibsonSingleLigation
        
        With obj
            .FirstSequence = Inputs(3, pre)
            .InsertedSequence = Inputs(4, pre) & Inputs(2, nex)
            .SecondSequence = Inputs(3, nex)
            .ForbiddenRegions = Inputs(5, pre)
        End With
                
        
            Outputs(1, j) = obj.Overlap
            Outputs(2, j) = Len(obj.Overlap)
            Outputs(3, j) = obj.dG
            Outputs(4, j) = obj.Tm
        
    Next j
    
    OutputR.Value = Outputs
    
    


End Sub
    
Sub GibsonOptimizeFrag()

    Dim InputR As Excel.Range
    Dim OutputR As Excel.Range
    Dim Inputs() As Variant
    Dim Outputs() As Variant
    
    Set InputR = Selection
    Set OutputR = InputR.Offset(InputR.Rows.Count + 2).Resize(5, InputR.Columns.Count)
    
    Dim Results As VBA.Collection
    
    Inputs = InputR.Value
    Outputs = OutputR.Value
    
    Dim pre As Long
    Dim nex As Long
    
    Dim i As Long
    Dim j As Long
    
    
    OutputR.ClearContents
    
    For j = 1 To OutputR.Columns.Count
        pre = j
        nex = j + 1
        If nex > OutputR.Columns.Count Then nex = 1
    
    
        Set Results = DNAGibsonOptimizeOverlaps(Inputs(3, pre), Inputs(4, pre) & Inputs(2, nex), Inputs(3, nex), Inputs(5, pre))
        
        
        
        
        
        For i = 1 To Results.Count
            Outputs(i, j) = Results.Item(i)
        Next i
    Next j
    
    OutputR.Value = Outputs
    
    


End Sub

'****************************************************************************************************
Public Function DNAGibsonOptimizeOverlaps( _
    ByVal Sequence1 As String, _
    ByVal Sequence2 As String, _
    ByVal Sequence3 As String, _
    Optional ByVal ForbiddenRegions As String = "", _
    Optional ByVal TargetLength As Long = 15, _
    Optional ByVal TargetTm As Double = 48, _
    Optional ByVal MaxDistanceFromCenter As Long = 50, _
    Optional ByVal MaxLength As Long = 25 _
    ) As VBA.Collection
    
'====================================================================================================
'Calculates the optimal Gibson overlap for annealing sequences 1 and 3, with inserting sequence 2 in between
'Juraj Ahel, 2016-12-20, for Gibson assembly

'output:
'Collection of
    'OVERLAP as String      'overlap sequence
    'DG as Double           'Gibbs energy for SS formation of overlap
    'TM as Double           'melting temperature of overlap
    '1 as String            'first fragment to be annealed
    '2 as String            'second fragment to be annealed

'====================================================================================================
    
    Dim i As Long

    Dim SequencesFound As Long
    Dim tOverlap As String
    
    Dim PutativeOverlaps As VBA.Collection
    
    Dim RNAFoldPath As String
    Dim RNAFoldCommand As String
    Dim TempInput As String
    Dim TempOutput As String
    Dim TempFilenameBase As String
    
        
    ' Construct collection of overlaps
        Set PutativeOverlaps = DNAGibsonConstructOverlaps( _
            Sequence1:=Sequence1, _
            Sequence2:=Sequence2, _
            Sequence3:=Sequence3, _
            ForbiddenRegions:=ForbiddenRegions, _
            TargetLength:=TargetLength, _
            TargetTm:=TargetTm, _
            MaxDistanceFromCenter:=MaxDistanceFromCenter, _
            MaxLength:=MaxLength _
            )
        
        
    TempFilenameBase = FileSystem_GetTempFolder(True) & "JA_Gibson_" & TempTimeStampName
    TempInput = TempFilenameBase & "_in"
    TempOutput = TempFilenameBase & "_out"
    Call CreateEmptyFile(TempInput)
    'Call CreateEmptyFile(TempOutput)
    
    'calculate secondary structure energies (ViennaRNA/RNAfold)
        
        For i = 1 To PutativeOverlaps.Count
            'Debug.Print (PutativeOverlaps.Item(i).Item(1))
            Call WriteTextFile(PutativeOverlaps.Item(i).Item(1), TempInput, True)
        Next i
        
        RNAFoldPath = "C:\ViennaRNA\"
        RNAFoldCommand = "RNAfold.exe"
        
        Call CallProgram(RNAFoldCommand, RNAFoldPath, "--noGU --noPS -T 10 --infile=" & TempInput & " --outfile=" & TempOutput)
         
        TempOutput = TempOutput & ".fold" 'ViennaRNA 2.3 does it'
    
    Dim OutputLines() As String
    Dim tempdGString As String
    Dim tempdG As Double
    Dim RegEx As New RegExp
    Dim SortedCollection As VBA.Collection
    Dim tColl As VBA.Collection
    Dim k As Long
    
    'extract results
        With RegEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = ".*\( *(-?[0-9]+\.[0-9]+)\)"
        End With
                
        OutputLines = Split(ReadTextFile(TempOutput), vbCrLf)
        Set SortedCollection = New VBA.Collection
        
    'sort results into a collection
        For i = 1 To PutativeOverlaps.Count
            tempdG = val(RegEx.Replace(OutputLines(2 * i - 1), "$1"))
            Set tColl = PutativeOverlaps.Item(i)
            tColl.Add tempdG
                           
            If i = 1 Then
                SortedCollection.Add tColl
            Else
                j = 1
                Do While j <= SortedCollection.Count
                    If tempdG < SortedCollection.Item(j).Item(3) Then
                        j = j + 1
                    Else
                        Exit Do
                    End If
                Loop
                
                If j > SortedCollection.Count Then
                    SortedCollection.Add tColl
                Else
                    SortedCollection.Add tColl, Before:=j
                End If
                
            End If
            
        Next i
        
    Set DNAGibsonOptimizeOverlaps = DNAGibsonExtractBestOverlaps(SortedCollection, Sequence1 & Sequence2 & Sequence3, TargetLength, TargetTm)
    
    For i = 1 To DNAGibsonOptimizeOverlaps.Count
        Debug.Print (DNAGibsonOptimizeOverlaps.Item(i))
    Next i
    
    Set SortedCollection = Nothing
    Set PutativeOverlaps = Nothing
    Set tColl = Nothing
    

End Function

Private Function DNAGibsonExtractBestOverlaps( _
    ByRef SortedCollection As VBA.Collection, _
    ByVal FinalSequence As String, _
    ByVal TargetLength As Long, _
    ByVal TargetTm As Double _
    ) As VBA.Collection
    
    Dim tColl As VBA.Collection
    Dim Fragment1 As String
    Dim Fragment2 As String
    Dim OverlapLength As Long
    Dim MinLength As Long
    Dim MindG As Double
    
    Dim i As Long
    Dim ChosenIndex As Long
    
    Dim SortColl2 As VBA.Collection
    
    Set tColl = SortedCollection.Item(1)
    Set SortColl2 = tColl
    
    MindG = tColl.Item(3)
    MinLength = Len(tColl.Item(1))
    ChosenIndex = 1
    
    'grab the shortest overlap with minimal dG
    For i = 2 To SortedCollection.Count
        Set tColl = SortedCollection.Item(i)
        If tColl.Item(3) >= MindG Then
            If Len(tColl.Item(1)) < MinLength Then
                Set SortColl2 = tColl
                MinLength = Len(SortColl2.Item(1))
                ChosenIndex = i
            End If
        End If
    Next i
    
    
    
    OverlapLength = MinLength
    Fragment1 = SubSequenceSelect(FinalSequence, 1, SortColl2.Item(2) + OverlapLength - 1)
    Fragment2 = SubSequenceSelect(FinalSequence, SortColl2.Item(2), Len(FinalSequence))
    
    Set tColl = New VBA.Collection
    
    With tColl
        .Add SortColl2.Item(1), "OVERLAP"
        .Add SortColl2.Item(3), "DG"
        .Add OligoTm(SortColl2.Item(1)), "TM"
        '.Add Fragment1, "1"
        '.Add Fragment2, "2"
    End With
    
    Set DNAGibsonExtractBestOverlaps = tColl
    
    Set tColl = Nothing
            
                
End Function



Sub Test()

Dim a As String
Dim b As String
Dim c As String

a = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
b = "GGGGGGG"
c = "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"

Dim i As Long
Dim x As VBA.Collection

Set x = DNAGibsonOptimizeOverlaps(a, b, c, "")

For i = 1 To DNAGIBs
Debug.Print (x.Item(i))
Next i

Set x = Nothing

End Sub


'****************************************************************************************************
Sub GibsonTest()

'====================================================================================================
'A huge procedure that generates the final result of Gibson overlap analysis by Florian's script
'It takes a range with prepared inputs, and directly outputs the results to 9 cells to the right
'This one should be made more modular / cleaned up, when I get the time
'Juraj Ahel, 2015-02-11, for Gibson assembly
'Last update 2016-02-17
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
    RunDir = ExcelExportFolder
    
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
            CallPythonScript FullPythonInputFilename, RunDir, FullPythonOutputFilename
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
                            OverlapEnergy = val(ExtractParameter(textline, "dG", "[]"))
                            OverlapTm = val(ExtractParameter(textline, "Tm", "[]"))
                        
                        Case "PRIMER1"
                            PrimerName(1) = ExtractParameter(textline, "PrimerName", "[]")
                            PrimerSequence(1) = ExtractParameter(textline, "Sequence", "[]")
                            PrimerTmNN(1) = val(ExtractParameter(textline, "Tm", "[]"))
                        Case "PRIMER2"
                            PrimerName(2) = ExtractParameter(textline, "PrimerName", "[]")
                            PrimerSequence(2) = ExtractParameter(textline, "Sequence", "[]")
                            PrimerTmNN(2) = val(ExtractParameter(textline, "Tm", "[]"))
                            
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
            Deleting = Not DebugMode
            
            If Deleting Then
                
                ExistenceTest = Dir(FullPythonInputFilename)
                If ExistenceTest <> "" Then Kill (FullPythonInputFilename)
                ExistenceTest = Dir(FullPythonOutputFilename)
                If ExistenceTest <> "" Then Kill (FullPythonOutputFilename)
                
            End If
                    
    
        End If '#############################################################################################################1
    Next cell
    
    MsgBox ("Done! Runtime: " & Round((Timer - StartTime), 2) & " seconds")
    
999     'Goto

End Sub

'****************************************************************************************************
Sub GibsonMother()
'====================================================================================================
'calls GibsonMonster for all areas in a selection
'Juraj Ahel, 2016-07-05
'
'====================================================================================================

    Dim inputRange As Range
    Dim SubRange As Range

    Set inputRange = Selection

    If inputRange.Areas.Count >= 1 Then
    
        For Each SubRange In inputRange.Areas
        
            SubRange.Select
            
            Call GibsonMonster
            
        Next SubRange
        
    End If

    Set SubRange = Nothing
    Set inputRange = Nothing

End Sub

'****************************************************************************************************
Sub GibsonMonster()
Attribute GibsonMonster.VB_ProcData.VB_Invoke_Func = "G\n14"

'====================================================================================================
'
'Juraj Ahel, 2016-06-27
'v0.5
'====================================================================================================
'2016-07-04 add PCR product length, make the program work for >2 sequences
'2016-07-05 add automatic FASTA formatted output of primers / construct
    
    'constants
    Const Table1Size As Long = 13
    
    Const Table2Size As Long = 12
    Const AssemblySize As Long = 5
    Const ORFTableSize As Long = 3
    
    Const GapSize As Long = 2
    
    Const conORFDetectNumber As Long = 7
    
    Const RequiredColumns As Long = 1
    Const ParameterNumber As Long = 9
    
    'positions of parameters in table1
    Const pName As Long = 1
    Const pSeq As Long = 7
    Const pFlorian As Long = 13
    
    'iterators
    Dim i As Long
    Dim j As Long
    Dim CurrColumn As Long
    Dim NextColumn As Long
    
    'tables
    Dim InputTable As Range
    Dim InputTableValues() As Variant
    
    Dim PrimersTable As Range
    Dim PrimersTableValues() As Variant
    
    Dim AssemblyTable As Range
    Dim AssemblyTableValues() As Variant
            
    Dim ORFTable As Range
    Dim ORFTableValues() As Variant
    
    Dim OutputTable As Range
    Dim OutputTableValues() As Variant
    
    'descriptors
    Dim FragmentCount As Long
    Dim ORFDetectNumber As Long
    
    'temp vars
    Dim PrevFragment As String
    Dim CurrFragment As String
    
    Dim CurrName As String
    Dim NextName As String
    
    Dim Addition As String
    Dim AllowedOverlap As String
    
    Dim FlorianParameter As String
    Dim GibsonScriptInput As String
    
    Dim GibsonResults() As Variant
    Dim ExpectedAssembly As String
    
    Dim ORFs As VBA.Collection
    
    'headers
    
    Dim RowHeaders1(1 To Table1Size, 1 To 1)
    Dim RowHeaders2(1 To Table2Size, 1 To 1)
    Dim RowHeaders3(1 To AssemblySize, 1 To 1)
    Dim RowHeaders4(1 To ORFTableSize, 1 To 1)
    
    RowHeaders1(1, 1) = "name"
    RowHeaders1(2, 1) = "linker/addition before"
    RowHeaders1(3, 1) = "start codon"
    RowHeaders1(4, 1) = "linker"
    RowHeaders1(5, 1) = "tag"
    RowHeaders1(6, 1) = "linker"
    RowHeaders1(7, 1) = "fragment sequence"
    RowHeaders1(8, 1) = "linker"
    RowHeaders1(9, 1) = "tag"
    RowHeaders1(10, 1) = "linker"
    RowHeaders1(11, 1) = "stop codon"
    RowHeaders1(12, 1) = "linker/adition after"
    RowHeaders1(13, 1) = "allowed overlap to next"
    
    RowHeaders2(1, 1) = "source sequence"
    RowHeaders2(2, 1) = "forward primer"
    RowHeaders2(3, 1) = "reverse primer"
    RowHeaders2(4, 1) = "forward name"
    RowHeaders2(5, 1) = "reverse name"
    RowHeaders2(6, 1) = "forward Tm"
    RowHeaders2(7, 1) = "reverse Tm"
    RowHeaders2(8, 1) = "forward Length"
    RowHeaders2(9, 1) = "reverse Length"
    RowHeaders2(10, 1) = "overlap previous"
    RowHeaders2(11, 1) = "overlap next"
    RowHeaders2(12, 1) = "dG next overlap"
    
    RowHeaders3(1, 1) = "PCR sequence"
    RowHeaders3(2, 1) = "length"
    RowHeaders3(3, 1) = "tags"
    RowHeaders3(4, 1) = "assembly"
    RowHeaders3(5, 1) = "comment"
    
    RowHeaders4(1, 1) = "nucleotides"
    RowHeaders4(2, 1) = "translation"
    RowHeaders4(3, 1) = "protein length"
    
    ORFDetectNumber = conORFDetectNumber
    
    
    ':::START:::
    
    Set InputTable = Selection
        
    'Parse Inputs
        With InputTable
            
            If .Rows.Count < Table1Size Then
                Err.Raise 1, , "Not enough rows"
            End If
            
            If .Columns.Count < RequiredColumns Then
                Err.Raise 1, , "Not enough columns"
            End If
            
            FragmentCount = .Columns.Count
            
            'if too much was selected by accident
            Set InputTable = InputTable.Resize(Table1Size, FragmentCount)
            
            InputTableValues = .Value
            
            'remove spaces and all other non-DNA characters that might be present from
            'relevant inputs
            For j = 1 To FragmentCount
                For i = 2 To Table1Size - 1
                    InputTableValues(i, j) = DNAParseTextInput(InputTableValues(i, j))
                    ExpectedAssembly = ExpectedAssembly & InputTableValues(i, j)
                Next i
            Next j
            
        End With
    
    'initialize Tables
        Set PrimersTable = InputTable.Offset(InputTable.Rows.Count + GapSize, 0).Resize(Table2Size, FragmentCount)
        PrimersTableValues = PrimersTable.Value
        
        Set AssemblyTable = PrimersTable.Offset(PrimersTable.Rows.Count + GapSize, 0).Resize(AssemblySize, FragmentCount)
        AssemblyTableValues = AssemblyTable.Value
        
        Set ORFTable = AssemblyTable.Offset(AssemblyTable.Rows.Count + GapSize, 0).Resize(ORFTableSize, ORFDetectNumber)
        ORFTableValues = ORFTable.Value
        
        Set OutputTable = ORFTable.Offset(ORFTable.Rows.Count + GapSize, 0).Resize(2 * FragmentCount + 1, 3)
        OutputTableValues = OutputTable.Value
    
        
    ':::::::::::::::::::::::::::::::::main:::::::::::::::::::::::::::::::::::
    CurrColumn = 1
    Do While CurrColumn <= FragmentCount
        
        'allow circularity
            Select Case CurrColumn
                Case FragmentCount
                    NextColumn = 1
                Case 1
                    NextColumn = CurrColumn + 1
                Case Else
                    NextColumn = CurrColumn + 1
            End Select
            
        'define the current fragment
            CurrFragment = InputTableValues(pSeq, CurrColumn)
            CurrName = InputTableValues(pName, CurrColumn)
                    
        'DNA to be added in between + name
        'DNA that will be added are the C-terminal (3') additions to current fragment
        'and the N-terminal (5') additions to the next fragment
            Addition = ""
            For i = 8 To 12
                Addition = Addition & InputTableValues(i, CurrColumn)
            Next i
            For i = 2 To 6
                Addition = Addition & InputTableValues(i, NextColumn)
            Next i
                
        'downstream region (next fragment)
            NextFragment = InputTableValues(pSeq, NextColumn)
            NextName = InputTableValues(pName, NextColumn)
            
        'overlaps parameter
            AllowedOverlap = InputTableValues(pFlorian, CurrColumn)
            Select Case True
                Case (AllowedOverlap Like "1")
                    FlorianParameter = "23"
                Case (AllowedOverlap Like "2")
                    FlorianParameter = "13"
                Case (AllowedOverlap Like "3")
                    FlorianParameter = "12"
                Case (AllowedOverlap Like "[12][12]")
                    FlorianParameter = "3"
                Case (AllowedOverlap Like "[13][13]")
                    FlorianParameter = "2"
                Case (AllowedOverlap Like "[23][23]")
                    FlorianParameter = "1"
                Case Else 'includes [123]*
                    FlorianParameter = ""
            End Select
              
              
        'construct input string for Gibson macro, run the macro
            GibsonScriptInput = Join(Array( _
                CurrFragment, _
                Addition, _
                NextFragment, _
                FlorianParameter, _
                CurrName, _
                NextName _
                ), _
                vbCrLf)
                
            ReDim GibsonResults(1 To ParameterNumber)
            Call GibsonRun(GibsonScriptInput, GibsonResults)
            
        'extract results
        'as of 20160627, results are Array(1..9)
        'results are in format: Overlap(Sequence|deltaG|Tm)||PrimerNext(Name|Seq|Tm)||PrimerCurr(%|%|%)
            
            Dim OverlapSequence As String
            Dim SSdG As Double
            Dim OverlapTm As Double
            
            Dim PrimerCurrName As String
            Dim PrimerCurrSeq As String
            Dim PrimerCurrTm As Double
            
            Dim PrimerNextName As String
            Dim PrimerNextSeq As String
            Dim PrimerNextTm As Double
            
            OverlapSequence = CStr(GibsonResults(1))
            SSdG = CDbl(GibsonResults(2))
            OverlapTm = CDbl(GibsonResults(3))
            
            PrimerNextName = CStr(GibsonResults(4))
            PrimerNextSeq = CStr(GibsonResults(5))
            PrimerNextTm = CDbl(GibsonResults(6))
            
            PrimerCurrName = CStr(GibsonResults(7))
            PrimerCurrSeq = CStr(GibsonResults(8))
            PrimerCurrTm = CDbl(GibsonResults(9))
            
        'output results to table
        
            PrimersTableValues(1, CurrColumn) = CurrFragment
            
            PrimersTableValues(3, CurrColumn) = PrimerCurrSeq
            PrimersTableValues(2, NextColumn) = PrimerNextSeq
            
            PrimersTableValues(5, CurrColumn) = PrimerCurrName
            PrimersTableValues(4, NextColumn) = PrimerNextName
            
            PrimersTableValues(7, CurrColumn) = PrimerCurrTm
            PrimersTableValues(6, NextColumn) = PrimerNextTm
            
            PrimersTableValues(9, CurrColumn) = Len(PrimerCurrSeq)
            PrimersTableValues(8, NextColumn) = Len(PrimerNextSeq)
            
            PrimersTableValues(11, CurrColumn) = OverlapSequence
            PrimersTableValues(10, NextColumn) = OverlapSequence
            
            PrimersTableValues(12, CurrColumn) = SSdG
                    
        CurrColumn = CurrColumn + 1
        
    Loop
    
    'when all the values have been extracted, do the following:
    
        Dim tTemplate As String
        Dim tFor As String
        Dim tRev As String
        
        Dim tNterm As String
        Dim tCterm As String
        
        Dim tResult As String
    
        For i = 1 To FragmentCount
        
            'perform in-silico PCR of all fragments
                
                tTemplate = PrimersTableValues(1, i)
                tFor = PrimersTableValues(2, i)
                tRev = PrimersTableValues(3, i)
                
                AssemblyTableValues(1, i) = PCRWithOverhangs(tTemplate, tFor, tRev, True)
                AssemblyTableValues(2, i) = Len(AssemblyTableValues(1, i))
            
            'annotate the tags/linkers
            
                tResult = ""
                tNterm = ""
                tCterm = ""
                
                For j = 3 To 6
                    tNterm = tNterm & InputTableValues(j, i)
                Next j
                
                For j = 8 To 11
                    tCterm = tCterm & InputTableValues(j, i)
                Next j
                
                'check which extension exists, apply it
                'check if any exist
                If Len(tCterm) > 0 Or Len(tNterm) > 0 Then
                
                    If Len(tNterm) > 0 Then
                        tNterm = "N-" & DNATranslate(tNterm)
                    End If
                    
                    If Len(tCterm) > 0 Then
                        tCterm = "C-" & DNATranslate(tCterm)
                    End If
                    
                    'if both
                    If Len(tNterm) > 0 And Len(tCterm) > 0 Then
                        tResult = tNterm & " / " & tCterm
                        
                    'if only one
                    Else
                        If Len(tNterm) > 0 Then
                            tResult = tNterm
                        Else
                            tResult = tCterm
                        End If
                    End If
                    
                End If
                                    
                AssemblyTableValues(3, i) = tResult
            
        Next i
        
    'ligate the fragments
        
        tResult = AssemblyTableValues(1, 1)
        
        'For i = 2 To FragmentCount
        '    tResult = DNAGibsonLigation(tResult, AssemblyTableValues(1, i))
        'Next i
        
        Dim tArray()
        ReDim tArray(1 To FragmentCount)
        
        For i = 1 To FragmentCount
            tArray(i) = AssemblyTableValues(1, i)
        Next i
        
        tResult = DNAGibsonLigation(tArray)
        
        AssemblyTableValues(4, 1) = tResult
        
        'check whether the in silico PCR gives the correct sequence
        If DNAEqual(AssemblyTableValues(4, 1), ExpectedAssembly) Then
            AssemblyTableValues(5, 1) = "Correct!"
        Else
            AssemblyTableValues(5, 1) = "Error in assembly!"
        End If
        
    
    'check the ORFs
        tResult = ""
        tTemplate = AssemblyTableValues(4, 1)
            
        Set ORFs = DNAFindORFs( _
            Sequence:=tTemplate, _
            Circular:=True, _
            MinimumORFLength:=50, _
            AllowORFOverlap:=False, _
            AllowReverseStrand:=True)
            
        If ORFDetectNumber > ORFs.Count Then
            ORFDetectNumber = ORFs.Count
        End If
            
        For i = 1 To ORFDetectNumber
            
            tResult = ORFs.Item(i)
                
            ORFTableValues(1, i) = tResult
            ORFTableValues(2, i) = DNATranslate(tResult)
            ORFTableValues(3, i) = Len(Replace(ORFTableValues(2, i), "*", ""))
            
        Next i
        
        
        OutputTableValues(1, 1) = "Assembly" & InputTable.Column
        OutputTableValues(1, 2) = AssemblyTableValues(4, 1)
        
        For i = 1 To FragmentCount
        
            OutputTableValues(2 * i, 1) = PrimersTableValues(4, i)
            OutputTableValues(2 * i, 2) = PrimersTableValues(2, i)
            
            OutputTableValues(2 * i + 1, 1) = PrimersTableValues(5, i)
            OutputTableValues(2 * i + 1, 2) = PrimersTableValues(3, i)
            
        Next i
        
        For i = LBound(OutputTableValues, 1) To UBound(OutputTableValues, 1)
          
            OutputTableValues(i, 3) = ">" & OutputTableValues(i, 1) & "###" & OutputTableValues(i, 2)
          
        Next i
                
    'repair table headers
                    
        With InputTable
            .Offset(-1, -1).Resize(1, 1).Value = "Inputs"
            .Offset(0, -1).Resize(Table1Size, 1).Value = RowHeaders1
        End With
                    
        With PrimersTable
            .Offset(-1, -1).Resize(1, 1).Value = "Primers"
            .Offset(0, -1).Resize(Table2Size, 1).Value = RowHeaders2
            .Value = PrimersTableValues
        End With
        
        With AssemblyTable
            .Offset(-1, -1).Resize(1, 1).Value = "PCR"
            .Offset(0, -1).Resize(AssemblySize, 1).Value = RowHeaders3
            .Value = AssemblyTableValues
        End With
        
        With ORFTable
            .Offset(-1, -1).Resize(1, 1).Value = "ORFs"
            .Offset(0, -1).Resize(ORFTableSize, 1).Value = RowHeaders4
            .Value = ORFTableValues
        End With
        
        With OutputTable
            .Offset(-1, -1).Resize(1, 1).Value = "Output"
            '.Offset(0, -1).Resize(ORFTableSize, 1).Value = RowHeaders4
            .Value = OutputTableValues
        End With
    
    
    'clean up
        
        Set InputTable = Nothing
        Set PrimersTable = Nothing
        Set AssemblyTable = Nothing
        Set ORFTable = Nothing
        Set ORFs = Nothing
        Set OutputTable = Nothing

End Sub

'****************************************************************************************************
Sub GibsonMacro()

'====================================================================================================
'wrapper for Gibson assembly
'Juraj Ahel, 2015-02-11, for Gibson assembly
'2016-06-27 separate to multiple Subs
'====================================================================================================
    
    Const ParameterNumber As Long = 9
    
    'Just for printing running time in the end
    Dim StartTime As Double
           
    Dim myRange As Range, cell As Range
    Dim tempResults() As Variant
    Dim TempOutput() As Variant
    
    Dim AssemblyCount As Long
    
    Dim i As Long
    Dim j As Long
        
    StartTime = Timer
    
    'Input
    Set myRange = Selection
    If Selection.Columns.Count > 1 Then
        MsgBox ("No. NO. NO. Just one Column! And to the right it must be free!")
        GoTo 999
    End If
    
    AssemblyCount = myRange.Cells.Count
    
    ReDim TempOutput(1 To AssemblyCount, 1 To ParameterNumber)
    ReDim tempResults(1 To ParameterNumber)
    
    j = 0
    For Each cell In myRange
    
        j = j + 1
        
        If cell.Value <> "" And cell.Value <> 0 Then
        
            Call GibsonRun(cell.Value, tempResults)
            
            For i = 1 To ParameterNumber
                TempOutput(j, i) = tempResults(i)
            Next i
            
        End If
        
    Next cell
    
    Set myRange = myRange.Offset(0, 1).Resize(AssemblyCount, ParameterNumber)
    
    myRange.Value = TempOutput
    
    Set myRange = Nothing
    

999     'Goto

    MsgBox ("Done! Runtime: " & Round((Timer - StartTime), 2) & " seconds")

End Sub


'****************************************************************************************************
Sub GibsonRun( _
    ByVal InputString As String, _
    ByRef ResultsArray() As Variant)

'====================================================================================================
'A huge procedure that generates the final result of Gibson overlap analysis by Florian's script
'It takes a range with prepared inputs, and directly outputs the results to 9 cells to the right
'This one should be made more modular / cleaned up, when I get the time
'Juraj Ahel, 2015-02-11, for Gibson assembly
'2016-06-27 separate to multiple Subs
'====================================================================================================
    
    'Constants
    Const tempExtension As String = ".jatmp"
    
    'Output variables
    Dim OverlapSequence As String
    Dim PrimerSequence(1 To 2) As String
    Dim PrimerName(1 To 2) As String
    Dim OverlapEnergy As Double
    Dim OverlapTm As Double
    Dim PrimerTmNN(1 To 2) As Double
            
    'Input and internal variables
    Dim FullPythonInputFilename As String, FullPythonOutputFilename As String
    Dim RunDir As String
    Dim textline As String
    Dim TempName As String
    Dim Sep As String
    
    Sep = Application.PathSeparator
    
    'Where the program will run and do its internal stuff, like leaving temp files
    RunDir = FileSystem_GetTempFolder
    
    'Temporary name to be used for temp files
    TempName = TempTimeStampName ' & "_R" & cell.Row & "_C" & cell.Column
       
    'Temporary name for the python script input and output
    FullPythonInputFilename = RunDir & Sep & TempName & tempExtension
    FullPythonOutputFilename = RunDir & Sep & TempName & "_out" & tempExtension
    
    'The python script needs an external file, at least I don't know how to pipe it in directly well
    Call WriteTextFile(InputString, FullPythonInputFilename)
    
    'This actually gives the result. The Python path is defined in the Subprocedure
    CallPythonScript FullPythonInputFilename, RunDir, FullPythonOutputFilename
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
                    OverlapEnergy = val(ExtractParameter(textline, "dG", "[]"))
                    OverlapTm = val(ExtractParameter(textline, "Tm", "[]"))
                
                Case "PRIMER1"
                    PrimerName(1) = ExtractParameter(textline, "PrimerName", "[]")
                    PrimerSequence(1) = ExtractParameter(textline, "Sequence", "[]")
                    PrimerTmNN(1) = val(ExtractParameter(textline, "Tm", "[]"))
                Case "PRIMER2"
                    PrimerName(2) = ExtractParameter(textline, "PrimerName", "[]")
                    PrimerSequence(2) = ExtractParameter(textline, "Sequence", "[]")
                    PrimerTmNN(2) = val(ExtractParameter(textline, "Tm", "[]"))
                    
            End Select
               
            
        End If
            
    Loop
    
    Close #1
    
    ResultsArray(1) = OverlapSequence
    ResultsArray(2) = OverlapEnergy
    ResultsArray(3) = OverlapTm
    ResultsArray(4) = PrimerName(1)
    ResultsArray(5) = PrimerSequence(1)
    ResultsArray(6) = PrimerTmNN(1)
    ResultsArray(7) = PrimerName(2)
    ResultsArray(8) = PrimerSequence(2)
    ResultsArray(9) = PrimerTmNN(2)
    
    'If Deleting is true, temp files are deleted. I might add an inputbox to choose whether to do it
    Dim Deleting As Boolean, ExistenceTest As String
    Deleting = Not DebugMode
    
    If Deleting Then
        
        ExistenceTest = Dir(FullPythonInputFilename)
        If ExistenceTest <> "" Then Kill (FullPythonInputFilename)
        ExistenceTest = Dir(FullPythonOutputFilename)
        If ExistenceTest <> "" Then Kill (FullPythonOutputFilename)
        
    End If

End Sub

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
Dim StartIndex As Long, EndIndex As Long

Dim Locs As Long, Loce As Long
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


Sub CallPythonScript(InputFile As String, RunDir As String, OutputFile As String)

'====================================================================================================
'Wrapper for calling the python script
'Juraj Ahel, 2015-02-11, for Gibson assembly and general purposes
'Last update 2016-02-16
'====================================================================================================
'bound to Module constants!

Dim prog As String, path As String, argum As String

prog = Python27ProgramName
path = PathToPython27

argum = jaQuote & PythonScriptPath & jaQuote & _
        " " & jaQuote & InputFile & jaQuote & _
        " " & jaQuote & RNAFoldPath & jaQuote & _
        " " & jaQuote & ExcelExportFolder & jaQuote

CallProgram ProgramCommand:=prog, _
            ProgramPath:=path, _
            ArgList:=argum, _
            WaitUntilFinished:=True, _
            WindowMode:="HIDE", _
            RunDirectory:=RunDir, _
            RunAsRawCmd:=True, _
            OutputFile:=OutputFile

End Sub


'****************************************************************************************************
Sub ExportToTXT(SourceData As Range, Optional FilePath As String = ExcelExportFolder, Optional FilenameBase As String = "ExcelOutput ", Optional Extension As String = ".txt")

'====================================================================================================
'Exports a separate text file for each cell in selection
'Still needs to be added modular file naming, now it's always "Fragment #.txt"
'
'Juraj Ahel, 2015-02-10, for Gibson assembly and general purposes
'Last update 2016-02-17
'====================================================================================================

Dim OutputFile As String
Dim DataSource As Range

    If Right(FilePath, 1) <> "\" Then FilePath = FilePath & "\"

    OutputFile = FilePath & FilenameBase & Extension
    Call WriteTextFile(SourceData(1, 1).Value, OutputFile)


End Sub



