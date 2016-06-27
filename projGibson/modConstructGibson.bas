Attribute VB_Name = "modConstructGibson"
'Juraj Ahel, 2015-02-01
'Last Update, 2016-02-17

Const DebugMode As Boolean = True

Const jaQuote As String = """"

'These need to be [Private] Dim and a final Wrapper function should ask
Const PythonScriptPath As String = "E:\PhD\Tools Alati Macros Scripts Programs Utilities\Gibson overlap by Florian Weissman\JA_GibsonOverlapScript_v160622.py"
Const Python27ProgramName As String = "python.exe"
Const PathToPython27 As String = "C:\Python27"
Const ExcelExportFolder As String = "C:\ExcelExports\GibsonMacro"
Const RNAFoldPath As String = "C:\Program Files (x86)\ViennaRNA Package\RNAfold.exe"


'****************************************************************************************************
Function TempTimeStampName() As String

'====================================================================================================
'A simple function that generates a timestamp string, containing full date and time without delimiters
'(YYYYMMDDhhmmss format)
'Juraj Ahel, 2015-02-11, for creating (almost certainly) unique files for GibsonTest
'Last update 2015-02-11
'====================================================================================================

    Dim t As String
    
    t = Now
    t = Replace(t, " ", "")
    t = Replace(t, ":", "")
    t = Replace(t, "-", "")
    
    TempTimeStampName = t

End Function

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
    Dim TempName As String, TempExtension As String: TempExtension = ".jatmp"
    
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
            Call ExportToTXT(cell, RunDir, TempName, TempExtension)
            
            'Temporary name for the python script input and output
            FullPythonInputFilename = RunDir & TempName & TempExtension
            FullPythonOutputFilename = RunDir & TempName & "_python" & TempExtension
            
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
Sub GibsonMonster()

'====================================================================================================
'
'CurrColumnuraCurrColumn Ahel, 2016-06-27
'
'====================================================================================================
    
    Const Table1Size As Long = 10
    Const Table2Size As Long = 10
    Const Table3Size As Long = 2
    Const AssemblySize As Long = 1
    
    Const GapSize As Long = 2
    
    Const RequiredRows As Long = 10
    Const RequiredColumns As Long = 2
    Const ParameterNumber As Long = 9
    
    Dim i As Long
    Dim j As Long
    Dim CurrColumn As Long
    Dim PrevColumn As Long
    
    Dim InputRange As Range
    Dim InputData() As Variant
    
    Dim FragmentCount As Long
    
    Dim PrevFragment As String
    Dim NextFragment As String
    
    Dim PrevName As String
    Dim NextName As String
    
    Dim Addition As String
    Dim AllowedOverlap As String
    
    Dim FlorianParameter As String
    Dim GibsonScriptInput As String
    
    Dim GibsonResults() As Variant
    
    
    Set InputRange = Selection
        
    'Parse Inputs
        With InputRange
            
            If .Rows.Count < RequiredRows Then
                Err.Raise 1, , "Not enough rows"
            End If
            
            If .Columns.Count < RequiredColumns Then
                Err.Raise 1, , "Not enough columns"
            End If
            
            FragmentCount = .Columns.Count
            
            InputData = .Value
            
        End With
    
    'initialize Table2
        Dim PrimersTable As Range
        Dim PrimersTableValues() As Variant
        
        Set PrimersTable = InputRange.Offset(Table1Size + GapSize, 0).Resize(Table2Size, FragmentCount)
        'shortcut to initialization of PrimersTable array
            PrimersTableValues = PrimersTable.Value
    
        
    ':::::::::::::::::::::::::::::::::main:::::::::::::::::::::::::::::::::::
    CurrColumn = 1
    Do While CurrColumn < FragmentCount
        
        'upstream DNA region + name
            If CurrColumn = 1 Then
                PrevColumn = FragmentCount
            Else
                PrevColumn = CurrColumn - 1
            End If
            
            PrevFragment = InputData(PrevColumn, 9)
            PrevName = InputData(PrevColumn, 1)
        
        'DNA to be added in between + name
            Addition = ""
            For i = 2 To 8
                Addition = Addition & InputData(CurrColumn, i)
            Next i
                
        'downstream region
            NextFragment = InputData(CurrColumn, 9)
            NextName = InputData(CurrColumn, 1)
            
        'overlaps parameter
            AllowedOverlap = InputData(CurrColumn, 10)
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
                PrevFragment, _
                Addition, _
                NextFragment, _
                FlorianParameter, _
                PrevName, _
                NextName _
                ), _
                vbCrLf)
                
            ReDim GibsonResults(1 To ParameterNumber)
            Call GibsonRun(GibsonScriptInput, GibsonResults)
            
        'extract results
        'as of 20160627, results are Array(1..9)
        'results are in format: Overlap(Sequence|deltaG|Tm)||PrimerNext(Name|Seq|Tm)||PrimerPrev(%|%|%)
            
            Dim OverlapSequence As String
            Dim SSdG As Double
            Dim OverlapTm As Double
            
            Dim PrimerPrevName As String
            Dim PrimerPrevSeq As String
            Dim PrimerPrevTm As Double
            
            Dim PrimerNextName As String
            Dim PrimerNextSeq As String
            Dim PrimerNextTm As Double
            
            OverlapSequence = CStr(GibsonResults(1))
            SSdG = CDbl(GibsonResults(2))
            OverlapTm = CDbl(GibsonResults(3))
            
            PrimerNextName = CStr(GibsonResults(4))
            PrimerNextSeq = CStr(GibsonResults(5))
            PrimerNextTm = CDbl(GibsonResults(6))
            
            PrimerPrevName = CStr(GibsonResults(7))
            PrimerPrevSeq = CStr(GibsonResults(8))
            PrimerPrevTm = CDbl(GibsonResults(9))
            
        'output results to table
        
            PrimersTableValues(CurrColumn, 1) = NextFragment
            
            PrimersTableValues(CurrColumn, 2) = PrimerNextSeq
            PrimersTableValues(PrevColumn, 3) = PrimerPrevSeq
            
            PrimersTableValues(CurrColumn, 4) = PrimerNextName
            PrimersTableValues(PrevColumn, 5) = PrimerPrevName
            
            PrimersTableValues(CurrColumn, 6) = PrimerNextTm
            PrimersTableValues(PrevColumn, 7) = PrimerPrevTm
            
            PrimersTableValues(CurrColumn, 9) = OverlapSequence
            PrimersTableValues(CurrColumn + 1, 8) = OverlapSequence
                    
        CurrColumn = CurrColumn + 1
        
    Loop
    
    


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
    Dim tempOutput() As Variant
    
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
    
    ReDim tempOutput(1 To AssemblyCount, 1 To ParameterNumber)
    ReDim tempResults(1 To ParameterNumber)
    
    j = 0
    For Each cell In myRange
    
        j = j + 1
        
        If cell.Value <> "" And cell.Value <> 0 Then
        
            Call GibsonRun(cell.Value, tempResults)
            
            For i = 1 To ParameterNumber
                tempOutput(j, i) = tempResults(i)
            Next i
            
        End If
        
    Next cell
    
    Set myRange = myRange.Offset(0, 1).Resize(AssemblyCount, ParameterNumber)
    
    myRange.Value = tempOutput
    
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
    Const TempExtension As String = ".jatmp"
    
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
    FullPythonInputFilename = RunDir & Sep & TempName & TempExtension
    FullPythonOutputFilename = RunDir & Sep & TempName & "_out" & TempExtension
    
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

Dim s As String, e As String
Dim StartIndex As Integer, EndIndex As Integer

Dim Locs As Integer, Loce As Integer
Dim Data As String

Select Case UCase(MarkerType)
    Case "[]", "[", "]", "SQUARE"
        s = "[" & ParameterName & "]"
        e = "[\" & ParameterName & "]"
        Off = Len(s)
    Case Else
        Data = "Not yet supported, sorry. Use ""[]"" for MarkerType"
        GoTo 90
End Select

Locs = InStr(1, Source, s, vbTextCompare)
Loce = InStr(1, Source, e, vbTextCompare)
StartIndex = Locs + Len(s)
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
    Call ExportDataToTextFile(SourceData(1, 1).Value, OutputFile)


End Sub



