Attribute VB_Name = "Macros"
'Juraj Ahel, 2015-02-01
'Last Update, 2016-02-17

Const DebugMode As Boolean = True

Const ExcelExportFolder As String = "C:\ExcelExports"

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
Sub ExportToTXTSequence(SourceData As Range, Optional FilePath As String = ExcelExportFolder, Optional FileNameBase As String = "ExcelOutput ", Optional Extension As String = ".txt")

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
    OutputFile = FilePath & FileNameBase & i & Extension
    Call ExportDataToTextFile(DataSource(i, 1).Value, OutputFile)
Next i

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

Dim Wsh As Object
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
Set Wsh = VBA.CreateObject("WSCript.Shell")

ParsedRunDirectory = RunDirectory
Wsh.CurrentDirectory = ParsedRunDirectory

If RunAsRawCmd Then RunCommand = "%comspec% /c " & RunCommand

'2>&1 at the end ensures that the error log will be appended to the result! Cool!
If OutputFile <> "" Then RunCommand = RunCommand & " >""" & OutputFile & """ 2>&1"

a = Wsh.Run(RunCommand, WindowVisibilityType, WaitOnReturn)

End Sub
Sub ExportToTXTMacro()

Dim SourceData As Range
Dim FilePath As String, FileNameBase As String

FilePath = ExcelExportFolder
FileNameBase = "Fragment_"

Set SourceData = Selection


Call ExportToTXTSequence(SourceData, FilePath, FileNameBase, ".txt")

End Sub

'****************************************************************************************************
Sub ExportToTXT(SourceData As Range, Optional FilePath As String = ExcelExportFolder, _
                Optional FileNameBase As String = "ExcelOutput ", Optional Extension As String = ".txt")

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

    OutputFile = FilePath & FileNameBase & Extension
    Call ExportDataToTextFile(SourceData(1, 1).Value, OutputFile)


End Sub


Sub ExportDataToTextFile(DataToOutput As String, OutputFilename As String)

Open OutputFilename For Output As #1

Print #1, DataToOutput
Close #1

End Sub

