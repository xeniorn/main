Attribute VB_Name = "XmodTools"
Option Explicit

'****************************************************************************************************
Public Function TempTimeStampName() As String

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
Public Sub CallProgram( _
                ProgramCommand As String, _
                Optional ProgramPath As String = "", _
                Optional ArgList As String = "", _
                Optional WaitUntilFinished As Boolean = True, _
                Optional WindowMode As String = "1", _
                Optional RunDirectory As String = "", _
                Optional RunAsRawCmd As Boolean = True, _
                Optional OutputFile As String = "", _
                Optional InputFile As String = "" _
               )

'====================================================================================================
'Calls an external program under the windows environment, using windows scripting host (Wsh)
'Takes more intuitive inputs and does all the complicated mimbo-jimbo so the code calling it is clean
'Juraj Ahel, 2015-02-11, for Gibson assembly and general purposes
'Last update 2015-02-11
'2016-06-28 tranfered to XmodTools module, added explicit var declaration
'====================================================================================================
'Made for Excel Professional Plus 2013 under Windows 8.1
'2016-12-20 add support for input files

    Dim wsh As Object
    Dim WaitOnReturn As Boolean: WaitOnReturn = WaitUntilFinished
    Dim WindowVisibilityType As Long
    Dim RunCommand As String, ProgramFullPath As String, ParsedArguments As String
    Dim ProgramCommandTemp As String, ProgramPathTemp As String
    
    Dim ParsedRunDirectory As String
    
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
    
    RunCommand = """" & ProgramFullPath & """ " & ParsedArguments
    
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
    If ParsedRunDirectory = "" Then ParsedRunDirectory = FileSystem_GetTempFolder
    
    wsh.CurrentDirectory = ParsedRunDirectory
    
    If RunAsRawCmd Then RunCommand = "%comspec% /c " & RunCommand
    'If RunAsRawCmd Then RunCommand = "%comspec% /k " & RunCommand
        
    If InputFile <> "" Then RunCommand = RunCommand & " <""" & InputFile & """"
    If OutputFile <> "" Then RunCommand = RunCommand & " >""" & OutputFile & """"
    
    '2>&1 at the end ensures that the error log will be appended to the result! Cool!
    RunCommand = RunCommand & " 2>&1"
    
    Call wsh.Run(RunCommand, WindowVisibilityType, WaitOnReturn)

End Sub

'****************************************************************************************************
Public Sub TestCallProgram( _
                ByVal ProgramCommand As String, _
                Optional ByVal ProgramPath As String = "", _
                Optional ByVal ArgList As String = "", _
                Optional ByVal WaitUntilFinished As Boolean = True, _
                Optional ByVal WindowMode As String = "1", _
                Optional ByVal RunDirectory As String = "", _
                Optional ByVal RunAsRawCmd As Boolean = True, _
                Optional ByVal OutputFile As String = "", _
                Optional ByVal InputFile As String = "" _
               )

    Dim RunCommand As String
    
    RunCommand = """" & ProgramPath & ProgramCommand & """ " & ArgList
    
    If InputFile <> "" Then RunCommand = RunCommand & " <""" & InputFile & """"
    If OutputFile <> "" Then RunCommand = RunCommand & " >""" & OutputFile & """"
    
    '2>&1 at the end ensures that the error log will be appended to the result! Cool!
    RunCommand = RunCommand & " 2>&1"
    
    Call Shell(RunCommand, vbNormalFocus)
    
    

End Sub

'****************************************************************************************************
Public Function DTT(x, Optional y = 0, Optional DateFormat As String = "YYMMDDhhmm", _
             Optional RoundingMode As Long = -1, Optional Output As String = "d")

'====================================================================================================
'Converts YYMMDDhhmm date/time format to excel's date format
'if there is y, then calculates difference x-y instead of absolute date
'other date formats possibly to be added
'Juraj Ahel, 2014-06-08, for Master's thesis
'Last update 2014-12-31
'====================================================================================================


    Dim yearx As Long, monthx As Long, dayx As Long, hourx As Long, minutex As Long, secondx As Single
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
        
        Dim yeary As Long, monthy As Long, dayy As Long, houry As Long, minutey As Long, secondy As Single
        
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
