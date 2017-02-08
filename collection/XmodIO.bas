Attribute VB_Name = "XmodIO"
Option Explicit

Private Const SystemSeparator = "\"

'****************************************************************************************************
Public Function FileSystem_GetDirContents( _
    ByVal InputPath As String, _
    Optional ByVal IncludePath As Boolean = True _
    ) As VBA.Collection
'===============================================================================
'gets a collection of all items in a directory (full path by default)
'Juraj Ahel, 2016-06-15
'Last update 2016-06-15
'===============================================================================

    Dim tString As String
    Dim tPrefix As String
    Dim tColl As VBA.Collection
    
    'InputPath = Left(InputPath, Len(InputPath) - 1)
    
    If Right(InputPath, 1) <> Application.PathSeparator Then
        InputPath = InputPath & Application.PathSeparator
    End If
    
    tString = Dir(InputPath, vbNormal)
    
    If IncludePath Then
        tPrefix = InputPath
    Else
        tPrefix = ""
    End If
    
    Set tColl = New VBA.Collection
    
    Do While tString <> ""
        tColl.Add tPrefix & tString
        tString = Dir
    Loop
    
    Set FileSystem_GetDirContents = tColl
    
    Set tColl = Nothing

End Function

'****************************************************************************************************
Function FileSystem_Unzip(ByVal ZipFilename As String, _
                    Optional ByVal TargetPath As String = "", _
                    Optional ByVal IgnoreExtension As Boolean = False) As String
'===============================================================================
'writes a raw binary string to a file
'Juraj Ahel, 2016-06-07
'Last update 2016-05-08
'===============================================================================
    
    Dim cShellObject As Shell
    'Dim FileNameExists As Boolean
    Dim targetFolder As Folder3
    Dim sourceFile As Folder3
    
    Dim FolderBase As String
    
    'FolderBase = "c:\temp\"
    
    FolderBase = FileSystem_GetTempFolder(IncludeTerminalSeparator:=True)
    
    If Len(TargetPath) = 0 Then
        TargetPath = FolderBase & "zip" & TempTimeStampName & "\" ' & FileSystem_GetFilename(ZipFilename, False) & "\"
    End If
    
    FileSystem_Unzip = ""
    
    'check if target file exists
    If FileSystem_FileExists(ZipFilename) Then
        
        'check if file is a zip file TODO: check this better, not by extension
        If (UCase(FileSystem_GetExtension(ZipFilename)) = "ZIP") Or (IgnoreExtension = True) Then
        
            TargetPath = FileSystem_CreatePath(TargetPath, False)
            
            'Extract the files into the newly created folder
            Set cShellObject = CreateObject("Shell.Application")
            
            'get the folder object corresponding to the path
            Set targetFolder = cShellObject.Namespace(TargetPath)
            Set sourceFile = cShellObject.Namespace(ZipFilename)
            
            'check if the zip is nonempty
            If sourceFile.Items.Count > 0 Then
                'confirm that the temp folder IS empty
                If targetFolder.Items.Count = 0 Then
                    
                    targetFolder.CopyHere sourceFile.Items, _
                    1024 + 16 + 4 'do not display errors + yes to all + don't display dialog box
                    
                    FileSystem_Unzip = TargetPath
                    
                Else
                
                    Debug.Print ("Target folder is not empty! It contains " & _
                        targetFolder.Items.Count & " files.")
                
                End If
            Else
            
                Debug.Print ("Zip file is empty!")
                
            End If
            
        Else
        
            Debug.Print ("Selected file does not have the right extension (.zip / .ZIP)." & _
                        " Extension is: " & _
                        FileSystem_GetExtension(ZipFilename))
            
        End If
        
    Else
    
        Debug.Print ("File not found (" & ZipFilename & ")")
        
    End If
    
    Set sourceFile = Nothing
    Set targetFolder = Nothing
    Set cShellObject = Nothing
    
End Function

'****************************************************************************************************
Function FileSystem_FileExists(ByVal FilePath As String) As Boolean
'===============================================================================
'checks if a file exists
'from StackOverflow http://stackoverflow.com/questions/16351249/vba-check-if-file-exists
'Juraj Ahel, 2016-06-08
'Last update 2016-05-08
'===============================================================================

    If Not Dir(FilePath, vbDirectory) = vbNullString Then
        FileSystem_FileExists = True
    Else
        FileSystem_FileExists = False
    End If

End Function

'****************************************************************************************************
Function FileSystem_DeleteFolder(ByVal FolderPath As String) As Boolean
'===============================================================================
'deletes a folder
'
'Juraj Ahel, 2016-06-08
'Last update 2016-06-08
'2016-06-12 format path string before delete
'===============================================================================

    Dim cFSO As FileSystemObject
    
    If FileSystem_FileExists(FolderPath) Then
        
        Set cFSO = New FileSystemObject
        
        If Right(FolderPath, 1) = SystemSeparator Then
            FolderPath = Left(FolderPath, Len(FolderPath) - 1)
        End If
        
        Call cFSO.DeleteFolder(FolderPath, True)
        
        FileSystem_DeleteFolder = True
        
    Else
        
        FileSystem_DeleteFolder = False
    
    End If
    
    Set cFSO = Nothing

End Function

'****************************************************************************************************
Sub WriteBinaryFileFromString(OutputBuffer As String, _
                    Optional OutputFilename As String = "c:\temp\exportexport.res", _
                    Optional ByVal ExistingFileHandle As Byte = 0)
'===============================================================================
'writes a raw binary string to a file
'Juraj Ahel, 2016-05-10
'Last update 2016-05-10
'===============================================================================


    Call WriteBinaryFile(OutputBuffer, OutputFilename, ExistingFileHandle)
    

End Sub

'****************************************************************************************************
Function CleanUpPath(ByVal FullPath As String, Optional ByVal TargetDirectoryDelimiter As String = "\") As String

'===============================================================================
'formats the path properly
'
'Juraj Ahel, 2016-06-02
'Last update 2016-06-02
'===============================================================================

    Dim DoubleDelimiter As String
    Dim UndesiredDelimiter As String
    
    'parse inputs
    Select Case TargetDirectoryDelimiter
        Case "\"
            UndesiredDelimiter = "/"
        Case "/"
            UndesiredDelimiter = "\"
        Case Else
            Err.Raise 1001, , "Delimiter must be \ or /"
    End Select
    
    DoubleDelimiter = String(2, TargetDirectoryDelimiter)
        
    'replace undesired delimiters with desired ones
    FullPath = Replace(FullPath, UndesiredDelimiter, TargetDirectoryDelimiter)
    
    'remove double delimiters
    'Do While StringCharCount(FullPath, DoubleDelimiter) > 0
    '    FullPath = Replace(FullPath, DoubleDelimiter, TargetDirectoryDelimiter)
    'Loop
    
    CleanUpPath = FullPath

End Function

'****************************************************************************************************
Function FileSystem_GetPath(ByVal FullPath As String, Optional ByVal TargetDirectoryDelimiter As String = "\") As String

'===============================================================================
'extracts just the folder path from full path
'
'Juraj Ahel, 2016-06-02
'Last update 2016-06-02
'===============================================================================
    
    Dim UndesiredDelimiter As String
    
    'parse inputs
    Select Case TargetDirectoryDelimiter
        Case "\"
            UndesiredDelimiter = "/"
        Case "/"
            UndesiredDelimiter = "\"
        Case Else
            Err.Raise 1001, "modIO.FileSystem_GetPath", "Delimiter must be \ or /"
    End Select
    
    FullPath = CleanUpPath(FullPath, TargetDirectoryDelimiter)
        
    FileSystem_GetPath = Left(FullPath, InStrRev(FullPath, TargetDirectoryDelimiter))
    
End Function

'****************************************************************************************************
Function FileSystem_GetExtension(ByVal FullPath As String _
                            ) As String

'===============================================================================
'extracts just the file extension from a full path
'
'Juraj Ahel, 2016-06-08
'Last update 2016-06-08
'===============================================================================
    
    Dim tempCutOff As Long

    tempCutOff = InStrRev(FullPath, ".") + 1
    
    If tempCutOff > 0 Then
        FileSystem_GetExtension = Mid(FullPath, tempCutOff, Len(FullPath))
    Else
        FileSystem_GetExtension = ""
    End If
    
End Function


'****************************************************************************************************
Function FileSystem_GetFilename(ByVal FullPath As String, _
                            Optional ByVal Extension As Boolean = True _
                            ) As String

'===============================================================================
'extracts just the filename from a full path
'
'Juraj Ahel, 2016-06-02
'Last update 2016-06-02
'===============================================================================
    
    Dim tempCutOff As Long
    
    FullPath = CleanUpPath(FullPath)
    
    
    If Not Extension Then
    
        tempCutOff = InStrRev(FullPath, ".") - 1
        
        If tempCutOff > 0 Then
            FullPath = Left(FullPath, tempCutOff)
        End If
        
    End If
    
    FileSystem_GetFilename = Mid(FullPath, InStrRev(FullPath, "\") + 1, Len(FullPath))
    
End Function

'****************************************************************************************************
Function FileSystem_GetTempFolder( _
                Optional ByVal IncludeTerminalSeparator As Boolean = False, _
                Optional wsh As FileSystemObject = Nothing) As String
'===============================================================================
'just grabs the windows default temp folder
'
'Juraj Ahel, 2016-06-08
'Last update 2016-06-08
'===============================================================================
'2016-12-20 Change fixed \ to Application.PathSeparator

    Dim WshWasNothing As Boolean
    
    Dim PathSeparator As String: PathSeparator = Application.PathSeparator
    
    If wsh Is Nothing Then
        WshWasNothing = True
    Else
        WshWasNothing = False
    End If
    
    If WshWasNothing Then
        Set wsh = VBA.CreateObject("Scripting.FileSystemObject")
    End If
    
    FileSystem_GetTempFolder = wsh.GetSpecialFolder(2) '= %temp%
    

    If IncludeTerminalSeparator Then
        FileSystem_GetTempFolder = FileSystem_GetTempFolder & PathSeparator
    End If
    
End Function

'****************************************************************************************************
Function FileSystem_FormatFilename( _
    ByVal FullPath As String, _
    Optional ByVal IncludePath As Boolean = True, _
    Optional ByVal IncludeExtension As Boolean = True, _
    Optional ByVal Quotes As Boolean = False _
    ) As String
'===============================================================================
'formats the full filename
'
'Juraj Ahel, 2016-06-09
'Last update 2016-06-09
'===============================================================================
    
    Const SystemPathSeparator As String = "\"
    
    Dim tempCutOff As Long
    
    If Not IncludePath Then
        tempCutOff = InStrRev(FullPath, SystemPathSeparator) + 1
        If tempCutOff > 0 Then
            FullPath = Mid(FullPath, tempCutOff, Len(FullPath))
        End If
    End If
    
    If Not IncludeExtension Then
        tempCutOff = InStrRev(FullPath, ".") - 1
        If tempCutOff > 0 Then
            FullPath = Left(FullPath, tempCutOff)
        End If
    End If
    
    'remove preexisting quotes
    FullPath = Replace(FullPath, """", "")
    If Quotes Then
        FullPath = """" & FullPath & """"
    End If
    
    FileSystem_FormatFilename = FullPath
    
End Function


'****************************************************************************************************
Function FileSystem_CreatePath(ByVal YourPath As String, _
            Optional ByVal IncludeTerminalSeparator As Boolean = False, _
                Optional wsh As WshShell = Nothing) As String
'===============================================================================
'for a given full file path, creates the full folder hierarchy where it doesn't
'already exist
'it's horrendously slow when calling many files
'Juraj Ahel, 2016-05-31
'Last update 2016-06-07
'===============================================================================


    Dim lastbkSlash As Long
    Dim WshWasNothing As Boolean
    'Dim wsh As Object
    
    If wsh Is Nothing Then
        WshWasNothing = True
    Else
        WshWasNothing = False
    End If
    
    If WshWasNothing Then
        Set wsh = VBA.CreateObject("WScript.Shell")
    End If
    
    lastbkSlash = InStrRev(YourPath, "\")
    
    YourPath = Left(YourPath, lastbkSlash)

    If Dir(YourPath, vbDirectory) = "" Then
        wsh.Run "cmd /c mkdir """ & YourPath & """", 0, True
    End If
    
    'If WshWasNothing Then
    '    Set wsh = Nothing
    'End If
    If IncludeTerminalSeparator Then
        FileSystem_CreatePath = YourPath
    Else
        FileSystem_CreatePath = Left(YourPath, Len(YourPath) - 1)
    End If
    
End Function

Private Function VarToByteString(InputVar As Variant) As String

    Select Case VarType(InputVar)
        Case vbString
            VarToByteString = CStr(InputVar)
        Case Else
            Err.Raise 1001, , "Variable type not yet supported for binary write"
    End Select

End Function

'****************************************************************************************************
Sub WriteBinaryFile(OutputVariable As Variant, _
                    Optional OutputFilename As String = "c:\temp\exportexport.res", _
                    Optional ByVal ExistingFileHandle As Byte = 0)
'===============================================================================
'writes a raw binary string to a file
'Juraj Ahel, 2016-05-07
'Last update 2016-06-09
'===============================================================================

    Dim ActiveFile As Byte
    Dim ByteArray() As Byte
    
    Dim OutputStream As String
    
    Dim i As Long
    
    OutputStream = VarToByteString(OutputVariable)
    
    ReDim ByteArray(0 To Len(OutputStream) - 1)
    
    For i = 0 To Len(OutputStream) - 1
        ByteArray(i) = Asc(Mid(OutputStream, i + 1, 1))
    Next i
        
    'If no file handle was specified
    If ExistingFileHandle = 0 Then
    
        'Get free file handle
        ActiveFile = VBA.FileSystem.FreeFile
        'Open file for writing
        Open OutputFilename For Binary Lock Write As ActiveFile
        'put the data in
        Put ActiveFile, , ByteArray
        Close ActiveFile
        
    Else
    
        'put the data in
        Put ExistingFileHandle, , ByteArray
        
    End If
    

End Sub

'****************************************************************************************************
Function ReadBinaryFile(InputFilename As String) As String
'===============================================================================
'reads the entire contents of a binary file into a string
'Juraj Ahel, 2016-05-06, for reading binary files
'Last update 2016-05-15
'===============================================================================
    
    Dim ActiveFile As Byte
    
    Dim tempString As String
    
    Dim Buffer As String
    
    'TODO: check if filename is valid!
    
    If InputFilename <> "" Then
        
        'Get free file handle
        ActiveFile = VBA.FileSystem.FreeFile
        
        'Open file for reading
        On Error Resume Next
        Open InputFilename For Binary Lock Read As ActiveFile
        
        If Err.Number <> 0 Then
            ReadBinaryFile = vbNullString
            Exit Function
        End If
        
        On Error GoTo 0
            
        
        
        'Set buffer size to be exactly the size of the file (LengthOfFile)
        Buffer = VBA.Strings.Space(VBA.FileSystem.LOF(ActiveFile))
        
        'Load entire file
        Get ActiveFile, , Buffer
        
        ReadBinaryFile = Buffer
        
        Close ActiveFile
        
    Else
        
        Call Err.Raise(1001, , "Filename to be opened cannot be blank")
        
    End If
       
    
End Function

'****************************************************************************************************
Function CreateEmptyFile(ByVal OutputFilename As String) As Boolean
'===============================================================================
'simply creates an empty file, replacing any existing ones if they exist
'Juraj Ahel, 2016-06-09, writing binary files
'Last update 2016-06-09
'===============================================================================

    Dim ActiveFile As Byte

    ActiveFile = VBA.FileSystem.FreeFile

    Open OutputFilename For Output As ActiveFile
    Close ActiveFile

End Function

'****************************************************************************************************
Function WriteTextFile(OutputText As String, OutputFilename As String, Optional ByVal Append As Boolean) As Boolean
'===============================================================================
'reads the entire contents of a binary file into a string
'poorly written, should be redone properly with checks, but there was an error
'and I didn't really feel like debugging
'Juraj Ahel, 2016-05-31, writing text files
'Last update 2016-05-31
'2016-06-27 changed default file handle to 10
'2016-12-20 make it select a free file handle
'===============================================================================
        
    Dim ActiveFile As Integer
        
    'Get free file handle
    ActiveFile = VBA.FileSystem.FreeFile()
        
    If Append Then
        Open OutputFilename For Append As #ActiveFile
    Else
        Open OutputFilename For Output As #ActiveFile
    End If
        

    Print #ActiveFile, OutputText
    Close #ActiveFile
    
End Function

'****************************************************************************************************
Function ReadTextFile(InputFilename As String) As String
'===============================================================================
'reads the entire contents of a binary file into a string
'Juraj Ahel, 2016-05-06, for reading binary files
'Last update 2016-05-15
'===============================================================================
    
    Dim ActiveFile As Byte
    
    Dim tempString As String
    
    Dim Buffer As String
    
    'TODO: check if filename is valid!
    
    If InputFilename <> "" Then
        
        'Get free file handle
        ActiveFile = VBA.FileSystem.FreeFile
        
        'Open file for reading
        Open InputFilename For Input As ActiveFile
        
        'Set buffer size to be exactly the size of the file (LengthOfFile)
        Buffer = VBA.Strings.Space(VBA.FileSystem.LOF(ActiveFile))
        
        'Load entire file
        Buffer = Input$(LOF(ActiveFile), ActiveFile)
        
        ReadTextFile = Buffer
        
        Close ActiveFile
        
    Else
    
        Call Err.Raise(1001, , "Filename to be opened cannot be blank")
        
    End If
       
    
End Function

'****************************************************************************************************
Sub ExportSeqToTXT()

'====================================================================================================
'Exports cell column pairs formated as [HEADER][SEQUENCE] in simple FASTA format
'under a file name corresponding to the header
'
'Juraj Ahel, 2015-02-10, for exporting sequences, to have an external database
'Last update 2015-08-26
'====================================================================================================

Dim FilePath As String, OutputFile As String
Dim DataSource As Range
Dim HeaderLine As String
Dim Sequence As String

Dim i As Long

FilePath = "C:\Excel_outputs\Sequences\"

Set DataSource = Selection

For i = 1 To DataSource.Rows.Count
    HeaderLine = ">" & CStr(DataSource(i, 1).Value)
    Sequence = DataSource(i, 2).Value
    OutputFile = FilePath & CStr(DataSource(i, 1).Value) & "_seq.txt"
    Call WriteTextFile(HeaderLine & vbCrLf & Sequence, OutputFile)
Next i

End Sub

'****************************************************************************************************
Sub ExportToTXT( _
                SourceData As Range, _
                Optional FilePath As String = "C:\Excel_outputs\", _
                Optional FilenameBase As String = "ExcelOutput ", _
                Optional Extension As String = ".txt" _
               )

'====================================================================================================
'Exports a separate text file for each cell in selection
'Still needs to be added modular file naming, now it's always "Fragment #.txt"
'
'Juraj Ahel, 2015-02-10, for Gibson assembly and general purposes
'Last update 2015-02-11
'====================================================================================================

Dim OutputFile As String
Dim DataSource As Range



    OutputFile = FilePath & FilenameBase & Extension
    Call WriteTextFile(SourceData(1, 1).Value, OutputFile)


End Sub



'****************************************************************************************************
Sub ExportToTXTSequence(SourceData As Range, Optional FilePath As String = "C:\Excel_outputs\", Optional FilenameBase As String = "ExcelOutput ", Optional Extension As String = ".txt")

'====================================================================================================
'Exports a separate text file for each cell in selection
'Still needs to be added modular file naming, now it's always "Fragment #.txt"
'
'Juraj Ahel, 2015-02-10, for Gibson assembly and general purposes
'Last update 2015-02-11
'====================================================================================================

Dim OutputFile As String
Dim DataSource As Range

Dim i As Long

Set DataSource = SourceData

For i = 1 To DataSource.Rows.Count
    OutputFile = FilePath & FilenameBase & i & Extension
    Call WriteTextFile(DataSource(i, 1).Value, OutputFile)
Next i

End Sub

Sub ExportToTXTMacro()

Dim SourceData As Range
Dim FilePath As String, FilenameBase As String

FilePath = "C:\Excel_outputs\"
FilenameBase = "Fragment "

Set SourceData = Selection


Call ExportToTXTSequence(SourceData, FilePath, FilenameBase, ".txt")

End Sub


