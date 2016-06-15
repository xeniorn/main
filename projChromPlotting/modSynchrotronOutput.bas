Attribute VB_Name = "modSynchrotronOutput"
Option Explicit


Sub ExportFiles()


    Dim InpRange As Excel.Range
    Dim InpArray()
    Dim wsh As Object
    
    Dim RowN As Long, i As Long, ColumnN As Long
    
    Dim TextToWrite As String
    Dim FinalPath As String
    Dim BasePath As String
    
    Dim ts As String
    
    Dim subfolder As String
    
    BasePath = "C:\Temp\Excel_outputs\test1\"
    
    Set InpRange = Selection
    
    Set wsh = VBA.CreateObject("WScript.Shell")
       
    RowN = InpRange.Rows.Count
    ColumnN = InpRange.Columns.Count
       
    ReDim InpArray(1 To RowN, 1 To ColumnN)
       
    InpArray = InpRange.Value
        
    'first create all the paths if necessary (that way it's faster
    'because it can be done asynchronously
    For i = 1 To RowN
    
        ts = InpArray(i, 1)
    
        FinalPath = BasePath & "\" & ts & ".description"
        
        Do While StringCharCount(FinalPath, "\\") > 0
            FinalPath = Replace(FinalPath, "\\", "\")
            InpArray(i, 1) = FinalPath
        Loop
        
        Call FileSystem_CreatePath(FinalPath, wsh)
        
    Next i
    
    'then write the actual files / contents
    For i = 1 To RowN
    
        ts = InpArray(i, 1)
    
        FinalPath = ts
        
        TextToWrite = Mid(ts, InStrRev(ts, "\") + 1, Len(ts))
    
        Call WriteTextFile(TextToWrite, FinalPath)
        
        Debug.Print (FinalPath)
        
    Next i
    
    Set wsh = Nothing


End Sub
