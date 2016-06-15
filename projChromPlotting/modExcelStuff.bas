Attribute VB_Name = "modExcelStuff"
Option Explicit

Function SheetExists(ByVal shtName As String, Optional ByVal wb As Workbook) As Boolean
'===============================================================================
'checks whether a sheet with a certain name exists in a workbook
'Juraj Ahel, 2016-05-15
'Last update 2016-05-15
'===============================================================================

   Dim sht As Excel.Worksheet

    If wb Is Nothing Then Set wb = Excel.ActiveWorkbook
    
    On Error Resume Next
       Set sht = wb.Sheets.Item(shtName)
    On Error GoTo 0
       SheetExists = Not sht Is Nothing
    
End Function


Function CreateSheetFromName(Optional ByVal SheetName As String = "Sheet", _
                            Optional ByVal Spacer As String = "_", _
                            Optional ByVal TargetWorkbook As Excel.Workbook = Nothing) As String
'===============================================================================
'creates a new sheet with a given name in a target workbook, adding _### iterations
'of the name if it exists already
'Juraj Ahel, 2016-05-15
'Last update 2016-05-18
'===============================================================================
    
    
    Const conDesiredSheetIndexFormat As String = "000"
    Const conMaximumSheetsAllowed As Long = 100
    
    Dim i As Long
    Dim TempName As String
    Dim AddedSheet As Excel.Worksheet
    
    i = 0
    TempName = SheetName
    
    Do While SheetExists(TempName) And i <= conMaximumSheetsAllowed
        i = i + 1
        TempName = SheetName & Spacer & Format(i, conDesiredSheetIndexFormat)
    Loop
    
    If i = conMaximumSheetsAllowed Then
        Call Err.Raise("1001", "You sure have a lot of sheets named " & SheetName & Spacer & "xxx...")
    Else
        If TargetWorkbook Is Nothing Then Set TargetWorkbook = Excel.ActiveWorkbook
        Set AddedSheet = TargetWorkbook.Sheets.Add
        AddedSheet.Name = TempName
        CreateSheetFromName = TempName
    End If
    

End Function
