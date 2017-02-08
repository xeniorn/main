Attribute VB_Name = "XModGraphs"
Option Explicit

'****************************************************************************************************
Sub ModifyTableData()

'====================================================================================================
'Truncates a data table in column or row dimension
'Thins columns and rows from the table using selected thinning factor
'Juraj Ahel, 2015-04-28
'Last update 2015-04-28
'====================================================================================================

    
    Dim InputTable As Range, OutputTable As Range
    Dim InputArray() As Variant, OutputArray() As Variant
    
    Dim ParameterString As String
    Dim StartRow As Long, EndRow As Long
    Dim StartColumn As Long, EndColumn As Long
    Dim RowThinFactor As Long, ColumnThinFactor As Long
    
    Dim Parameters
    
    Set InputTable = Application.InputBox("Select the table (or a part of it) to modify:", "Input selection", Type:=8)
    
    ParameterString = InputBox("Start row; End row; Start column; End column; Keep every N-th row; Keep every N-th column." _
                            & vbCrLf & "Defaults give out an exact copy. Zero-end means full-length.", _
                            "Input parameters", _
                            "1;0;1;0;1;1")
    
    Set OutputTable = Application.InputBox("Select the top left corner of the location where you want the processed table to be output.", _
                                            "Target output", Type:=8)
    
    
    Parameters = Split(ParameterString, ";")
    
    StartRow = Parameters(0)
    EndRow = Parameters(1)
    StartColumn = Parameters(2)
    EndColumn = Parameters(3)
    RowThinFactor = Parameters(4)
    ColumnThinFactor = Parameters(5)
    
    Dim InputTableWidth As Long, InputTableHeight As Long
    Dim OutputTableWidth As Long, OutputTableHeight As Long
    
    InputTableWidth = InputTable.Columns.Count
    InputTableHeight = InputTable.Rows.Count
    
    If EndColumn > InputTableWidth Or EndColumn = 0 Then EndColumn = InputTableWidth
    If EndRow > InputTableHeight Or EndRow = 0 Then EndRow = InputTableHeight
    
    OutputTableWidth = 1 + EndColumn - StartColumn
    OutputTableWidth = 1 + Int((OutputTableWidth - 1) / ColumnThinFactor)
    OutputTableHeight = 1 + EndRow - StartRow
    OutputTableHeight = 1 + Int((OutputTableHeight - 1) / RowThinFactor)
    
    'If OutputTableWidth * OutputTableHeight = 0 Then
    '    MsgBox ("Zero-length output. Check input parameters and table selection.")
    '    GoTo 999
    'End If
    
    ReDim InputArray(1 To InputTableHeight, 1 To InputTableWidth)
    ReDim OutputArray(1 To OutputTableHeight, 1 To OutputTableWidth)
    
    Dim TempFormula As String, TempColor As Long, TempPattern As Long
    Dim ProceedCheck
    
    TempFormula = OutputTable.Value
    TempColor = OutputTable.Interior.Color
    TempPattern = OutputTable.Interior.Pattern
    
    OutputTable.Formula = "!SELECTION!"
    OutputTable.Interior.Color = vbRed
    OutputTable.Interior.Pattern = xlSolid
    
    
    ProceedCheck = MsgBox( _
                          "The macro will replace the area " & _
                          OutputTableHeight & " columns below and " & _
                          OutputTableWidth & " rows to the right of the top-left cell of the selection" _
                          & vbCrLf & "Any contents will be lost. Proceed?", _
                          vbYesNo, _
                          "Confirm replacement:" _
                         )
    
    OutputTable.Formula = TempFormula
    OutputTable.Interior.Color = TempColor
    OutputTable.Interior.Pattern = TempPattern
    
    If ProceedCheck = vbNo Then
        MsgBox ("Input aborted.")
        GoTo 999
    End If
    
    InputArray = InputTable.Value
    
    Dim i As Long, j As Long, m As Long, N As Long
    
    For i = 1 To OutputTableHeight
        m = StartRow + (i - 1) * RowThinFactor
        For j = 1 To OutputTableWidth
            N = StartColumn + (j - 1) * ColumnThinFactor
            OutputArray(i, j) = InputArray(m, N)
        Next j
    Next i
    
    Set OutputTable = OutputTable.Resize(OutputTableHeight, OutputTableWidth)
    
    OutputTable.Value = OutputArray

999 End Sub

