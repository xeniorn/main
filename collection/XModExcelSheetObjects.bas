Attribute VB_Name = "XModExcelSheetObjects"
Option Explicit


'****************************************************************************************************
Sub DeleteAllChartsOnSheet()
'====================================================================================================
'Deletes all charts on the active sheet
'Juraj Ahel, 2015-04-24
'Last update 2015-04-24
'====================================================================================================


Dim chtObj As ChartObject
For Each chtObj In ActiveSheet.ChartObjects
chtObj.Delete
Next

End Sub



'****************************************************************************************
Sub RangeColumnInvert()

'====================================================================================================
'Inverts the row order within the selected columns
'Copies formulas "stupidly", raw-copy style, so "A2 + A3" will stay as such and won't switch to "A3 + A4"
'Juraj Ahel, 2015-04-17, for general purposes
'Last update 2015-04-17
'====================================================================================================

Dim TargetRange As Range
Dim NRow As Long, NColumn As Long
Dim RangeDataIn() As Variant, RangeDataOut() As Variant
Dim i As Long, j As Long, k As Long

Set TargetRange = Selection

NRow = TargetRange.Rows.Count
NColumn = TargetRange.Columns.Count

ReDim RangeDataIn(1 To NRow, 1 To NColumn)
ReDim RangeDataOut(1 To NRow, 1 To NColumn)

RangeDataIn = TargetRange.Formula

For i = 1 To NRow

    k = NRow - i + 1
    
    For j = 1 To NColumn
        RangeDataOut(k, j) = RangeDataIn(i, j)
    Next j
    
Next i

TargetRange.Formula = RangeDataOut

End Sub
