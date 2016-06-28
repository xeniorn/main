Attribute VB_Name = "XmodConfirmedCodeCollection"
Option Explicit

'****************************************************************************************************
Sub ChemicalFormulaFormat()

    '====================================================================================================
    'Resets subscripts, and then sets all numbers in selected cells to subscripts
    'Juraj Ahel, 2015-02-05, for KnowledgeBase empirical formula formatting
    'Last update 2015-02-05
    '====================================================================================================
    
    Dim List As Range, cell As Range
    Dim Formula As String
    Dim i As Long
    
    Set List = Application.Selection
    Dim tmp As String
    
    'List.Font.Subscript = True
    'List.Font.Subscript = False
    
    For Each cell In List
        cell.Font.Subscript = False
        For i = 1 To Len(cell)
            tmp = Mid(cell, i, 1)
            If tmp Like "#" Then
                cell.Characters(i, 1).Font.Subscript = True
            Else
                cell.Characters(i, 1).Font.Subscript = False
            End If
        Next i
    Next cell


End Sub


