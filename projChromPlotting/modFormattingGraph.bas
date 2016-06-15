Attribute VB_Name = "modFormattingGraph"
Option Explicit

Private Const Pi As Double = 3.14159265358979

Sub Graphing_FormattedGraph()

    Dim SelectedRange As Excel.Range
    Dim GroupDefinitionRange As Excel.Range
    Dim XTempRange As Excel.Range, YTempRange As Excel.Range
    Dim TitlesRange As Excel.Range, TempTitleRange As Excel.Range
    
    Dim tempFont As Excel.Font
    
    Dim TargetChart As Excel.ChartObject
    Dim TempSeries As Excel.Series
    
    Dim i As Long
    Dim H As Double, S As Double, V As Double
    
    Dim tempGroupCounter As String
    Dim currentGroup As String
    
    'On Error GoTo ErrorHandler
    '    Application.EnableEvents = False
            
        Set SelectedRange = Selection
        
        Set GroupDefinitionRange = SelectedRange.Offset(0, 1).Resize(1, SelectedRange.Columns.Count - 1)
        Set TitlesRange = GroupDefinitionRange.Offset(1, 0)
        
        Set XTempRange = SelectedRange.Offset(2, 0).Resize(SelectedRange.Rows.Count - 2, 1)
        
        Set TargetChart = Chromatography_AddChart(TargetWorksheet:=ActiveWorkbook.ActiveSheet)
        
        
        H = 0
        S = 0.9
        V = 1
        
        For i = 1 To TitlesRange.Count
            
            Set TempTitleRange = TitlesRange.Offset(0, i - 1).Resize(1, 1)
            Set YTempRange = XTempRange.Offset(0, i)
            
            currentGroup = CStr(GroupDefinitionRange.Cells(1, i).Value)
            tempGroupCounter = tempGroupCounter & currentGroup
            
            
            If currentGroup = "" Then
                S = 0
                H = 0
                V = 0.5
            Else
                S = 1
                'hue is the same for all members of a plot group
                H = (Val(currentGroup) - 1) * Pi / 3
                'color gets darker and darker with each plot
                V = (1 - (StringCharCount(tempGroupCounter, currentGroup) - 1) / 5)
            End If
            
            Set TempSeries = ChartAddSeriesToChart(XTempRange, _
                                                    YTempRange, _
                                                    TargetChart, _
                                                    TempTitleRange)
            With TempSeries
            
                With .Format.Line
                    .ForeColor.RGB = ColorFromHSV(H, S, V)
                    .Weight = xlHairline
                End With
                
                .Smooth = False
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerBackgroundColor = ColorFromHSV(H, S, V)
                .MarkerSize = 3
                
            End With
            
        Next i
        
        With TargetChart.Chart
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = XTempRange.Offset(-1, 0).Resize(1, 1).Value
            .Legend.Font.Size = .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size * 0.75
            
            For i = 1 To .Legend.LegendEntries.Count
                Set tempFont = .Legend.LegendEntries(i).Font
                With tempFont
                    .Color = TargetChart.Chart.SeriesCollection(i).Format.Line.ForeColor
                    
                    '.Line.Visible = msoTrue
                    '.Line.Weight = 0
                    '.Line.Color = vbBlack
                End With
            Next i
        End With
        
        SelectedRange.Select
    
ErrorHandler:
    Application.EnableEvents = True
        
    
    Set XTempRange = Nothing
    Set YTempRange = Nothing
    Set SelectedRange = Nothing
    Set GroupDefinitionRange = Nothing
    Set TitlesRange = Nothing
    Set TempTitleRange = Nothing
    Set TargetChart = Nothing
    Set TempSeries = Nothing

End Sub
