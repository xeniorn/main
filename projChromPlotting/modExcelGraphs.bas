Attribute VB_Name = "modExcelGraphs"
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2016-xx-xx
'Last update 2016-05-18
'====================================================================================================
Option Explicit

Private Const conClassName As String = "modExcelGraphs"


'****************************************************************************************************
Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)
'====================================================================================================
'shows an inputbox - pasted code will be converted to a string declaration to be used in VBA
'Juraj Ahel, 2016-05-09
'Last update 2016-05-09
'====================================================================================================

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub

Sub ChartApplySettingsToSeries(ByRef TargetSeries As Excel.Series, _
                                ByRef Settings As clsSeriesFormatSettings)
                                
    With TargetSeries
        
        'general
        .Smooth = Settings.LineSmooth
        
        'marker
        .MarkerSize = Settings.MarkerSize
        .MarkerStyle = Settings.MarkerType 'includes "no marker as a possible value
        .MarkerForegroundColor = Settings.MarkerColor
                
        'line
        With .Format.Line
            .Visible = Settings.LineVisible
            .Transparency = Settings.Transparency
            .Weight = Settings.LineWeight
            .ForeColor.RGB = Settings.LineColor
        End With
    
    End With
        
End Sub

'****************************************************************************************************
Function ChartAddSeriesToChartDirect(ByRef x1() As Double, _
                        ByRef y1() As Double, _
                        ByRef TargetChart As Excel.ChartObject, _
                        Optional ByRef titlerange As String = "", _
                        Optional ByVal AxisChoice As Long = xlPrimary, _
                        Optional ByVal Settings As clsSeriesFormatSettings _
                        ) As Excel.Series

'====================================================================================================
'adds a series (x range, y range, title) to a chart, either to primary or secondary axis
'directly from an array of data
'allows immediate application of settings
'Juraj Ahel, 2016-05-15
'Last update 2016-05-18
'====================================================================================================

    Dim ChartSeries As Excel.Series
    
    Dim xcount As Long, ycount As Long
    Dim xcolumns As Long, ycolumns As Long
    Dim xAddress As String, yAddress As String, labelAddress As String
    
    Dim NextSeriesIndex As Long
    
    Dim Sep As String
    
    'separator in formulas can be ";" or ",", depending on the computer!
    'this defines it programatically
    'Sep = Excel.Application.International(xlListSeparator)
    'DOESN'T WORK, BECAUSE IN VBA IT ALWAYS USES COMMA, REGARDLESS OF LOCALE!
    Sep = ","
    
    '================Input parsing================
    If (AxisChoice <> xlPrimary) And (AxisChoice <> xlSecondary) Then
        Call ErrorReport(, "Axis choice must be either xlPrimary or xlSecondary")
    End If
    
    If TargetChart Is Nothing Then
        Call ErrorReport(, "TargetChart cannot be 'Nothing'.")
    End If
    
    xcount = UBound(x1) - LBound(x1) + 1
    ycount = UBound(y1) - LBound(y1) + 1
    
    'check array dimensions...
            
    If xcount = 0 Or ycount = 0 Then
        Call ErrorReport(, "Range to be imported into the series cannot be empty")
    End If
    
    If xcount <> ycount Then
        Call ErrorReport(, "x and y data ranges must be of the same length")
    End If
    
    If Settings Is Nothing Then Set Settings = New clsSeriesFormatSettings
    
    '================Main================
    With TargetChart.Chart
        
        'add a new series
        NextSeriesIndex = .FullSeriesCollection.Count + 1
        Set ChartSeries = .SeriesCollection.NewSeries
        '.SeriesCollection.NewSeries
        
        'define the new series

        With ChartSeries
            .XValues = x1
            .Values = y1
            .Name = titlerange
            .AxisGroup = AxisChoice
        End With
        
        Call ChartApplySettingsToSeries(ChartSeries, Settings)
            
    End With


    '================Cleanup================
    Set ChartAddSeriesToChartDirect = ChartSeries
    Set ChartSeries = Nothing

End Function


'****************************************************************************************************
Function ChartAddSeriesToChart(ByRef x1 As Excel.Range, _
                        ByRef y1 As Excel.Range, _
                        ByRef TargetChart As Excel.ChartObject, _
                        Optional ByRef titlerange As Excel.Range = Nothing, _
                        Optional ByVal AxisChoice As Long = xlPrimary, _
                        Optional ByVal DefineValuesDirectly As Boolean = False _
                        ) As Excel.Series

'====================================================================================================
'adds a series (x range, y range, title) to a chart, either to primary or secondary axis
'it can also either draw it normally (referencing cells), or putting the values into the chart itself
'it also returns the series object, for convenience, if desired
'Juraj Ahel, 2016-05-15
'Last update 2016-05-17
'====================================================================================================

    Dim ChartSeries As Excel.Series
    
    Dim xcount As Long, ycount As Long
    Dim xcolumns As Long, ycolumns As Long
    Dim xAddress As String, yAddress As String, labelAddress As String
    
    Dim NextSeriesIndex As Long
    
    Dim Sep As String
    
    'separator in formulas can be ";" or ",", depending on the computer!
    'this defines it programatically
    'Sep = Excel.Application.International(xlListSeparator)
    'DOESN'T WORK, BECAUSE IN VBA IT ALWAYS USES COMMA, REGARDLESS OF LOCALE!
    Sep = ","
    
    '================Input parsing================
    If (AxisChoice <> xlPrimary) And (AxisChoice <> xlSecondary) Then
        Call ErrorReport(, "Axis choice must be either xlPrimary or xlSecondary")
    End If
    
    If TargetChart Is Nothing Then
        Call ErrorReport(, "TargetChart cannot be 'Nothing'.")
    End If
    
    If x1 Is Nothing Or y1 Is Nothing Then
        Call ErrorReport(, "Ranges containing x and y data cannot be 'nothing'.")
    End If
    
    xcount = x1.Count
    ycount = y1.Count
    
    xcolumns = x1.Columns.Count
    ycolumns = y1.Columns.Count
    
    If xcolumns > 1 Or ycolumns > 1 Then
        Call ErrorReport(, "Both axes data ranges need to be a single-column range")
    End If
        
    If xcount = 0 Or ycount = 0 Then
        Call ErrorReport(, "Range to be imported into the series cannot be empty")
    End If
    
    If xcount = 0 Or ycount = 0 Then
        Call ErrorReport(, "Range to be imported into the series cannot be empty")
    End If
    
    If xcount <> ycount Then
        Call ErrorReport(, "x and y data ranges must be of the same length")
    End If
    
    
    '================Main================
    With TargetChart.Chart
        
        'add a new series
        NextSeriesIndex = .FullSeriesCollection.Count + 1
        Set ChartSeries = .SeriesCollection.NewSeries
        '.SeriesCollection.NewSeries
        
        'define the new series
        If DefineValuesDirectly Then
        
            With ChartSeries
                .XValues = x1.Value
                .Values = y1.Value
                .Name = titlerange.Value
                .AxisGroup = AxisChoice
            End With
            
        Else
            
            xAddress = x1.Worksheet.Name & "!" & x1.Address(ReferenceStyle:=xlR1C1)
            yAddress = y1.Worksheet.Name & "!" & y1.Address(ReferenceStyle:=xlR1C1)
            labelAddress = titlerange.Worksheet.Name & "!" & titlerange.Address(ReferenceStyle:=xlR1C1)
            Dim tempstr As String
            tempstr = "=SERIES(" & labelAddress & Sep & xAddress & Sep & yAddress & Sep & NextSeriesIndex & ")"
            
            With ChartSeries
                '.Formula = "=SERIES(" & labelAddress & Sep & xAddress & Sep & yAddress & Sep & NextSeriesIndex & ")"
                '.Formula = "=SERIES(Unicorn_002!R9C6,Unicorn_002!R1C1:R150C1,Unicorn_002!R1C2:R150C2,7)"
                .Formula = tempstr
                '.XValues = x1
                '.Values = y1
                '.Name = titlerange 'TODO: this still doesn't reference the name...
                .AxisGroup = AxisChoice
            End With
            
        End If
            
    End With


    '================Cleanup================
    Set ChartAddSeriesToChart = ChartSeries
    Set ChartSeries = Nothing

End Function

Sub Chromatography_RemoveHack(TargetChart As Excel.Chart)

    Dim srs As Excel.Series
    Dim i As Long
    
    Do
    
        For i = 1 To TargetChart.SeriesCollection.Count
            Set srs = TargetChart.SeriesCollection.Item(i)
            If srs.Name = "I_AM_A_HACK_DELETE_ME" Then
                srs.Delete
                Exit For
            End If
        Next i
        
    Loop Until i >= TargetChart.SeriesCollection.Count
    
    Set srs = Nothing
    
End Sub

'****************************************************************************************************
Function Chromatography_AddChart(Optional ByVal IncludeSecondaryAxis As Boolean = False, _
                                    Optional TargetWorksheet As Excel.Worksheet = Nothing, _
                                    Optional Settings As clsChartFormatSettings = Nothing, _
                                    Optional AxisSettings As clsAxisFormatSettings = Nothing _
                                ) As Excel.ChartObject

'====================================================================================================
'
'
'
'Juraj Ahel, 2014-06-08, for Master's thesis
'Last update 2016-05-22
'====================================================================================================
'TODO: gather all the settings (axis weight, color, tick width, etc...) into an object file which can be
'passed to the function, and stored separately / exported / edited!

    Dim cht As Excel.ChartObject
    Dim srs As Series
    
    'if no Settings object provided, use the default settings (class is initialized with default values
    If Settings Is Nothing Then Set Settings = New clsChartFormatSettings
    
    'if no AxisSettings object provided, use the default settings (class is initialized with default values
    If AxisSettings Is Nothing Then Set AxisSettings = New clsAxisFormatSettings
       
    'if target Worksheet is not specified, do it in the worksheet in focus
    If TargetWorksheet Is Nothing Then Set TargetWorksheet = ActiveWorkbook.ActiveSheet
    
    'if I want it in the sheet
    Set cht = TargetWorksheet.ChartObjects.Add(Left:=0, Width:=800, Top:=0, Height:=500)
    
    'ChartName = Mid(cht.Name, InStr(1, cht.Name, "Chart"), 1000)
    
    'if I want it as a separate object
    'Set cht = Charts.Add()
    
    With cht.Chart
        
        'e.g. XY scatter
        .ChartType = Settings.ChartType
        '.ChartType = xlXYScatter
        
        With .Axes(xlValue, xlPrimary)
            
            .HasTitle = AxisSettings.HasTitle
            .AxisTitle.Font.Color = AxisSettings.AxisColor
            .AxisTitle.Font.Size = AxisSettings.TitleFontSize
            
            .MajorTickMark = AxisSettings.MajorTickType
            .MinorTickMark = AxisSettings.MinorTickType
            
            .Border.Weight = AxisSettings.AxisWeight 'before it was xlThick (=4)
            .Border.Color = AxisSettings.AxisColor
            
        End With
        
        If IncludeSecondaryAxis Then
            
            'a hack, just add an invisible series...
                Set srs = .SeriesCollection.NewSeries
                    srs.XValues = Array(0)
                    srs.Values = Array(0)
                    srs.AxisGroup = xlPrimary
                    srs.Name = "I_AM_A_HACK_DELETE_ME"
                    
                Set srs = .SeriesCollection.NewSeries
                    srs.XValues = Array(0)
                    srs.Values = Array(0)
                    srs.AxisGroup = xlSecondary
                    srs.Name = "I_AM_A_HACK_DELETE_ME"
                
                .HasAxis(xlCategory, xlPrimary) = True
                .HasAxis(xlValue, xlPrimary) = True
                            
                '.HasAxis(xlCategory, xlSecondary) = True
                .HasAxis(xlValue, xlSecondary) = True
            
            With .Axes(xlValue, xlSecondary)
            
                .HasTitle = AxisSettings.HasTitle
                .AxisTitle.Font.Color = AxisSettings.AxisColor
                .AxisTitle.Font.Size = AxisSettings.TitleFontSize
                
                .MajorTickMark = AxisSettings.MajorTickType
                .MinorTickMark = AxisSettings.MinorTickType
                
                .Border.Weight = AxisSettings.AxisWeight
                .Border.Color = AxisSettings.AxisColor
                
            End With
            
        End If
        
        With .Axes(xlCategory)
            
            .HasTitle = AxisSettings.HasTitle
            .AxisTitle.Font.Color = AxisSettings.AxisColor
            .AxisTitle.Font.Size = AxisSettings.TitleFontSize
            
            .MajorTickMark = AxisSettings.MajorTickType
            .MinorTickMark = AxisSettings.MinorTickType
            
            .Border.Weight = AxisSettings.AxisWeight 'before it was xlThick (=4)
            .Border.Color = AxisSettings.AxisColor
                
        End With
          
        .Axes(xlCategory).TickLabels.Font.Size = AxisSettings.TickLabelSize
        .Axes(xlValue, 1).TickLabels.Font.Size = AxisSettings.TickLabelSize
        
        If IncludeSecondaryAxis Then
        
            .Axes(xlValue, 2).TickLabels.Font.Size = AxisSettings.TickLabelSize
            .Axes(xlValue, 2).HasMajorGridlines = AxisSettings.HasMajorGridlines
            .Axes(xlValue, 2).HasMinorGridlines = AxisSettings.HasMinorGridlines
            
        End If
            
        '.Legend.Font.Size = 20
        '.Legend.Delete
        
        .ChartArea.Border.LineStyle = Settings.BorderLineStyle
        
        .Axes(xlCategory).HasMajorGridlines = AxisSettings.HasMajorGridlines
        .Axes(xlCategory).HasMinorGridlines = AxisSettings.HasMinorGridlines
        
        .Axes(xlValue, 1).HasMajorGridlines = AxisSettings.HasMajorGridlines
        .Axes(xlValue, 1).HasMinorGridlines = AxisSettings.HasMinorGridlines
            
        'For Each srs In .SeriesCollection
        '    srs.Format.Line.Weight = 1
        'Next srs
        
        .ChartArea.Fill.Visible = Settings.ChartAreaBgVisible
        .PlotArea.Fill.Visible = Settings.PlotAreaBgVisible
        
        .Axes(xlCategory).TickLabels.Font.Color = AxisSettings.AxisColor
        .Axes(xlValue, 1).TickLabels.Font.Color = AxisSettings.AxisColor
        If IncludeSecondaryAxis Then .Axes(xlValue, 2).TickLabels.Font.Color = AxisSettings.AxisColor
            
    End With
    
    
    
    Set Chromatography_AddChart = cht
    
    Set cht = Nothing
    Set srs = Nothing

End Function

'****************************************************************************************************
Function ChartFormatSeries(TargetSeries As Excel.Series) As Excel.Chart

'====================================================================================================
'
'
'
'Juraj Ahel, 2014-06-08, for Master's thesis
'Last update 2016-05-17
'====================================================================================================

    Dim cht As Object
    Dim srs As Series
    
    AxisWeight = xlMedium '(Hairline, Thin, Medium, Thick)
    AxisColor = RGB(147, 149, 152) 'determined from Illustrator
    UV1Color = RGB(0, 0, 0)
    ConcBColor = AxisColor
    
    TitleSize = 25
    TickLabelSize = 25
    
    
    'if I want it in the sheet
    Set cht = ActiveSheet.ChartObjects.Add(Left:=0, Width:=800, Top:=0, Height:=500).Chart
    
    'ChartName = Mid(cht.Name, InStr(1, cht.Name, "Chart"), 1000)
    
    'if I want it as a separate object
    'Set cht = Charts.Add()
    
    With cht
        '.ChartTitle.text = "NiNTA"
        '.ChartTitle.Text = ""
        '.Type = xlXYScatter
        
        'introduce the series
        If RunType = 1 Then
            .ChartType = xlXYScatterSmoothNoMarkers
        End If
           
        With srs
            '.Name = "280 nm absorbance"
            .Format.Line.Weight = 2
            .Border.Color = UV1Color
            
        End With
        
        With .Axes(xlValue, 1)
            .HasTitle = True
            .MinimumScale = 0
            .AxisTitle.Font.Color = AxisColor
            With .AxisTitle
                .Caption = "A280"
                .Font.Size = TitleSize
                .Characters(1, 1).Font.Italic = True
                .Characters(2, 3).Font.Subscript = True
            End With
            .MinorUnit = .MajorUnit / 2
            .MajorTickMark = xlTickMarkOutside
            .MinorTickMark = xlTickMarkOutside
            .Border.Weight = AxisWeight 'before it was xlThick (=4)
            .Border.Color = AxisColor
            
        End With
        
        If RunType = 1 Then
            Set srs = ChartAddSeriesToChart(x1, y1, cht, , xlSecondary, False)
            With srs
                .Name = "% elution buffer"
                .Format.Line.Weight = xlHairline
                .Format.Line.BackColor.RGB = ConcBColor
                .Border.Color = ConcBColor
            End With
            
            With .Axes(xlValue, 2)
                .HasTitle = True
                .AxisTitle.Font.Color = AxisColor
                .MinimumScale = 0
                .MaximumScale = 100
                .AxisTitle.Caption = "% of elution buffer"
                .AxisTitle.Font.Size = TitleSize
                .Border.Weight = AxisWeight
                .Border.Color = AxisColor
                .MajorUnit = 20
                .MinorUnit = .MajorUnit / 2
                .MajorTickMark = xlTickMarkOutside
                .MinorTickMark = xlTickMarkOutside
            End With
        End If
               
        Select Case RunType
        Case 1 'NiNTA
            ScaleMax = ((x1(x1.Count, 1).Value \ 10) + 1) * 10
            If ScaleMax > 200 Then MajUN = 50 Else MajUN = 25
            MinUN = MajUN / 5
        Case 2 'S200 10/300
            MajUN = RoundToNearestX(Volume / 5, 1)
            MinUN = RoundToNearestX(MajUN / 5, 0.05)
            'ScaleMax = x1(x1.Count, 1).Value \ 1 + 1
            ScaleMax = ((x1(x1.Count, 1).Value \ 10) + 1) * 10
        Case 3 'Ettan SEC P3.2
            MajUN = 0.5
            MinUN = 0.1
            'ScaleMax = Round(x1(x1.Count, 1).Value, 1) + 0.1
            ScaleMax = ((x1(x1.Count, 1).Value \ 10) + 1) * 10
        End Select
        
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Font.Color = AxisColor
            .MajorUnit = MajUN
            .MinorUnit = MinUN
            .MajorTickMark = xlTickMarkCross
            .MinorTickMark = xlTickMarkOutside
            .MinimumScale = 0
            .MaximumScale = ScaleMax
            .AxisTitle.Caption = "Volume / mL"
            .AxisTitle.Font.Size = TitleSize
            .Border.Weight = AxisWeight
            .Border.Color = AxisColor
            If Volume < 10 Then
                .TickLabels.NumberFormat = "0,0"
            Else
                .TickLabels.NumberFormat = "0"
            End If
                
        End With
          
        .Axes(xlCategory).TickLabels.Font.Size = TickLabelSize
        .Axes(xlValue, 1).TickLabels.Font.Size = TickLabelSize
        If RunType = 1 Then
            .Axes(xlValue, 2).TickLabels.Font.Size = TickLabelSize
            
            .Axes(xlValue, 2).HasMajorGridlines = False
            .Axes(xlValue, 2).HasMinorGridlines = False
        End If
            
        '.Legend.Font.Size = 20
        .Legend.Delete
        .ChartArea.Border.LineStyle = xlNone
        
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlCategory).HasMinorGridlines = False
        
        .Axes(xlValue, 1).HasMajorGridlines = False
        .Axes(xlValue, 1).HasMinorGridlines = False
            
        'For Each srs In .SeriesCollection
        '    srs.Format.Line.Weight = 1
        'Next srs
        
        .ChartArea.Fill.Visible = False
        .PlotArea.Fill.Visible = False
        
        .Axes(xlCategory).TickLabels.Font.Color = AxisColor
        .Axes(xlValue, 1).TickLabels.Font.Color = AxisColor
        If RunType <= 1 Then .Axes(xlValue, 2).TickLabels.Font.Color = AxisColor
        
            
    End With

End Function
