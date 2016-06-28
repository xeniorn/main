Attribute VB_Name = "XmodFoldIndex"

                        

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

'****************************************************************************************************
Function FoldIndex( _
                    InputSequence As String, _
                    Optional WindowSize As Integer = 51, _
                    Optional Separator As String = vbTab _
                    ) As String

'====================================================================================================
'Calculates the order-disorder factor in the way FoldIndex web app does
'
'Juraj Ahel, 2015-04-23, for easier drawing from FoldIndex
'Last update 2015-04-23
'====================================================================================================

Dim Hydropathy As Double, Charge As Double
Dim SequenceLength As Long, SubstringLength As Long
Dim SubString As String, CurrentResidue As String
Dim i As Long, z As Long, j As Long
Dim TempResult()

Dim AminoAcids As String: AminoAcids = "ACDEFGHIKLMNPQRSTVWY"
Dim HydropathyIndex()

HydropathyIndex = Array(0.7, 0.778, 0.111, 0.111, 0.811, 0.456, 0.144, 1#, 0.067, 0.922, 0.711, 0.111, 0.333, 0.111, 0#, 0.411, 0.578, 0.967, 0.4, 0.356)

SequenceLength = Len(InputSequence)

'Just a basic check for meaningfulness of window
If WindowSize > SequenceLength Then WindowSize = SequenceLength
If WindowSize = 0 Then WindowSize = 1

IndexStart = Int(WindowSize / 2) + 1
IndexEnd = SequenceLength - Int((WindowSize - 1) / 2)

ReDim TempResult(1 To SequenceLength)


For i = 1 To SequenceLength - WindowSize + 1

    SubString = Mid(InputSequence, i, WindowSize)
    SubstringLength = Len(SubString)
    z = i + IndexStart - 1
    
    Hydropathy = 0
    Charge = 0
    
    For j = 1 To 20
        Hydropathy = Hydropathy + _
        HydropathyIndex(j - 1) * StringCharCount(SubString, Mid(AminoAcids, j, 1))
    Next j
           
    Charge = StringCharCount(SubString, "D", "E") - StringCharCount(SubString, "K", "R")
    
    'FoldIndex formula
    TempResult(z) = Round(((2.785 * Hydropathy - Abs(Charge)) / WindowSize - 1.151), 4)
    
Next i

For i = 1 To (IndexStart - 1): TempResult(i) = 0: Next i
For i = (IndexEnd + 1) To SequenceLength: TempResult(i) = 0: Next i

'FoldIndex = Join(TempResult, Chr(13) & Chr(10))
FoldIndex = Join(TempResult, Separator)
'Call ExportToTXT FoldIndex
End Function

'****************************************************************************************************
Sub FoldIndexDraw(WindowSize As Long, PlotRange As Range, _
                    LeftOffset, GraphWidth, _
                    TopOffset, GraphHeight, _
                    GraphMaximum, GraphMinimum, _
                    TickSpace, LabelSpace, DisplayGrid, _
                    Mode)

'====================================================================================================
'Draws the graphs for FoldIndexMacro
'Juraj Ahel, 2015-04-24, for more automated FoldIndex-ing
'Last update 2015-04-25
'====================================================================================================

Dim myChart As Object
Dim srs As Series

'if I want it in the sheet
'If mode = 0 Then Set myChart = ActiveSheet.ChartObjects.Add(Left:=0, Width:=800, Top:=0, Height:=500).Chart

If Mode = 1 Then SeriesNumber = 1
If Mode = 2 Then SeriesNumber = 2

Set myChart = ActiveSheet.ChartObjects.Add(Left:=LeftOffset, Width:=GraphWidth, _
                                            Top:=TopOffset, Height:=GraphHeight) _
                                            .Chart

'How big the labels and markers on axes will be
TitleSize = 25
TickLabelSize = 25

Dim ChartColor(1 To 2)
Dim Data(1 To 2) As Range

'Green and Red, respectively for positive and negative series from FoldIndexMacro
ChartColor(1) = RGB(25, 190, 25)
ChartColor(2) = RGB(200, 25, 25)

If Mode = 1 Then ChartColor(1) = 13998939 '-that bluish color

Set Data(1) = PlotRange
Set Data(2) = PlotRange.Offset(0, 1)

With myChart
    '.ChartTitle.text = "NiNTA"
    .HasTitle = True
    .ChartTitle.Text = CStr(WindowSize)
    '.Type = xlXYScatter
         
    'remove possible old series
    For Each srs In .SeriesCollection
        srs.Delete
    Next srs
    
    .ChartType = xlArea
        
    For i = 1 To SeriesNumber
        'introduce the series
        .SeriesCollection.NewSeries
        
           
        With .SeriesCollection(i)
            .Values = Data(i)
            .Format.Fill.ForeColor.RGB = ChartColor(i)
        End With
    Next i
    
    If DisplayGrid Then
        With .Axes(xlValue, 1)
            .HasTitle = True
            .MinimumScale = GraphMinimum
            .MaximumScale = GraphMaximum
            With .AxisTitle
                .Caption = "Fold Index"
                .Font.Size = TitleSize
            End With
            .MajorTickMark = xlTickMarkOutside
            .MinorTickMark = xlTickMarkOutside
            .Border.Weight = xlThick
            .Border.Color = RGB(0, 0, 0)
        End With
        Else
        With .Axes(xlValue, 1)
            .MinimumScale = GraphMinimum
            .MaximumScale = GraphMaximum
        End With
    End If
    
    If DisplayGrid Then
        With .Axes(xlPrimary)
            .HasTitle = True
            '.MinimumScale = 1
            '.MaximumScale = PlotRange.Rows.Count
            With .AxisTitle
                .Caption = "residue number"
                .Font.Size = TitleSize
            End With
            '.MinorUnit = .MajorUnit / 2
            .MajorTickMark = xlTickMarkOutside
            .MinorTickMark = xlTickMarkOutside
            .Border.Weight = xlThick
            .Border.Color = RGB(0, 0, 0)
            .TickMarkSpacing = TickSpace
            .TickLabelSpacing = LabelSpace
        End With
    Else
        With .Axes(xlPrimary)
            .MajorTickMark = xlNone
            .MinorTickMark = xlNone
            .Border.Weight = xlThin
            .Border.Color = RGB(0, 0, 0)
            .TickLabelPosition = xlTickLabelPositionNone
        End With
    End If
    
                    
    .Axes(xlCategory).TickLabels.Font.Size = TickLabelSize
    .Axes(xlValue, 1).TickLabels.Font.Size = TickLabelSize
    .Axes(xlValue).MajorGridlines.Delete
    '.Axes(xlValue).MinorGridlines.Delete
            
    '.Legend.Font.Size = 20
    .Legend.Delete
    .ChartArea.Border.LineStyle = xlNone
        
    'For Each srs In .SeriesCollection
    '    srs.Format.Line.Weight = 1
    'Next srs
        
    If Not DisplayGrid Then
        '.HasAxis(xlPrimary) = False
        .HasAxis(xlValue) = False
    End If
    
End With


End Sub

'****************************************************************************************************
Sub FoldIndexMacro()

'====================================================================================================
'Performs the FoldIndex calculation and generates the graphs to be imported in photoshop for overlaying
'Plots positive and negative values separately (different colors!)
'All the graphs have the same min / max x and y axes, so should be easy to overlay!
'The idea is to export the images using Daniel's XL Toolbox, and import them to Photoshop
'and overlaying them, with blend mode "Multiply" and then finetuning opacity to get optimal saturation
'Juraj Ahel, 2015-04-24
'Last update 2015-04-25
'====================================================================================================

Dim SeparatePositiveAndNegativeByColor As Boolean

SeparatePositiveAndNegativeByColor = True

Dim InputCell As Object, OutputRange As Range
Dim OutputTable() As Double
Dim WindowSizeList(), ScaleList()

Dim GraphNumber As Integer
Dim i As Long, SequenceLength As Long

Dim InputSequence As String, TempResult As String
Dim FoldIndexValues() As String

Dim MaxWindow As Long, MinWindow As Long, NumberOfWindows As Long
Dim SeppFactor As Double

MinWindow = 50
MaxWindow = 250
NumberOfWindows = 10

ReDim WindowSizeList(0 To NumberOfWindows - 1)

'Classic windows size list (first successful Mys1a overlay):
'WindowSizeList = Array(5, 25, 51, 75, 101, 151, 201)

'Generate equally log-spaced windows:
If NumberOfWindows > 1 Then
    SeppFactor = Log(MaxWindow / MinWindow) / (NumberOfWindows - 1)
End If

For i = 0 To NumberOfWindows - 1
    WindowSizeList(i) = Round(Exp(Log(MinWindow) + SeppFactor * i), 0)
Next i

'Essentially equal to NumberOfWindows
GraphNumber = UBound(WindowSizeList) - LBound(WindowSizeList) + 1

ReDim ScaleList(0 To GraphNumber - 1)

'Set scales proportional to window width
For i = 1 To GraphNumber
    ScaleList(GraphNumber - i) = WindowSizeList(GraphNumber - i) / WindowSizeList(GraphNumber - 1)
Next i

Set InputCell = Selection
'Set InputCell = Application.InputBox("Select cell containing input sequence:","Input selection",Type:=8)
                                    
InputSequence = CStr(InputCell.Value)
SequenceLength = Len(InputSequence)

ReDim OutputTable(1 To SequenceLength + 1, 1 To 1 + 2 * GraphNumber)

'First column is used just to generate the last graph (the axes without the profile)
'For i = 1 To SequenceLength: OutputTable(i, 1) = i: Next i
For i = 1 To SequenceLength: OutputTable(i, 1) = 0: Next i

'Other columns are scaled FoldIndex profiles
For i = 1 To GraphNumber
    
    TempResult = FoldIndex(InputSequence, CLng(WindowSizeList(i - 1)), vbTab)
    FoldIndexValues = Split(TempResult, vbTab)
    
    OutputTable(1, 2 * i) = WindowSizeList(i - 1)
    OutputTable(1, 2 * i + 1) = WindowSizeList(i - 1)
    For j = 1 To SequenceLength
        TempNumber = CDbl(FoldIndexValues(j - 1)) * CDbl(ScaleList(i - 1))
        
        OutputTable(j + 1, 2 * i) = TempNumber
        OutputTable(j + 1, 2 * i + 1) = TempNumber
        'Positive and negative ones on separate series - data in separate columns!
        If SeparatePositiveAndNegativeByColor And TempNumber < 0 Then
            OutputTable(j + 1, 2 * i) = 0
        Else
            OutputTable(j + 1, 2 * i + 1) = 0
        End If
        
    Next j
    
    
Next i
    
Set OutputRange = InputCell.Offset(1, 0).Resize(SequenceLength, 1 + 2 * GraphNumber)
OutputRange.Value = OutputTable

Dim PlotRange As Range
Dim GraphMaximum As Double, GraphMinimum As Double
Dim DrawMode As Integer


'To get meaningfully scaled visuals, graphs are drawn between 110 % global minimum and 110 % global maximum
GraphMaximum = WorksheetFunction.Max(OutputRange.Offset(1, 1).Resize(SequenceLength, 2 * GraphNumber))
GraphMaximum = RoundToNearestX(1.1 * GraphMaximum, 0.01)
GraphMinimum = WorksheetFunction.Min(OutputRange.Offset(1, 1).Resize(SequenceLength, 2 * GraphNumber))
GraphMinimum = RoundToNearestX(1.1 * GraphMinimum, 0.01)

'The graphs are drawn one below the other using a separate drawing Sub
'graphs are drawn without axes for ease of overlaying, axes are available as a separate graph

For i = 1 To GraphNumber
    
    Set PlotRange = OutputRange.Offset(1, 2 * i - 1).Resize(SequenceLength, 1)
    
    'Where the graph is drawn and how big it is'
    LeftOffset = 0
    GraphWidth = 2000
    TopOffset = 0 + 275 * (i - 1)
    GraphHeight = 250
    
    'Spacing between major markers and labels. There is a minor marker between 2 major ones
    TickSpace = 500
    LabelSpace = 500
    
    If SeparatePositiveAndNegativeByColor Then DrawMode = 2 Else DrawMode = 1
    Call FoldIndexDraw(CLng(WindowSizeList(i - 1)), _
                        PlotRange, _
                        LeftOffset, GraphWidth, TopOffset, GraphHeight, _
                        GraphMaximum, GraphMinimum, _
                        TickSpace, LabelSpace, DisplayGrid:=False, Mode:=DrawMode)
Next i

'In the end, also draw a separate plot with the axes

Set PlotRange = OutputRange.Offset(1, 0).Resize(SequenceLength, 1)
    
    LeftOffset = 0
    GraphWidth = 2000
    TopOffset = 0 + 275 * GraphNumber
    GraphHeight = 250
    
    TickSpace = 500
    LabelSpace = 500
    
    Call FoldIndexDraw(0, _
                        PlotRange, _
                        LeftOffset, GraphWidth, TopOffset, GraphHeight, _
                        GraphMaximum, GraphMinimum, _
                        TickSpace, LabelSpace, DisplayGrid:=True, Mode:=1)


End Sub

