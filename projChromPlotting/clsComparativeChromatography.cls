VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComparativeChromatography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2016-05-22, for this and that purpose
'Last update 2016-05-22
'====================================================================================================
Option Explicit

Implements IExperiment
Implements IChromExp

Private Const conClassName As String = "clsComparativeChromatography"


Const conFileTypeTSV As String = "TSV"
Const conFileTypeCSV As String = "CSV"

'[Var Declaration]
Private pChromatograms As clsChromatograms

'[Object Initialization]
Private Sub Class_Initialize()

    Set pChromatograms = New clsChromatograms

End Sub

'[Object Dereferencing]
Private Sub Class_Terminate()

    Set pChromatograms = Nothing
    
End Sub


Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub


'****************************************************************************************************
'****************************************************************************************************
Public Sub ImportTable(ByVal InputFilename As String, _
                    Optional ByVal FileType As String = conFileTypeTSV, _
                    Optional ByVal Delimiter As String = vbTab, _
                    Optional ByVal Arrangement As String = "X1Y1X2Y2X3Y3", _
                    Optional ByVal NumberOfHeaderLines As Long = 1, _
                    Optional ByVal MainHeaderLine As Long = 1)
                    
                    'or"XY1Y2Y3

    
'todo: change the DataToProcess to a type-safe collection!!!
           
    'iterators
    Dim i As Long, j As Long
    
    Dim DoubleDelimiter As String
    Dim NewlineDelimiter As String
    Dim DelimiterNewline As String
    Dim DoubleNewLine As String
        
    'temps
    Dim tempLine As String
    
    'Input
    Dim ResFile As String
    
    '
    Dim Lines() As String
    Dim tempLineDecomposition() As String
    Dim Headers() As String
    Dim ArrayOfGraphData() As clsGraphData
    Dim tempChromatogram As clsChromatogram
    Dim TempGraphArray() As Double
    
    Dim NumberOfDataLines As Long
    Dim NumberOfColumns As Long
    Dim NumberOfDataSets As Long
    
    Dim tempValue As Double, tempLastValue As Double
    Dim XDataConst As Double
    Dim tempAxisInfo As AxisDataType
    
    DoubleDelimiter = Delimiter & Delimiter
    NewlineDelimiter = vbCrLf & Delimiter
    DelimiterNewline = Delimiter & vbCrLf
    DoubleNewLine = vbCrLf & vbCrLf
    
    'Load the input file
    ResFile = ReadTextFile(InputFilename)
    
    'replace empty values with zeroes (adjacent delimiters, delimiters "touching" newline
        Do While StringCharCountLong(ResFile, DoubleDelimiter) > 0
            ResFile = Replace(ResFile, DoubleDelimiter, Delimiter & "0" & Delimiter)
        Loop
        
        Do While StringCharCountLong(ResFile, DelimiterNewline) > 0
            ResFile = Replace(ResFile, DelimiterNewline, Delimiter & "0" & vbCrLf)
        Loop
        
        Do While StringCharCountLong(ResFile, NewlineDelimiter) > 0
            ResFile = Replace(ResFile, NewlineDelimiter, vbCrLf & "0" & Delimiter)
        Loop
    
    'eliminate empty rows
        Do While StringCharCountLong(ResFile, DoubleNewLine) > 0
            ResFile = Replace(ResFile, DoubleNewLine, vbCrLf)
        Loop
    
    'eliminate terminal vbCrLf
        Do While Right(ResFile, 1) = vbCr Or Right(ResFile, 1) = vbLf
            ResFile = Left(ResFile, Len(ResFile) - 1)
        Loop
    
    Lines = Split(ResFile, vbCrLf)
    
    NumberOfDataLines = 1 + UBound(Lines) - LBound(Lines) - NumberOfHeaderLines
    
    'read headers
    For i = 0 To NumberOfHeaderLines - 1
        If i = MainHeaderLine - 1 Then
            Headers = Split(Lines(i), Delimiter)
            NumberOfColumns = UBound(Headers) - LBound(Headers) + 1
        End If
    Next i
    
    'Define number of final datasets
    Select Case Arrangement
        Case "XY1Y2Y3"
            NumberOfDataSets = NumberOfColumns - 1
        Case "X1Y1X2Y2X3Y3"
            NumberOfDataSets = NumberOfColumns / 2
    End Select
    
    ReDim ArrayOfGraphData(0 To NumberOfDataSets - 1)
    ReDim TempGraphArray(1 To NumberOfDataLines)
    
    'Initialize Graph collections + set axis labels
    For i = 0 To NumberOfDataSets - 1
    
        Set ArrayOfGraphData(i) = New clsGraphData
        ArrayOfGraphData(i).XDataArray = TempGraphArray
        ArrayOfGraphData(i).YDataArray = TempGraphArray
        
        Select Case Arrangement
            Case "XY1Y2Y3"
                tempAxisInfo.Label = Headers(0)
                ArrayOfGraphData(i).XAxisInfo = tempAxisInfo
                tempAxisInfo.Label = Headers(1 + i)
                ArrayOfGraphData(i).YAxisInfo = tempAxisInfo
             
            Case "X1Y1X2Y2X3Y3"
                tempAxisInfo.Label = Headers(2 * i)
                ArrayOfGraphData(i).XAxisInfo = tempAxisInfo
                tempAxisInfo.Label = Headers(2 * i + 1)
                ArrayOfGraphData(i).YAxisInfo = tempAxisInfo
        End Select
        
    Next i
    
    'Read the data
    For i = NumberOfHeaderLines To UBound(Lines)
        
        tempLine = Lines(i)
        
        tempLineDecomposition = Split(tempLine, Delimiter)
        
        Select Case Arrangement
            Case "XY1Y2Y3"
                XDataConst = CDbl(tempLineDecomposition(0))
                For j = 0 To NumberOfDataSets - 1
                    ArrayOfGraphData(j).XData(1 + i - NumberOfHeaderLines) = XDataConst
                    ArrayOfGraphData(j).YData(1 + i - NumberOfHeaderLines) = CDbl(tempLineDecomposition(1 + j))
                Next j
            
            Case "X1Y1X2Y2X3Y3"
                For j = 0 To NumberOfDataSets - 1
                    If CDbl(tempLineDecomposition(2 * j)) > tempLastValue Then
                        tempValue = CDbl(tempLineDecomposition(2 * j))
                        tempLastValue = tempValue
                    Else
                        tempValue = tempLastValue
                    End If
                    
                    ArrayOfGraphData(j).XData(1 + i - NumberOfHeaderLines) = tempValue
                    ArrayOfGraphData(j).YData(1 + i - NumberOfHeaderLines) = CDbl(tempLineDecomposition(2 * j + 1))
                    
                Next j
                
        End Select
        
    Next i
    
    For i = 0 To NumberOfDataSets - 1
    
        Set tempChromatogram = New clsChromatogram
        Set tempChromatogram.ChromatogramData = ArrayOfGraphData(i)
        Set ArrayOfGraphData(i) = Nothing
        
        Call AddChromatogram(tempChromatogram, CStr(i))
    
    Next i
    
    Set tempChromatogram = Nothing

End Sub

Public Sub NormalizeAll(Xstart As Double, Xend As Double)
    
    Dim i As Long
    Dim tempChrom As clsChromatogram
    
    For i = 1 To Chromatograms.Count
        
        Set tempChrom = Chromatograms.Item(i)
        
        If Not (tempChrom Is Nothing) Then
            Call tempChrom.NormalizeToMaxValue(Xstart, Xend)
        End If
        
    Next i

    Set tempChrom = Nothing
    
End Sub

Public Sub ThinData(ByVal DistanceBetweenPoints As Double)

    Dim i As Long
    Dim tempChrom As clsChromatogram
    
    For i = 1 To Chromatograms.Count
        
        Set tempChrom = Chromatograms.Item(i)
        
        If Not (tempChrom Is Nothing) Then
            tempChrom.ThinData (DistanceBetweenPoints)
        End If
        
    Next i

    Set tempChrom = Nothing

End Sub

Private Function GenerateRangeArray(SourceArray() As Double) As Double()

    Dim tempArray() As Double
    Dim i As Long
    
    ReDim tempArray(LBound(SourceArray) To UBound(SourceArray), 1 To 1)
    
    For i = LBound(SourceArray) To UBound(SourceArray)
        tempArray(i, 1) = SourceArray(i)
    Next i
    
    GenerateRangeArray = tempArray

End Function


Public Sub tempDraw()

    Dim tempChart As Excel.ChartObject
    Dim TempSeries As Excel.Series
    Dim tempChromatogramData As clsGraphData
    Dim XRange As Excel.Range
    Dim YRange As Excel.Range
    Dim LabelRange As Excel.Range
    
    Dim i As Long

    Set tempChart = Chromatography_AddChart
    
    tempChart.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = Chromatograms.Item(1).XAxisInfo.Label
    
    For i = 1 To Chromatograms.Count
        Set tempChromatogramData = Chromatograms.Item(i).ChromatogramData
        With tempChromatogramData
        
            Set XRange = Range("A3").Offset(0, 2 * (i - 1)).Resize(.Count, 1)
            Set YRange = Range("A3").Offset(0, 2 * (i - 1) + 1).Resize(.Count, 1)
            Set LabelRange = Range("A2").Offset(0, 2 * (i - 1)).Resize(1, 1)
            
            XRange.Value = GenerateRangeArray(.XDataArray)
            YRange.Value = GenerateRangeArray(.YDataArray)
            LabelRange.Value = .YAxisLabel
            
            Set TempSeries = ChartAddSeriesToChart(XRange, YRange, tempChart, LabelRange)
            
            TempSeries.Format.Line.Weight = xlHairline
            
            
        End With
        'tempSeries.XValues = tempChromatogramData.XDataArray
        'tempSeries.Values = tempChromatogramData.YDataArray
    Next i
    
    Set tempChart = Nothing
    Set TempSeries = Nothing
    Set tempChromatogramData = Nothing
    

End Sub

Public Sub tempDrawDirect()

    Dim tempChart As Excel.ChartObject
    Dim TempSeries As Excel.Series
    Dim tempChromatogramData As clsGraphData
    
    Dim i As Long

    Set tempChart = Chromatography_AddChart
    
    tempChart.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = Chromatograms.Item(1).XAxisInfo.Label
    
    For i = 1 To Chromatograms.Count
        Set tempChromatogramData = Chromatograms.Item(i).ChromatogramData
        With tempChromatogramData
            Set TempSeries = ChartAddSeriesToChartDirect(.XDataArray, .YDataArray, tempChart, .YAxisLabel)
        End With
        'tempSeries.XValues = tempChromatogramData.XDataArray
        'tempSeries.Values = tempChromatogramData.YDataArray
    Next i
    
    Set tempChart = Nothing
    Set TempSeries = Nothing
    Set tempChromatogramData = Nothing
    

End Sub


Public Property Get Chromatograms() As clsChromatograms
    Set Chromatograms = pChromatograms
End Property


Public Sub AddChromatogram(ByVal InpChromatogram As clsChromatogram, ChromatogramID As String)

    Call Chromatograms.Add(InpChromatogram, ChromatogramID)

End Sub

Private Sub IChromExp_AddChromatogram(ByVal InpChromatogram As clsChromatogram)

    AddChromatogam (InpChromatogram)

End Sub

Public Property Get ExperimentDate() As String

End Property

Public Property Get ExperimentScientist() As String

End Property


Public Property Get ExperimentType() As String

End Property


Private Function IChromExp_GetChromatogram(Optional ByVal Index As Variant = 1) As clsChromatogram

End Function

Private Property Get IExperiment_ExpDate() As String
    IExperiment_ExpDate = ExperimentDate
End Property

Private Property Get IExperiment_ExpName() As String
    IExperiment_ExpName = ExperimentName
End Property

Private Property Get IExperiment_ExpType() As String
    IExperiment_ExpType = ExperimentType
End Property

Private Property Get IExperiment_Scientist() As String
    IExperiment_ExpType = ExperimentScientist
End Property

