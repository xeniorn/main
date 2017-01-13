VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChromatogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2016-05-06, for this and that purpose
'Last update 2016-05-24
'2016-06-13 add data alignment function to be able to prevent results like
'           0,012655    0,112655    0,312655    0,412655 ... (turn it into 0,0 0,1 0,2 0,3 ...)
'====================================================================================================
'2016-07-08 add some "ByVal"s
'2017-01-12 XToIndex change default value to 1, not 0

Option Explicit

Const conClassName As String = "clsChromatogram"

'[Var Declaration]
Private pChromatogramData As clsGraphData
Private pChromatogramData_Original As clsGraphData
Private pChromatogramData_UndoBuffer As clsGraphData

'[Object Initialization]
Private Sub Class_Initialize()

End Sub

'[Object Dereferencing]
Private Sub Class_Terminate()

    Set pChromatogramData = Nothing
    Set pChromatogramData_Original = Nothing
    Set pChromatogramData_UndoBuffer = Nothing

End Sub

Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub

'[Property Gets and Lets and Sets]
Public Property Get XAxisInfo() As AxisDataType
    XAxisInfo = ChromatogramData.XAxisInfo
End Property

Public Property Let XAxisInfo(inp_XAxis As AxisDataType)
    ChromatogramData.XAxisInfo = inp_XAxis
End Property

Public Property Get YAxisInfo() As AxisDataType
    YAxisInfo = ChromatogramData.YAxisInfo
End Property

Public Property Let YAxisInfo(inp_YAxis As AxisDataType)
    ChromatogramData.YAxisInfo = inp_YAxis
End Property

Public Property Get ChromatogramData() As clsGraphData
    
    Set ChromatogramData = pChromatogramData
       
End Property

Public Property Set ChromatogramData(inp_ChromatogramData As clsGraphData)

    Set pChromatogramData = inp_ChromatogramData
    
    If ChromatogramData_Original Is Nothing Then
        Set ChromatogramData_Original = New clsGraphData
        Call CopyChromatogramData(ChromatogramData, ChromatogramData_Original)
    End If
    
End Property

Private Property Get ChromatogramData_Original() As clsGraphData
    Set ChromatogramData_Original = pChromatogramData_Original
End Property

Private Property Set ChromatogramData_Original(inp_ChromatogramData_Original As clsGraphData)
    Set pChromatogramData_Original = inp_ChromatogramData_Original
End Property

Private Property Get ChromatogramData_UndoBuffer() As clsGraphData
    Set ChromatogramData_UndoBuffer = pChromatogramData_UndoBuffer
End Property

Private Property Set ChromatogramData_UndoBuffer(inp_ChromatogramData_UndoBuffer As clsGraphData)
    Set pChromatogramData_UndoBuffer = inp_ChromatogramData_UndoBuffer
End Property

Public Property Get Xmax() As Double
    Xmax = ChromatogramData.XData(UBound(ChromatogramData.XDataArray))
End Property

Public Property Get Ymax() As Double
    
    'this should call my external max thingie...
    
        Dim i As Long
        Dim tempmax As Double
        
        Dim IndexStart As Long, IndexEnd As Long
                
        With ChromatogramData
            
            IndexStart = LBound(.YDataArray)
            IndexEnd = UBound(.YDataArray)
                
            tempmax = .YData(IndexStart)
            
            For i = 2 To IndexEnd
                If .YData(i) > tempmax Then tempmax = .YData(i)
            Next i
            
        End With
        
        Ymax = tempmax
    
End Property

Public Property Get NumberOfPoints() As Long
    NumberOfPoints = UBound(ChromatogramData.XDataArray) - LBound(ChromatogramData.XDataArray) + 1
End Property

Public Property Get Xmin() As Double
    Xmin = ChromatogramData.XData(LBound(ChromatogramData.XDataArray))
End Property

Public Property Get Ymin() As Double
    
    'this should call my external min thingie...
    
    Dim i As Long
    Dim tempmin As Double
    
    Dim IndexStart As Long, IndexEnd As Long
            
    With ChromatogramData
        
        IndexStart = LBound(.YDataArray)
        IndexEnd = UBound(.YDataArray)
            
        tempmin = .YData(IndexStart)
        
        For i = 2 To IndexEnd
            If .YData(i) < tempmin Then tempmin = .YData(i)
        Next i
        
    End With
    
    Ymin = tempmin
    
End Property


'****************************************************************************************************
Public Sub ThinData(ByVal DistanceBetweenPoints As Double)
    
    Debug.Print ("Thinning chromatogram data")
    'it is biased towards the higher

    Dim i As Long, j As Long
    Dim tempYSum As Double
    Dim tempCount As Long
    
    Dim IndexStart As Long, IndexEnd As Long
    
    Dim XAnchorValue As Double
    Dim XShift As Double
    Dim XLowBound As Double, XHighBound As Double
    Dim XPlusMinus As Double
    
    Dim NewMaxIndex As Long
    
    
    Dim tempXData() As Double, tempYData() As Double
    
    IndexStart = LBound(ChromatogramData.XDataArray)
    IndexEnd = UBound(ChromatogramData.XDataArray)
    
    XPlusMinus = DistanceBetweenPoints / 2
    
    If Xmax > Xmin Then
        NewMaxIndex = Int((Xmax - Xmin) / DistanceBetweenPoints)
    Else
        Call ErrorReport(, "Max value smaller than min value")
    End If
    
    'if there is a significant number of point that wouldn't be included, increase MaxIndex by 1
    'i.e. more than 50 % of the new DistanceBetweenPoints would be truncated
    If ((Xmax - Xmin) - (NewMaxIndex - 1) * DistanceBetweenPoints) / DistanceBetweenPoints >= 0.5 Then
        NewMaxIndex = NewMaxIndex + 1
    End If
    
        
    
    ReDim tempXData(1 To NewMaxIndex)
    ReDim tempYData(1 To NewMaxIndex)
        
    i = IndexStart
    j = 1
        
    XAnchorValue = ChromatogramData.XData(IndexStart)
                
    XLowBound = XAnchorValue - XPlusMinus
    XHighBound = XAnchorValue + XPlusMinus
    tempYSum = 0
    tempCount = 0
    
    Do Until i = IndexEnd
    
        With ChromatogramData
            
            'if the data is well-ordered
            If .XData(i) >= XLowBound Then
                
                'sum all the data and increase count
                tempYSum = tempYSum + .YData(i)
                tempCount = tempCount + 1
                
                'until high bound is reached
                If .XData(i) >= XHighBound Then
                    
                    'then put the data in the calculated arrays
                    tempXData(j) = .XData(IndexStart) + (j - 1) * DistanceBetweenPoints
                    tempYData(j) = tempYSum / tempCount
                    
                    'set new high and low bounds
                    XLowBound = XAnchorValue - XPlusMinus + j * DistanceBetweenPoints
                    XHighBound = XAnchorValue + XPlusMinus + j * DistanceBetweenPoints
                    
                    'increase the counter of the calculated array, reset temp variables
                    j = j + 1
                    tempCount = 0
                    tempYSum = 0
                    
                End If
                            
            'if the data wasn't well ordered...
            Else
                
                'raise error, as the calculation assumes monotonously growing dataset in X!
                Call Err.Raise("1001", "clsChromatogram", "Source chromatogram data isn't monotonous in the XAxis")
                    
            End If
    
        End With
        
        i = i + 1
    
    Loop 'next index
    
    'add the last point, which might be partial-length - if it exists
    If tempCount <> 0 And j = NewMaxIndex Then
        tempXData(j) = XAnchorValue + (j - 1) * DistanceBetweenPoints
        tempYData(j) = tempYSum / tempCount
    End If
    
    Call CreateUndo
    
    ChromatogramData.XDataArray = tempXData
    ChromatogramData.YDataArray = tempYData
        
    

End Sub

'****************************************************************************************************
Public Function TotalIntegral()

    TotalIntegral = Integral(Xmin, Xmax)

End Function

Private Function XToIndex(XValue As Double)

    Dim i As Long
    
    XToIndex = 1
    
    For i = 1 To NumberOfPoints
        If ChromatogramData.XData(i) >= XValue Then
            XToIndex = i
            Exit For
        End If
    Next i


End Function

'****************************************************************************************************
Public Function Integral(Xstart As Double, Xend As Double)
'this is tricky to do right, but even this basic version should work in cases where
'the X-spacing is not excessively low, nor total integrals too large!

    Dim i As Long
    Dim tempIntegral As Double
    
    Dim PointStart As Long, PointEnd As Long
    
    PointStart = XToIndex(Xstart)
    PointEnd = XToIndex(Xend)
    
    
    For i = PointStart To PointEnd - 1
        With ChromatogramData
            tempIntegral = (.XData(i + 1) - .XData(i + 1)) * ((.YData(i) + .YData(i + 1)) / 2)
        End With
    Next i
    
    Integral = tempIntegral / (Xend - Xstart)
    
        
End Function

Public Function MaxInRange(Xstart As Double, Xend As Double)
    
    Dim PointStart As Long, PointEnd As Long
    
    Dim i As Long
    
    Dim tempmax As Double
    
    PointStart = XToIndex(Xstart)
    PointEnd = XToIndex(Xend)
            
    If PointStart > PointEnd Then
    
        Call ErrorReport(, "Maximum must be higher or equal to minimum in a range")
        
    Else
                
        With ChromatogramData
                       
            tempmax = .YData(PointStart)
            
            For i = PointStart + 1 To PointEnd
                If .YData(i) > tempmax Then tempmax = .YData(i)
            Next i
            
        End With
        
    MaxInRange = tempmax
        
    End If

End Function

'****************************************************************************************************
Public Sub NormalizeToMaxValue(Xstart As Double, Xend As Double)
    
    Dim tempmax As Double
    
    tempmax = Abs(MaxInRange(Xstart, Xend))
    
    If tempmax > 0 Then
        Call ScaleYAxis(1 / MaxInRange(Xstart, Xend))
    End If

End Sub

'****************************************************************************************************
Public Sub NormalizeToTotalIntegral()

    Call ScaleYAxis(1 / TotalIntegral())

End Sub

'****************************************************************************************************
Public Sub NormalizeToAreaUnderCurve(Xstart As Double, Xend As Double)

    Call ScaleYAxis(1 / Integral(Xstart, Xend))

End Sub

'****************************************************************************************************
Public Sub TruncateXAxisLeft(XLow As Double)

    Call TruncateXAxis(XLow, Xmax)

End Sub

'****************************************************************************************************
Public Sub TruncateXAxisRight(XHigh As Double)

    Call TruncateXAxis(Xmin, XHigh)

End Sub

'****************************************************************************************************
Public Sub AlignDataStartToZero(Optional ByVal ZeroValue As Double = 0)

    Call ShiftXAxis(ZeroValue - ChromatogramData.XData(1))

End Sub


'****************************************************************************************************
Public Sub TruncateXAxis(ByVal XLow As Double, Optional ByVal XHigh As Double)

    Dim i As Long
    Dim j As Long
            
    Dim PointStart As Long, PointEnd As Long
    
    Dim tempDataArrayX() As Double
    Dim tempDataArrayY() As Double
    
    
    CreateUndo
    
    PointStart = XToIndex(XLow)
    PointEnd = XToIndex(XHigh)
    
    'kind of a HACKY hack TODO:
    If PointEnd <= PointStart Then PointEnd = UBound(ChromatogramData.XDataArray)
    
    ReDim tempDataArrayX(1 To PointEnd - PointStart + 1)
    ReDim tempDataArrayY(1 To PointEnd - PointStart + 1)
    
    For i = PointStart To PointEnd
        tempDataArrayX(1 + i - PointStart) = ChromatogramData.XData(i)
        tempDataArrayY(1 + i - PointStart) = ChromatogramData.YData(i)
    Next i
    
    ChromatogramData.XDataArray = tempDataArrayX
    ChromatogramData.YDataArray = tempDataArrayY

End Sub

'****************************************************************************************************
Public Sub ScaleXAxis(ByVal ScaleMultiplier As Double)

    Dim i As Long
    Dim tempDataArray() As Double
    
    tempDataArray = ChromatogramData.XDataArray
    
    CreateUndo
    
    For i = 1 To NumberOfPoints
        tempDataArray(i) = tempDataArray(i) * ScaleMultiplier
    Next i

    ChromatogramData.XDataArray = tempDataArray

End Sub

'****************************************************************************************************
Public Sub ScaleYAxis(ByVal ScaleMultiplier As Double)

    Dim i As Long
    Dim tempDataArray() As Double
    
    tempDataArray = ChromatogramData.YDataArray
    
    CreateUndo
    
    For i = 1 To NumberOfPoints
        tempDataArray(i) = tempDataArray(i) * ScaleMultiplier
    Next i
    
    ChromatogramData.YDataArray = tempDataArray

End Sub


'****************************************************************************************************
Public Sub ShiftXAxis(ByVal ShiftAmount As Double)
    
    Dim i As Long
    Dim tempDataArray() As Double
    
    tempDataArray = ChromatogramData.XDataArray
    
    CreateUndo
    
    For i = 1 To NumberOfPoints
        tempDataArray(i) = tempDataArray(i) + ShiftAmount
    Next i
    
    ChromatogramData.XDataArray = tempDataArray

End Sub

'****************************************************************************************************
Public Sub OffsetYAxis(ByVal OffsetAmount As Double)
    
    Dim i As Long
    Dim tempDataArray() As Double
    
    tempDataArray = ChromatogramData.YDataArray
    
    CreateUndo
    
    For i = 1 To NumberOfPoints
        tempDataArray(i) = tempDataArray(i) + OffsetAmount
    Next i
    
    ChromatogramData.YDataArray = tempDataArray

End Sub

'****************************************************************************************************
Private Sub CopyChromatogramData(Source As clsGraphData, Destination As clsGraphData)

    If Source Is Nothing Then
    
        Call Err.Raise("1001", "clsChromatogram", "Source Chromatogram cannot be empty")
        
    Else
        
        If Destination Is Nothing Then Set Destination = New clsGraphData
        
        Destination.XDataArray = Source.XDataArray
        Destination.YDataArray = Source.YDataArray
        
        'HACKY, this isn't handled well. I need to make it so that AxisInfo can be reverted, but
        'the initial values are included into "original dataset", instead of the original containing
        'only "" empty strings, because it was created at the time of import of the data itself
        'Destination.XAxisInfo = Source.XAxisInfo
        'Destination.YAxisInfo = Source.YAxisInfo
        
    End If

End Sub

'****************************************************************************************************
Sub CommitChangesToOriginal()

    Call CopyChromatogramData(ChromatogramData, ChromatogramData_Original)
    
End Sub

'****************************************************************************************************
Sub RevertToOriginal()
    
    Call CopyChromatogramData(ChromatogramData, ChromatogramData_UndoBuffer)
    Call CopyChromatogramData(ChromatogramData_Original, ChromatogramData)
    
End Sub

'****************************************************************************************************
Sub UndoChange()

    Dim tempcGraphData As clsGraphData
    
    Call CopyChromatogramData(ChromatogramData, tempcGraphData)
    Call CopyChromatogramData(ChromatogramData_UndoBuffer, ChromatogramData)
    Call CopyChromatogramData(tempcGraphData, ChromatogramData_UndoBuffer)

End Sub

'****************************************************************************************************
Sub CreateUndo()

    'saves the current state of the ChromatogramData object to the Undo Buffer
    
    If ChromatogramData_UndoBuffer Is Nothing Then
        Set ChromatogramData_UndoBuffer = New clsGraphData
    End If
    
    Call CopyChromatogramData(ChromatogramData, ChromatogramData_UndoBuffer)
    
    'If ChromatogramData_Original Is Nothing Then Call CopyChromatogramData(ChromatogramData, ChromatogramData_Original)
    
End Sub
