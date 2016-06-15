VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUnicorn3Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2016-05-08, for importing unicorn .res files
'Juraj Ahel, 2016-05-18
'2016-06-13 change injection annotation to "1,2,3, ... N" instead of "N, N, N, ...", where N is N of injections
'====================================================================================================
'TODO: move all import commands to Unicorn3Importer, make UnicornData just carry the data (as is done for Unicorn6Importer)

Option Explicit

Private Const conDefaultWavelength As Double = 280#

Private Const conClassName As String = "clsUnicorn3Data"

'Elution volumes are stored as 100*V/mL, need to be divided by 100
Private Const conVolumeScaling As Double = 100#

Private pLabel As String
Private pPointNumber As Long

Private pHeader As clsUnicorn3Header
Private pGraph As clsGraphData
Private pAnnotation As clsGraphAnnotation
Private pRegions As clsGraphRegions

Private pXAxisType As AxisDataType
Private pYAxisType As AxisDataType

Private pLogBook As String

Private pUVWavelength As Double

Private Sub Class_Initialize()

    pUVWavelength = conDefaultWavelength

End Sub

Private Sub Class_Terminate()
    
    Set pAnnotation = Nothing
    Set pGraph = Nothing
    Set pHeader = Nothing
    Set pRegions = Nothing

End Sub

Private Sub ErrorReport(Optional ErrorNumber As Long = 0, Optional ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub

Public Property Get Name() As String

    Name = pHeader.Name

End Property

Public Property Get ChromName() As String

    ChromName = pHeader.ChromName

End Property

Public Property Get UVWavelength() As Double
    UVWavelength = pUVWavelength
End Property

Public Property Get XAxisType() As AxisDataType
    XAxisType = pXAxisType
End Property

Public Property Get YAxisType() As AxisDataType
    YAxisType = pYAxisType
End Property


Private Property Let PointNumber(N As Long)
    pPointNumber = N
End Property

Public Property Get PointNumber() As Long
    PointNumber = pPointNumber
End Property


Public Property Get CurveType() As eUnicornCurveType
    CurveType = pHeader.CurveType
End Property

Public Property Get DataType() As eUnicornDataType
    DataType = pHeader.DataType
End Property

Private Property Get Annotation() As clsGraphAnnotation
    Set Annotation = pAnnotation
End Property

Private Property Set Annotation(GraphInp As clsGraphAnnotation)
    Set pAnnotation = GraphInp
End Property


Private Property Set Graph(GraphInp As clsGraphData)
    Set pGraph = GraphInp
End Property


Public Property Get Graph() As clsGraphData
    Set Graph = pGraph
End Property

Private Property Get AnnotatedRegions() As clsGraphRegions
    Set AnnotatedRegions = pRegions
End Property

Private Property Set AnnotatedRegions(GraphInp As clsGraphRegions)
    Set pRegions = GraphInp
End Property


Public Sub ImportHeader(HI As clsUnicorn3Header)
'imports a defined header object
    Set pHeader = HI
End Sub

Public Sub ImportFromRaw(ByVal FullRawString As String)
    
    Dim DataBlockString As String
    Dim DataLength As Long
    
    DataBlockString = Mid(FullRawString, pHeader.DataAddress + 1, pHeader.DataSize)
    DataLength = Len(DataBlockString)
    
    
    Select Case pHeader.DataType
    
        Case uniData
        
            If DataLength <= conDataMeta_L Then
                Call ErrorReport(, "The input source is not long enough - it is not possible to extract any data")
            Else
                Call ImportDataMeta(Left(DataBlockString, conDataMeta_L))
                Call ImportData(Right(DataBlockString, DataLength - conDataMeta_L))
            End If
                
        Case uniInjection
        
            'Injection, Fraction, and LogBook have same length of meta (Meta1)
            If DataLength <= conMeta1Meta_L Then
                Call ErrorReport(, "The input source is not long enough - it is not possible to extract any data")
            Else
                Call ImportInjectionMeta(Left(DataBlockString, conMeta1Meta_L))
                Call ImportInjection(Right(DataBlockString, DataLength - conMeta1Meta_L))
            End If
                
        Case uniFraction
        
            'Injection, Fraction, and LogBook have same length of meta (Meta1)
            If DataLength <= conMeta1Meta_L Then
                Call ErrorReport(, "The input source is not long enough - it is not possible to extract any data")
            Else
                Call ImportFractionMeta(Left(DataBlockString, conMeta1Meta_L))
                Call ImportFraction(Right(DataBlockString, DataLength - conMeta1Meta_L))
            End If
            
        Case uniLogBook
        
            'Injection, Fraction, and LogBook have same length of meta (Meta1)
            If DataLength <= conMeta1Meta_L Then
                Call ErrorReport(, "The input source is not long enough - it is not possible to extract any data")
            Else
                Call ImportLogBookMeta(Left(DataBlockString, conMeta1Meta_L))
                Call ImportLogBook(Right(DataBlockString, DataLength - conMeta1Meta_L))
            End If
            
        Case uniMeta
        
            'entire string is the meta
            Call ImportMeta(DataBlockString)
        
        Case Else
            Call ErrorReport(, "Header does not have a valid data type")
            
    End Select
            

    
End Sub

'====================================================================================================
'META IMPORT
'====================================================================================================
Private Sub ImportDataMeta(ByVal MetaString As String)
 
    Const MagicID_First_Pos     As String = 1
    Const MagicID_First_L       As String = 10
    
    Const MagicID_Second_Pos    As String = 77
    Const MagicID_Second_L      As String = 8
    
    Const MagicID_Third_Pos     As String = 225
    Const MagicID_Third_L       As String = 16
    
    Const TimeUnit_Pos          As String = 51
    Const TimeUnit_L            As String = 16
    
    Const VolumeUnit_Pos        As String = 129
    Const VolumeUnit_L          As String = 16
    
    Const YAxisUnit_Pos         As String = 208
    Const YAxisUnit_L           As String = 17
    
    Const DataStreamName_Pos    As String = 167
    Const DataStreamName_L      As String = 40
    
     
    Dim tempMagicID_First       As String
    Dim tempMagicID_Second      As String
    Dim tempMagicID_Third       As String
    
    Dim tempTimeUnit            As String
    Dim tempVolumeUnit          As String
    Dim tempYAxisUnit           As String
    
    Dim tempDataStreamName      As String
    
    'many of them not used at the moment, might use them later
    
    tempMagicID_First = Mid(MetaString, MagicID_First_Pos, MagicID_First_L)
    tempMagicID_Second = Mid(MetaString, MagicID_Second_Pos, MagicID_Second_L)
    tempMagicID_Third = Mid(MetaString, MagicID_Third_Pos, MagicID_Third_L)
    
    tempTimeUnit = Mid(MetaString, TimeUnit_Pos, TimeUnit_L)
    tempVolumeUnit = Mid(MetaString, VolumeUnit_Pos, VolumeUnit_L)
    tempYAxisUnit = Mid(MetaString, YAxisUnit_Pos, YAxisUnit_L)
    
    tempDataStreamName = Mid(MetaString, DataStreamName_Pos, DataStreamName_L)
    
    'extract data:
    
    pXAxisType.Unit = RTrim$(StringRemoveNonPrintable(tempVolumeUnit))
    pYAxisType.Unit = RTrim$(StringRemoveNonPrintable(tempYAxisUnit))
    
    Select Case pXAxisType.Unit
        Case "ml": pXAxisType.Unit = "mL"
        Case "l": pXAxisType.Unit = "L"
        Case "ul": pXAxisType.Unit = "uL"
        Case "nl": pXAxisType.Unit = "nL"
    End Select
 
End Sub

Private Sub ImportInjectionMeta(ByVal MetaString As String)

End Sub

Private Sub ImportFractionMeta(ByVal MetaString As String)

End Sub

Private Sub ImportLogBookMeta(ByVal MetaString As String)

End Sub

'====================================================================================================
'DATA IMPORT
'====================================================================================================
Private Sub ImportData(ByVal DataString As String)

'imports the binary data from UNICORN 3.1 style datablock (int32-int32 volume-signal)

    Dim XData() As Double, YData() As Double
        
    Dim i As Long
    
    Dim tempData As String, tempVolume As Long, tempSignal As Long
    
    Dim ScalingFactor As Double
    
    'variables are scaled in some way, consistently within type
    Select Case CurveType
        Case uniUV, uniCond, uniFlow
            ScalingFactor = 1000#
        Case uniPressure
            ScalingFactor = 100#
        Case uniConc, uniOther
            ScalingFactor = 10#
    End Select
    
    'data must obviously be a multiple of 8 bytes
    If Len(DataString) Mod conDataPoint_L <> 0 Then
        
        Call ErrorReport(, "Data String is not a multiple of 8 bytes - check input")
    
    Else
    
        'number of data points
        PointNumber = Len(DataString) \ conDataPoint_L
        
        'Debug.Print ("PointNumber: " & PointNumber)
        
        ReDim XData(1 To PointNumber)
        ReDim YData(1 To PointNumber)
        
        For i = 1 To PointNumber
            
            'get a small sub buffer out, and separate it into components
            tempData = Mid(DataString, 1 + (i - 1) * conDataPoint_L, conDataPoint_L)
            tempVolume = BinaryStringToInt32(Mid(tempData, conDataVolume_IntPos, conDataVolume_L), LittleEndian:=False)
            tempSignal = BinaryStringToInt32(Mid(tempData, conDataSignal_IntPos, conDataSignal_L), LittleEndian:=False)
            
            'variables are scaled in some way
            XData(i) = CDbl(tempVolume) / conVolumeScaling
            YData(i) = CDbl(tempSignal) / ScalingFactor
                                    
        Next i
        
        Set Graph = New clsGraphData
        
        pGraph.XDataArray = XData
        pGraph.YDataArray = YData
        
        'Call tempexport(pGraph.XData, pGraph.YData)
    
    End If

End Sub

Private Function ImportAnnotation(ByVal DataString As String) As clsGraphAnnotation
                            
                            
'imports the binary data from UNICORN 3.1 style datablock (int32-int32 volume-signal)

    Dim i As Long
    
    Dim tempData As String
    Dim tempVolume As Double, tempTime As Double
    Dim tempAnnotation As String
    Dim tempcAnnotation As clsGraphAnnotation
    
    Dim tempStringArray() As String
    Dim tempVolumeArray() As Double
    
       
    'data must obviously be a multiple of conMeta1Point_L bytes
    If Len(DataString) Mod conMeta1Point_L <> 0 Then
        
        Call ErrorReport(, "Data String is not a multiple of 8 bytes - check input")
    
    Else
    
        'number of data points
        PointNumber = Len(DataString) \ conMeta1Point_L
        
        'Debug.Print ("PointNumber: " & PointNumber)
        
        ReDim tempStringArray(1 To PointNumber)
        ReDim tempVolumeArray(1 To PointNumber)
                
        'TODO - use a general-use function for this!
        For i = 1 To PointNumber
            
            'get a small sub buffer out, and separate it into components
            tempData = Mid(DataString, 1 + (i - 1) * conMeta1Point_L, conMeta1Point_L)
            tempTime = BinaryStringToDouble(Mid(tempData, conMeta1Time_IntPos, conMeta1Time_L), LittleEndian:=True)
            tempVolume = BinaryStringToDouble(Mid(tempData, conMeta1Volume_IntPos, conMeta1Volume_L), LittleEndian:=True)
            tempAnnotation = Mid(tempData, conMeta1Text_IntPos, conMeta1Text_L)
            
            tempVolumeArray(i) = tempVolume
            'time?
            tempStringArray(i) = StringRemoveNonPrintable(tempAnnotation)
                                    
        Next i
        
     Set tempcAnnotation = New clsGraphAnnotation
     
     With tempcAnnotation
        .XDataArray = tempVolumeArray
        .AnnotationArray = tempStringArray
    End With

    Set ImportAnnotation = tempcAnnotation
    
    End If
    
    Set tempcAnnotation = Nothing

End Function

Private Sub ImportInjection(ByVal DataString As String)
    'Injection and fraction have exactly the same data formatting
    
    Dim tempcGraphAnnotation As clsGraphAnnotation
    Dim tempSurplusTimeData() As Double
    Dim tempArray() As String
    
    Dim i As Long
    
    Set tempcGraphAnnotation = New clsGraphAnnotation
    
    Set tempcGraphAnnotation = ImportAnnotation(DataString)
    tempcGraphAnnotation.SetType (conUni3InjectionLabel)
        
    Set Annotation = tempcGraphAnnotation
            
    'Unicorn 3 labels each injection with N, where N is number of injections.
    'this kind of annotation is useless, so I replace it by 1, 2, 3, etc
            
        Call ArrayCopy(Annotation.AnnotationArray, tempArray)
        
        For i = 1 To Annotation.Count
            tempArray(i) = CStr(i)
        Next i
        
        Annotation.AnnotationArray = tempArray
       
    Set tempcGraphAnnotation = Nothing
    
End Sub

Private Sub ImportFraction(ByVal DataString As String)

    Dim tempcGraphRegions As clsGraphRegions
    Dim tempcGraphAnnotation As clsGraphAnnotation
    Dim tempSurplusTimeData() As Double
    Dim tempDoubleArray() As Double
    Dim tempStartEndArray() As Double
    Dim tempFinalStrings() As String
    Dim tempStringArray() As String
    
    Dim CurrentlyInWaste As Boolean
    
    Dim FractionsCount As Long
    
    Dim IndexS As Long, IndexE As Long
    
    Dim i As Long, j As Long
    
    Set tempcGraphRegions = New clsGraphRegions
    
    Set tempcGraphAnnotation = ImportAnnotation(DataString)
    
    tempDoubleArray = tempcGraphAnnotation.XDataArray
    tempStringArray = tempcGraphAnnotation.AnnotationArray
    
    IndexS = LBound(tempDoubleArray)
    IndexE = UBound(tempDoubleArray)
    
    FractionsCount = IndexE - IndexS + 1
    
    For i = IndexS To IndexE
        If tempStringArray(i) = conUni3FractionWaste Then
            FractionsCount = FractionsCount - 1
        End If
    Next i
    
    ReDim tempStartEndArray(1 To FractionsCount, 1 To 2)
    ReDim tempFinalStrings(1 To FractionsCount)
        
    j = 1
    
    'CurrentlyInWaste = (tempStringArray(IndexS) = conUni3FractionWaste)
    'If not CurrentlyInWaste then tempStartArray(1) = tempDoubleArray tempStringArray(IndexS)
    
    For i = IndexS To IndexE - 1
                        
        If tempStringArray(i) <> conUni3FractionWaste Then
            tempStartEndArray(j, 1) = tempDoubleArray(i)
            tempStartEndArray(j, 2) = tempDoubleArray(i + 1)
            tempFinalStrings(j) = tempStringArray(i)
            j = j + 1
        End If
        
    Next i
    
    'last one should always be waste fraction. If it's not, that means fractionation wasn't stopped by "Fractionation Stop"
    'but by prematurely ended unicorn run
    If tempStringArray(IndexE) <> conUni3FractionWaste Then
        tempStartEndArray(j, 1) = tempDoubleArray(IndexE)
        tempStartEndArray(j, 2) = 0 'placeholder for improperly terminated fractionation - this has to be corrected using external info!
        tempFinalStrings(j) = tempStringArray(IndexE)
    End If
            
    With tempcGraphRegions
        .XStartEndArray = tempStartEndArray
        .AnnotationArray = tempFinalStrings
        .SetType (conUni3FractionLabel)
    End With
    
    Set AnnotatedRegions = tempcGraphRegions
       
    Set tempcGraphRegions = Nothing
    Set tempcGraphAnnotation = Nothing
    
End Sub

Private Sub ImportLogBook(ByVal DataString As String)
    
    Dim tempcGraphAnnotation As clsGraphAnnotation
    Dim tempSurplusTimeData() As Double
    
    Set tempcGraphAnnotation = New clsGraphAnnotation
    
    With tempcGraphAnnotation
        Call ImportAnnotation(DataString)
    End With
    
    Set Annotation = tempcGraphAnnotation
       
    Set tempcGraphAnnotation = Nothing

End Sub

Private Sub ImportMeta(ByVal MetaString As String)

    Select Case pHeader.Name
        Case "LogBook"
            pLogBook = MetaString
        Case Else
            '"Importing of " & pHeader.Name & " not yet implemented, just skip it, maybe I will never implement most of them
            'Call ErrorReport(, "Importing of " & pHeader.Name & " not yet implemented, or it is not a valid header.")
    End Select

End Sub

Public Property Get ResultName() As String
    
    Const colon As String = ":"
    
    Dim tempString As String
    
    tempString = pHeader.Name
        
    tempString = Left$(tempString, InStr(1, tempString, colon) - 1)
    
    ResultName = tempString

End Property

Public Function GetTitle() As String

    Const uni_meta_L As Long = 224
    Const uni_title_L As Long = 31

    GetTitle = Mid(pLogBook, uni_meta_L + 1, uni_title_L)

End Function

Public Function GetFractions() As clsGraphRegions

    Set GetFractions = AnnotatedRegions

End Function

Public Function GetInjection() As clsGraphAnnotation

    Set GetInjection = Annotation

End Function

Public Function GetData(DataType As eUnicornCurveType) As clsChromatogram

    Dim XInfo As AxisDataType
    Dim YInfo As AxisDataType
    
    Dim ctempChromatogram As clsChromatogram
     
    XInfo = XAxisType
    YInfo = YAxisType
    
    Select Case DataType
        
        Case uniUV
            XInfo.Label = "Elution Volume"
            YInfo.Label = "A" & VBA.Format$(UVWavelength, "0")
        
        Case uniConc
            XInfo.Label = "Elution Volume"
            YInfo.Label = "% of elution buffer"
        
        Case uniCond
            XInfo.Label = "Elution Volume"
            YInfo.Label = "Conductivity"
        
        Case Else
            
            '
    End Select
    
    Set ctempChromatogram = New clsChromatogram
    
    With ctempChromatogram
    
        Set .ChromatogramData = Graph
        .XAxisInfo = XInfo
        .YAxisInfo = YInfo
        
    End With
    
    Set GetData = ctempChromatogram
    
    Set ctempChromatogram = Nothing
    
End Function

Public Function GetUV() As clsChromatogram

    Dim tempAxisData As AxisDataType
    
    Dim ctempChromatogram As clsChromatogram
   
    Set ctempChromatogram = New clsChromatogram
    
    With ctempChromatogram
    
        Set .ChromatogramData = Graph
        
        tempAxisData.Label = "Elution Volume"
        tempAxisData.Unit = XAxisType.Unit
        
        .XAxisInfo = tempAxisData
        
        '.XAxis.Label = "Elution Volume"
        '.XAxis.Unit = "" 'read from meta
        
        tempAxisData.Label = "A" & VBA.Format$(UVWavelength, "0")
        tempAxisData.Unit = YAxisType.Unit
        
        .YAxisInfo = tempAxisData
        
        '.YAxis.Label = "A" & "280" 'read from metadata!
        '.YAxis.Unit = "" 'read from meta
        
    End With
    
    Set GetUV = ctempChromatogram
    
    Set ctempChromatogram = Nothing

End Function



Public Sub tempexport(x, y)
    
    Dim W As Long
    Dim H As Long
    
    Dim i As Long
    Dim j As Long
    
    Dim expy() As Double
    
    W = 2
    H = UBound(x, 1) - LBound(x, 1) + 1
    
    ReDim expy(1 To H, 1 To 2)
    
    For i = 1 To H
        expy(i, 1) = x(i)
        expy(i, 2) = y(i)
    Next i
    
    Range("H20").Resize(H, W).Value = expy

End Sub


