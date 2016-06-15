VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSizeExclusionChromatography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'DEPENDS ON: modBinary (2016-05-08)

'****************************************************************************************************
'====================================================================================================
'Class for SEC experiments
'Juraj Ahel, 2016-05-08, for automatic handling of SEC data
'Last update 2016-05-13
'====================================================================================================
Option Explicit

Implements IExperiment
Implements IChromExp

Const conClassName As String = "clsSizeExclusionChromatography"

Const conFileTypeUNI3 As String = "UNICORN3"

'[Var Declaration]
Private pChromatograms As clsChromatograms
Private pMetadata As clsSECMetaData
Private pLabeledRegions As clsGraphRegions
Private pLabeledPoints As clsGraphAnnotation
Private pFractions As clsGraphRegions
Private pInjections As clsGraphAnnotation

'[Object Initialization]
Private Sub Class_Initialize()

    Set pChromatograms = New clsChromatograms
    Set pMetadata = New clsSECMetaData

End Sub

'[Object Dereferencing]
Private Sub Class_Terminate()

    Set pChromatograms = Nothing
    Set pMetadata = Nothing
    Set pLabeledRegions = Nothing
    Set pLabeledPoints = Nothing
    Set pFractions = Nothing
    Set pInjections = Nothing

End Sub

Private Sub ErrorReport(Optional ErrorNumber As Long = 0, Optional ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub

'[Property Gets and Lets and Sets]
Public Property Get LabeledRegions() As clsGraphRegions
    Set LabeledRegions = pLabeledRegions
End Property

Private Property Set LabeledRegions(GReg As clsGraphRegions)
    Set pLabeledRegions = GReg
End Property

Public Property Get LabeledPoints() As clsGraphAnnotation
    Set LabeledPoints = pLabeledPoints
End Property

Private Property Set LabeledPoints(GRPoint As clsGraphAnnotation)
    Set pLabeledPoints = GRPoint
End Property

Public Property Get Injections() As clsGraphAnnotation
    Set Injections = pInjections
End Property

Private Property Set Injections(GRPoint As clsGraphAnnotation)
    Set pInjections = GRPoint
End Property

Public Property Get Fractions() As clsGraphRegions
    Set Fractions = pFractions
End Property

Private Property Set Fractions(GReg As clsGraphRegions)
    Set pFractions = GReg
End Property

Public Property Get Chromatograms() As clsChromatograms
    Set Chromatograms = pChromatograms
End Property

Private Property Set Chromatograms(ByVal inp_Chromatograms As clsChromatograms)
    Set pChromatograms = inp_Chromatograms
End Property

Public Property Get Metadata() As clsSECMetaData
    Set Metadata = pMetadata
End Property

Private Property Set Metadata(ByVal inp_Metadata As clsSECMetaData)
    Set pMetadata = inp_Metadata
End Property

Public Property Get ExperimentName() As String
    ExperimentName = Metadata.ExperimentName
End Property

Public Property Get ExperimentDate() As String
    ExperimentDate = Metadata.ExperimentDate
End Property

Public Property Get ExperimentScientist() As String
    ExperimentScientist = Metadata.ExperimentScientist
End Property

Public Property Get SampleVolume() As Double
    SampleVolume = pSampleVolume
End Property

Private Property Let SampleVolume(ByVal inp_SampleVolume As Double)
    pSampleVolume = inp_SampleVolume
End Property

Public Property Get SampleName() As String
    SampleName = pSampleName
End Property

Private Property Let SampleName(ByVal inp_SampleName As String)
    pSampleName = inp_SampleName
End Property

'****************************************************************************************************
Public Sub ShiftDataX(ByVal ShiftAmount As Double)

    Dim tempcChromatogram As clsChromatogram
    Dim i As Long
    
    For i = 1 To Chromatograms.Count
        Set tempcChromatogram = Chromatograms.Item(i)
        If Not (tempcChromatogram Is Nothing) Then Call tempcChromatogram.ShiftXAxis(ShiftAmount)
    Next i
    
    If Not (Injections Is Nothing) Then Call Injections.ShiftX(ShiftAmount)
    If Not (Fractions Is Nothing) Then Call Fractions.ShiftX(ShiftAmount)
    If Not (LabeledPoints Is Nothing) Then Call LabeledPoints.ShiftX(ShiftAmount)
    If Not (LabeledRegions Is Nothing) Then Call LabeledRegions.ShiftX(ShiftAmount)
    
    Set tempcChromatogram = Nothing
    
End Sub

'****************************************************************************************************
Public Sub TruncateToVolumeRange(ByVal StartVolume As Double, ByVal EndVolume As Double)

    Dim tempcChromatogram As clsChromatogram
    Dim i As Long
    
    For i = 1 To Chromatograms.Count
        Set tempcChromatogram = Chromatograms.Item(i)
        If Not (tempcChromatogram Is Nothing) Then Call tempcChromatogram.TruncateXAxis(StartVolume, EndVolume)
    Next i
    
    'If Not (Injections Is Nothing) Then Call Injections.TruncateXAxis(StartVolume, EndVolume)
    'If Not (Fractions Is Nothing) Then Call Fractions.TruncateXAxis(StartVolume, EndVolume)
    'If Not (LabeledPoints Is Nothing) Then Call LabeledPoints.ShiftX(ShiftAmount)
    'If Not (LabeledRegions Is Nothing) Then Call LabeledRegions.ShiftX(ShiftAmount)
    
    Set tempcChromatogram = Nothing
    
End Sub


'****************************************************************************************************
Public Sub TruncateToColumnVolume()
    
    If Metadata.ColumnVolume > 0 Then
        Call TruncateToVolumeRange(0, Metadata.ColumnVolume)
    End If

End Sub


'****************************************************************************************************
Public Sub NormalizeVolumeToInjectionNumber(N As Long)

    Call ShiftDataX(-1 * Injections.XData(N))
    
End Sub

'****************************************************************************************************
Public Sub NormalizeVolumeToLastInjection()

    Call NormalizeVolumeToInjectionNumber(UBound(Injections.XDataArray))

End Sub

'****************************************************************************************************
Public Sub NormalizeToIntegralInRange(ByVal StartVolume As Double, ByVal EndVolume As Double, ByVal ChromatogramNumber As Long)
    
    Dim tempcChromatogram As clsChromatogram
    
    If ChromatogramNumber > 0 And ChromatogramNumber <= Chromatograms.Count Then
    
        Set tempcChromatogram = Chromatograms.Item(ChromatogramNumber)
        If Not (tempcChromatogram Is Nothing) Then
            Call tempcChromatogram.NormalizeToAreaUnderCurve(StartVolume, EndVolume)
        End If
        
    Else
    
        Call ErrorReport(, "Tried to grab Chromatogram number that doesn't exist in collection")
        
    End If
    
    Set tempcChromatogram = Nothing
        
End Sub


'****************************************************************************************************
Public Sub NormalizeToMaxValueInRange(ByVal StartVolume As Double, ByVal EndVolume As Double, ByVal ChromatogramNumber As Long)
    
    Dim tempcChromatogram As clsChromatogram
    
    If ChromatogramNumber > 0 And ChromatogramNumber <= Chromatograms.Count Then
    
        Set tempcChromatogram = Chromatograms.Item(ChromatogramNumber)
        If Not (tempcChromatogram Is Nothing) Then
            Call tempcChromatogram.NormalizeToMaxValue(StartVolume, EndVolume)
        End If
        
    Else
    
        Call ErrorReport(, "Tried to grab Chromatogram number that doesn't exist in collection")
        
    End If
    
    Set tempcChromatogram = Nothing
        
End Sub

'****************************************************************************************************
Public Sub NormalizeToGlobalMaximum(ByVal ChromatogramNumber As Long)
    
    Dim tempcChromatogram As clsChromatogram
    
    If ChromatogramNumber > 0 And ChromatogramNumber <= Chromatograms.Count Then
    
        Set tempcChromatogram = Chromatograms.Item(ChromatogramNumber)
        If Not (tempcChromatogram Is Nothing) Then
            Call NormalizeToMaxValueInRange(tempcChromatogram.Xmin, tempcChromatogram.Xmax, ChromatogramNumber)
        End If
        
    Else
    
        Call ErrorReport(, "Tried to grab Chromatogram number that doesn't exist in collection")
        
    End If
    
    Set tempcChromatogram = Nothing
       
        
End Sub


'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Public Sub ImportFile(ByVal InputFilename As String, Optional ByVal FileType As String = conFileTypeUNI3)
    
    'check if file name is of acceptable format
    'check if file exists
        
    Select Case UCase(FileType)
    
        Case conFileTypeUNI3
            Call ImportUNI3(InputFilename)
            
        Case Else
            Call ErrorReport(, "Unrecognized file type. Check input parameters. Filetype possibly not supported.")
            
    End Select

End Sub


'****************************************************************************************************
Private Sub ImportUNI3(ByVal InputFilename As String)

    Dim DataToProcess As VBA.Collection
    Dim Uni3Importer As clsUnicorn3Importer
   
    Dim ctempChromatogram As clsChromatogram
    Dim tempObject As Object
    
    Set Uni3Importer = New clsUnicorn3Importer
    
    Set DataToProcess = New VBA.Collection
    With DataToProcess
        .Add Uni3Importer.DefineFullType(uniData, uniUV)
        .Add Uni3Importer.DefineFullType(uniMeta, uniOther)
        .Add Uni3Importer.DefineFullType(uniLogBook, uniOther)
        .Add Uni3Importer.DefineFullType(uniInjection, uniOther)
        .Add Uni3Importer.DefineFullType(uniFraction, uniOther)
    End With
    
    Call Uni3Importer.ImportUNI3(InputFilename, DataToProcess)
    
            'Set ctempChromatogram = Uni3Importer.GetUV
            Call AddChromatogram(Uni3Importer.GetData(uniUV), "UV")
            
            Set tempObject = Uni3Importer.GetInjections
            If Not (tempObject Is Nothing) Then Set Injections = tempObject
            
            Set tempObject = Uni3Importer.GetFractions
            If Not (tempObject Is Nothing) Then Set Fractions = tempObject
            
            Set tempObject = Nothing
            
            'Call AddChromatogram(Uni3Importer.GetUV)
            
            With Metadata
                .ExperimentScientist = "AutoImport not supported yet!"
                .ExperimentDate = Uni3Importer.GetDateTime '"2016-04-12 17:30" 'import from data!
                .ExperimentName = Uni3Importer.GetResultName 'import from data!
                .SampleVolume = 0# 'import from data!
                .SystemUsed = "AutoImport not supported yet!" 'import from data!
                .ColumnVolume = 0# 'import!!!!
            End With
            
    'temp DEBUG
    'Set ctempChromatogram = Chromatograms.Item(1)
    'ctempChromatogram.ThinData (0.5)
    'Metadata.ColumnVolume = 130
    'Call NormalizeVolumeToLastInjection
    'Call TruncateToColumnVolume
    'Call NormalizeToGlobalMaximum(1)
    'Call tempexport(ctempChromatogram.ChromatogramData.XDataArray, ctempChromatogram.ChromatogramData.YDataArray)
    
    Set tempObject = Nothing
    Set DataToProcess = Nothing
    Set Uni3Importer = Nothing
    Set ctempChromatogram = Nothing
    
End Sub

Public Sub TempCleanUp()

    Chromatograms.Item(1).ThinData (0.5)
    Metadata.ColumnVolume = 130
    Call NormalizeVolumeToLastInjection
    Call TruncateToColumnVolume
    Call NormalizeToGlobalMaximum(1)

End Sub

Public Function TempGetGraph() As Double()

    Dim W As Long
    Dim H As Long
    
    Dim i As Long
    Dim j As Long
    
    Dim x() As Double
    Dim y() As Double
    
    
    Dim expy() As Double
    
    x = Chromatograms.Item(1).ChromatogramData.XDataArray
    y = Chromatograms.Item(1).ChromatogramData.YDataArray
    
    W = 2
    H = UBound(x, 1) - LBound(x, 1) + 1
    
    ReDim expy(1 To H, 1 To 2)
    
    For i = 1 To H
        expy(i, 1) = x(i)
        expy(i, 2) = y(i)
    Next i
    
    TempGetGraph = expy

End Function


'****************************************************************************************************
Public Sub tempexport(ByVal x, ByVal y)
    
    Const Ran As String = "H20"
    
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
    
    Dim aaa As clsChromatogram
    
    Range("H:H").ClearContents
    Range("I:I").ClearContents
       
    
    Set aaa = pChromatograms.Item(1)
    
    'Range(Ran).Offset(-3, 0).Value = a
    'Range(Ran).Offset(-3, 1).Value = a
    Range(Ran).Offset(-2, 0).Value = aaa.XAxisInfo.Label
    Range(Ran).Offset(-2, 1).Value = aaa.YAxisInfo.Label
    Range(Ran).Offset(-1, 0).Value = aaa.XAxisInfo.Unit
    Range(Ran).Offset(-1, 1).Value = aaa.YAxisInfo.Unit
            
    Range(Ran).Resize(H, W).Value = expy

End Sub

Public Sub AddChromatogram(ByVal InpChromatogram As clsChromatogram, CustomKey As String)

    Chromatograms.Add InpChromatogram, CustomKey

End Sub

'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Private Sub IChromExp_AddChromatogram(ByVal InpChromatogram As clsChromatogram)

    AddChromatogam (InpChromatogram)

End Sub

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
