VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGeneralizedChromatography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'DEPENDS ON: modBinary (2016-05-08)

'****************************************************************************************************
'====================================================================================================
'Class for SEC experiments
'Juraj Ahel, 2016-05-08, for automatic handling of SEC data
'Last update 2016-06-12
'2016-06-12 add data alignment function
'====================================================================================================
Option Explicit

Implements IExperiment
Implements IChromExp

Const conClassName As String = "clsGeneralizedChromatography"

Const conFileTypeUNI3 As String = "UNICORN3"
Const conFileTypeUNI6 As String = "UNICORN6"
Const conFileTypeTSV As String = "TSV"
Const conFileTypeCSV As String = "CSV"

'[Var Declaration]
Private pChromatograms As clsChromatograms
Private pMetadata As clsSECMetaData
Private pLabeledRegions As clsGraphRegions
Private pLabeledPoints As clsGraphAnnotation
Private pFractions As clsGraphRegions
Private pInjections As clsGraphAnnotation

Private WithEvents Uni3Importer As clsUnicorn3Importer
Attribute Uni3Importer.VB_VarHelpID = -1
Private WithEvents Uni6Importer As clsUnicorn6Importer
Attribute Uni6Importer.VB_VarHelpID = -1

Private pInputIsInvalid As Boolean

Public Event InvalidInput()

'[Object Initialization]
Private Sub Class_Initialize()

    'Set pLabeledRegions = New clsGraphRegions
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
     
    Set Uni3Importer = Nothing
    Set Uni6Importer = Nothing

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

Public Property Set LabeledRegions(GReg As clsGraphRegions)
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
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************



'****************************************************************************************************
Public Sub ImportFile(ByVal InputFilename As String, _
                    Optional ByVal FileType As String = conFileTypeUNI3)
    
    'check if file name is of acceptable format
    'check if file exists
        
    Debug.Print ("Chromatography: Start file input")
        
    Select Case UCase(FileType)
    
        Case conFileTypeUNI3
            Call ImportUNI3(InputFilename)
            
        Case conFileTypeUNI6
            Call ImportUNI6(InputFilename)
            
        Case Else
            Call ErrorReport(, "Unrecognized file type. Check input parameters. Filetype possibly not supported.")
            
    End Select
    


End Sub

'****************************************************************************************************
Private Sub ImportUNI6(ByVal InputFilename As String)
        
    Dim DataToProcess As VBA.Collection
    
    Dim tempObj As Object
    
    'initialize the importer
    Set Uni6Importer = New clsUnicorn6Importer
    
    'set up the data import parameters (which curves are to be imported)
    Set DataToProcess = New VBA.Collection
        
    With DataToProcess
    
        .Add Uni6Importer.DefineFullType("UV", "UV")
        .Add Uni6Importer.DefineFullType("Conduction", "Cond")
        .Add Uni6Importer.DefineFullType("Other", "Conc B")
        .Add Uni6Importer.DefineFullType("Fraction", "Fraction")
        .Add Uni6Importer.DefineFullType("Injection", "Injection")
        .Add Uni6Importer.DefineFullType("Logbook", "Run Log")
        
    End With
    
    'import the data
    Call Uni6Importer.ImportUNI6ZipFile(InputFilename, DataToProcess)
    
    If Not pInputIsInvalid Then
        
        'TODO:
        'do a check whether multiple chromatograms are present in file, and allow to pick which one you want / process all
        
        With Uni6Importer
            
            'add data chromatograms
            Call AddChromatogram(.GetData("UV"), "UV1")
            Call AddChromatogram(.GetData("CONC"), "CONC")
            Call AddChromatogram(.GetData("COND"), "COND")
            
            'add curve metadata
            Set Injections = .GetInjections
            Set Fractions = .GetFractions
            
            'TODO: add global experiment metadata
            Metadata.ColumnVolume = 0
            Metadata.SampleVolume = 0
            Metadata.ExperimentDate = .GetDateTime
            Metadata.ExperimentName = .GetResultName
            Metadata.ExperimentScientist = "AutoImport not supported yet!"
            Metadata.SystemUsed = "AutoImport not supported yet!"
            
        End With
        
    End If
    
    Set Uni6Importer = Nothing
    Set DataToProcess = Nothing

End Sub


'****************************************************************************************************
Private Sub ImportUNI3(ByVal InputFilename As String)

    Dim DataToProcess As VBA.Collection
   
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
        .Add Uni3Importer.DefineFullType(uniData, uniConc)
        .Add Uni3Importer.DefineFullType(uniData, uniCond)
    End With
    
    Call Uni3Importer.ImportUNI3(InputFilename, DataToProcess)
    
    'only if input is valid (passed by events), proceed to grab the data from the importer
    If Not pInputIsInvalid Then
    
            'Set ctempChromatogram = Uni3Importer.GetUV
            Call AddChromatogram(Uni3Importer.GetData(uniUV), "UV1")
            Call AddChromatogram(Uni3Importer.GetData(uniConc), "CONC")
            Call AddChromatogram(Uni3Importer.GetData(uniCond), "COND")
            
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
            
    End If
            
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

Public Sub RevertAllToOriginal()

    Dim i As Long
        
    For i = 1 To Chromatograms.Count
        Chromatograms.Item(i).RevertToOriginal
    Next i

End Sub



Public Sub TempCleanUp(SelectedInjection As Long)
    
    Dim i As Long
        
    'Metadata.ColumnVolume = 130
    Call NormalizeVolumeToInjectionNumber(SelectedInjection)
    
    Call TruncateToColumnVolume
    Call AlignDataStartToZero
    
    For i = 1 To Chromatograms.Count
        Chromatograms.Item(i).ThinData (0.5)
    Next i
    
    
    'Call NormalizeToGlobalMaximum(1)

End Sub

Public Function TempGetGraph(ChromatogramLabel As String) As clsGraphData

    Dim W As Long
    Dim H As Long
    
    Dim i As Long
    Dim j As Long
    
    Dim x() As Double
    Dim y() As Double
    
    Dim expy() As Double
    
    Dim tempcChromatogram As clsChromatogram
    
    Set tempcChromatogram = Chromatograms.Item(ChromatogramLabel)
        
    Set TempGetGraph = tempcChromatogram.ChromatogramData
    
    Set tempcChromatogram = Nothing

End Function

Public Function TempGetGraphDirect(ChromatogramLabel As String) As Double()

    Dim W As Long
    Dim H As Long
    
    Dim i As Long
    Dim j As Long
    
    Dim x() As Double
    Dim y() As Double
    
    Dim expy() As Double
    
    Dim tempcChromatogram As clsChromatogram
    
    Set tempcChromatogram = Chromatograms.Item(ChromatogramLabel)
    
    x = tempcChromatogram.ChromatogramData.XDataArray
    y = tempcChromatogram.ChromatogramData.YDataArray
    
    W = 2
    H = UBound(x, 1) - LBound(x, 1) + 1
    
    ReDim expy(1 To H, 1 To W)
    
    For i = 1 To H
        expy(i, 1) = x(i)
        expy(i, 2) = y(i)
    Next i
    
    TempGetGraphDirect = expy
    
    Set tempcChromatogram = Nothing

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



Public Sub AddChromatogram(ByVal InpChromatogram As clsChromatogram, ChromatogramID As String)

    Call Chromatograms.Add(InpChromatogram, ChromatogramID)

End Sub

'DATA MANIPULATION
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************

'****************************************************************************************************
Public Sub AlignDataStartToZero(Optional ByVal ZeroValue As Double = 0)

    Dim tempcChromatogram As clsChromatogram
    Dim i As Long
    Dim ShiftAmount As Double
    
    Debug.Print ("Aligning chromatogram(1) start to " & ZeroValue)
    
    'one cannot call align to zero for all chromatograms, because each would shift independently,
    'breaking any coupling of the data! Chromatogram(1) is shifted to zero, others just the same amount
    'Chrom1 is the one used to define the X axis in graphing usually. This is quite WD40 + ducky tape, but
    'can't really be helped AFAIK
    Set tempcChromatogram = Chromatograms.Item(1)
    If Not tempcChromatogram Is Nothing Then
        ShiftAmount = ZeroValue - tempcChromatogram.ChromatogramData.XData(1)
    End If
    
    Call ShiftDataX(ShiftAmount)
    
    Set tempcChromatogram = Nothing
    

End Sub

'****************************************************************************************************
Public Sub ShiftDataX(ByVal ShiftAmount As Double)

    Dim tempcChromatogram As clsChromatogram
    Dim i As Long
    
    Debug.Print ("Shifting X-axis by " & ShiftAmount)
    
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
    
    Debug.Print ("Truncating data to volume range " & StartVolume & " - " & EndVolume)
    
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
    
    If Metadata.ColumnVolume <= 0 Then
        Call TruncateToVolumeRange(0, 0)
    Else
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

Public Sub ThinData(ByVal DistanceBetweenPoints As Double)

    Dim i As Long
    
    For i = 1 To Chromatograms.Count
        Chromatograms.Item(i).ThinData (0.5)
    Next i

End Sub

'****************************************************************************************************
Public Sub NormalizeToIntegralInRange(ByVal StartVolume As Double, ByVal EndVolume As Double, ByVal ChromatogramNumber As Long)
    
    Dim tempcChromatogram As clsChromatogram
    
    If ChromatogramNumber > 0 And ChromatogramNumber <= Chromatograms.Count Then
    
        Set tempcChromatogram = Chromatograms.Item(ChromatogramNumber)
        
        If Not (tempcChromatogram Is Nothing) Then
        
            Call tempcChromatogram.NormalizeToAreaUnderCurve(StartVolume, EndVolume)
            
            With tempcChromatogram.YAxisInfo
                .Description = "Relative " & .Description
                .Unit = ""
            End With
            
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
            
            With tempcChromatogram.YAxisInfo
                .Description = "Relative " & .Description
                .Unit = ""
            End With
            
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

Private Sub Uni3Importer_InvalidInput()
    pInputIsInvalid = True
    RaiseEvent InvalidInput
End Sub

Private Sub Uni6Importer_InvalidInput()
    pInputIsInvalid = True
    RaiseEvent InvalidInput
End Sub
