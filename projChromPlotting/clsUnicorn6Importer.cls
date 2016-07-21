VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUnicorn6Importer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2016-06-09
'Last update 2016-06-12
'====================================================================================================
Option Explicit

Private Const conClassName As String = "clsUnicorn6Importer"

'appended to subfolders that contain the data to be extracted, from the level-2 zip files in the result file
Const conDataFolderExt As String = "_unzip"
'self-explanatory
Const conDefaultSystemSeparator As String = "\"
'offset required to start reading the data in the data files (there are 35 control chars in the beginning + 1 in the end)
Const conDataOffset As Long = 35

Private pFullFileName As String
Private pTempPath As String
Private pDataSets As Collection
Private pXMLObject As MSXML2.DOMDocument60

Private SystemSeparator As String
'Private pParentExperiment As IChromExp

Private pFileValid As Boolean

Public Event InvalidInput()

'pDataSets structure: Collection of (1:[VBA.Collection:the data type] 2:[clsChromatogram:the actual data] 3:[String]:base name of data file)
'   data type collection is as defined in the DefineFullType (1:[String:CurveType] 2:[String:CurveName])

Private Sub Class_Initialize()

    Set pXMLObject = New MSXML2.DOMDocument60
    Set pDataSets = New Collection
    
    SystemSeparator = conDefaultSystemSeparator

End Sub

Private Sub Class_Terminate()

    Set pXMLObject = Nothing
    Set pDataSets = Nothing
    
    Call FileSystem_DeleteFolder(pTempPath)

End Sub

Private Sub ErrorReport(Optional ErrorNumber As Long = 0, Optional ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub

Private Property Get XML() As MSXML2.DOMDocument60
    Set XML = pXMLObject
End Property

'****************************************************************************************************
Public Function GetInjections() As clsGraphAnnotation
    
    Const conCurveType = "Injection"
    
    Dim tXPath As String
    
    Dim NodeSelection As IXMLDOMNodeList
    Dim tNode As IXMLDOMNode
            
    Dim InjectionsNumber As Long
    
    Dim tempInjections As clsGraphAnnotation
    
    Dim tAnnotationArray() As String
    Dim tVolumesArray() As Double
    
    Dim tempText As String
        
    Dim i As Long
    
    'get path to injection events
        tXPath = "/Chromatogram/EventCurves/EventCurve[@EventCurveType=""" & conCurveType & """]/Events"
    
    'select all subnodes there
        Set NodeSelection = XML.SelectSingleNode(tXPath).ChildNodes
    
    'if there are any, import injections
        If Not NodeSelection Is Nothing Then
        
            InjectionsNumber = NodeSelection.Length
            
            'initialize the Graph Annotation object and containers
            Set tempInjections = New clsGraphAnnotation
            tempInjections.SetType ("INJECTION")
            ReDim tAnnotationArray(1 To InjectionsNumber)
            ReDim tVolumesArray(1 To InjectionsNumber)
            
            'loop through all injection events inside the XML
            For i = 1 To InjectionsNumber
                
                'select single event
                Set tNode = NodeSelection.Item(i - 1)
                
                'grab volume
                    tempText = tNode.SelectSingleNode("EventVolume").Text
                    tVolumesArray(i) = Val(tempText)
                'set annotation
                    tAnnotationArray(i) = CStr(i)
                
            Next i
            
            'store the temp containers into the graph annotation object
            With tempInjections
                .AnnotationArray = tAnnotationArray
                .XDataArray = tVolumesArray
            End With
            
        End If
                
            
    
    'output result
        Set GetInjections = tempInjections
        
    'dereference objects
        Set tempInjections = Nothing
        Set NodeSelection = Nothing
        Set tNode = Nothing

End Function

'****************************************************************************************************
Public Function GetFractions() As clsGraphRegions
        
    
    Const conCurveType As String = "Fraction"
    Const conWaste As String = "WASTE"
    
    Dim tXPath As String
    
    Dim NodeSelection As IXMLDOMNodeList
    Dim tNode As IXMLDOMNode
            
    Dim FractionsNumber As Long
    Dim RealFractionsNumber As Long
    
    Dim tempFractions As clsGraphRegions
    
    Dim tAnnotationArray() As String
    Dim tVolumesArray() As Double
    Dim FinalAnnotationArray() As String
    Dim FinalVolumesArray() As Double
    
    Dim tempText As String
        
    Dim i As Long
    Dim j As Long
    
    'get path to Fraction events
        tXPath = "/Chromatogram/EventCurves/EventCurve[@EventCurveType=""" & conCurveType & """]/Events"
    
    'select all subnodes there
        Set NodeSelection = XML.SelectSingleNode(tXPath).ChildNodes
    
    'if there are any, import fractions
        If Not NodeSelection Is Nothing Then
        
            FractionsNumber = NodeSelection.Length
            RealFractionsNumber = FractionsNumber
            
            'initialize the Graph Annotation object and containers
            Set tempFractions = New clsGraphRegions
            tempFractions.SetType ("FRACTION")
            ReDim tAnnotationArray(1 To FractionsNumber)
            ReDim tVolumesArray(1 To FractionsNumber)
            
            'loop through all fraction events inside the XML
            For i = 1 To FractionsNumber
                
                'select single event
                Set tNode = NodeSelection.Item(i - 1)
                
                'grab volume
                    tempText = tNode.SelectSingleNode("EventVolume").Text
                    tVolumesArray(i) = Val(tempText)
                'grab fraction name + parse (remove dots)
                    tempText = tNode.SelectSingleNode("EventText").Text
                    tempText = Replace(tempText, ".", "")
                    If UCase(tempText) Like "*WASTE*" Then
                        tempText = conWaste
                        RealFractionsNumber = RealFractionsNumber - 1
                    End If
                    tAnnotationArray(i) = tempText
                
            Next i
            
            'remove "waste" from fractions, resize the arrays
            ReDim FinalVolumesArray(1 To RealFractionsNumber, 1 To 2)
            ReDim FinalAnnotationArray(1 To RealFractionsNumber)
            For i = 1 To FractionsNumber
                If tAnnotationArray(i) <> conWaste Then
                    j = j + 1
                    FinalVolumesArray(j, 1) = tVolumesArray(i)
                    FinalVolumesArray(j, 2) = tVolumesArray(i + 1)
                    FinalAnnotationArray(j) = tAnnotationArray(i)
                End If
            Next i
                
            
            
            'store the temp containers into the graph annotation object
            With tempFractions
                .AnnotationArray = FinalAnnotationArray
                .XStartEndArray = FinalVolumesArray
            End With
            
        End If
                
            
    
    'output result
        Set GetFractions = tempFractions
        
    'dereference objects
        Set tempFractions = Nothing
        Set NodeSelection = Nothing
        Set tNode = Nothing
      

End Function

'****************************************************************************************************
Private Function GetCurveType(DatasetItem As VBA.Collection) As String
    
    If Not DatasetItem Is Nothing Then
        GetCurveType = DatasetItem.Item(1).Item(1)
    End If

End Function

'****************************************************************************************************
Private Function GetCurveName(DatasetItem As VBA.Collection) As String
    
    If Not DatasetItem Is Nothing Then
        GetCurveName = DatasetItem.Item(1).Item(2)
    End If

End Function

'****************************************************************************************************
Private Function GetCurveFilenameBase(DatasetItem As VBA.Collection) As String
    
    If Not DatasetItem Is Nothing Then
        GetCurveFilenameBase = DatasetItem.Item(3)
    End If

End Function

'****************************************************************************************************
Private Function GetChromatogramObject(DatasetItem As VBA.Collection) As clsChromatogram
    
    If Not DatasetItem Is Nothing Then
        Set GetChromatogramObject = DatasetItem.Item(2)
    End If

End Function


'****************************************************************************************************
Public Function GetData(ByVal DataType As String) As clsChromatogram

    Dim CurveType As String
    Dim CurveName As String
    Dim i As Long
    
    Dim tempChrom As clsChromatogram
    
    'translate data type to CurveType + CurveName
    Select Case DataType
        Case "UV"
            CurveType = "UV"
            CurveName = "UV"
        Case "CONC"
            CurveType = "Other"
            CurveName = "Conc B"
        Case "COND"
            CurveType = "Conduction"
            CurveName = "Cond"
        Case Else
            ErrorReport , "DataType invalid or not yet implemented: " & DataType
    End Select
    
    'Pick the chromatogram from the correct data item
    With pDataSets
        For i = 1 To .Count
            If GetCurveType(.Item(i)) = CurveType Then
                If GetCurveName(.Item(i)) = CurveName Then
                    Set tempChrom = pDataSets.Item(i).Item(2)
                    Exit For
                End If
            End If
        Next i
    End With
    
    If Not tempChrom Is Nothing Then
        Set GetData = tempChrom
    Else
        ErrorReport , "Tried to pick data that isn't there: " & CurveType & ":" & CurveName
    End If
    
    Set tempChrom = Nothing
        
End Function

'****************************************************************************************************
Public Function GetResultName() As String

    GetResultName = FileSystem_FormatFilename(pFullFileName, False, False, False)
    
End Function

'****************************************************************************************************
Public Function GetDateTime() As String
   
   Dim tempString As String
   
   'extract date from XML
   tempString = XML.SelectSingleNode("Chromatogram/Created").Text
   
   'replace the middle T wih a space (format in XML is YYYY-MM-DD"T"HH:mm:ss.sss)
   tempString = Replace(tempString, "T", " ")
   
   'truncate at seconds
   tempString = Left(tempString, 19)
   
   GetDateTime = tempString
   
End Function

Private Function GetYAxisLabel(ByVal DatumToProcess As VBA.Collection, ByVal XMLCurveName As String)

    Dim tWavelength As String
    Dim tStr As String

    Select Case DatumToProcess.Item(1)
        Case "Conduction"
            Select Case DatumToProcess.Item(2)
                Case "Cond"
                    GetYAxisLabel = "Conductivity"
            End Select
        Case "UV"
            Select Case DatumToProcess.Item(2)
                Case "UV"
                    GetYAxisLabel = "UV280"
                Case Else
                    tStr = XMLCurveName
                    tWavelength = Mid(tStr, InStr(1, tStr, "_"), Len(tStr))
                    GetYAxisLabel = "UV" & tWavelength
            End Select
        Case "Other"
            Select Case DatumToProcess.Item(2)
                Case "Conc B"
                    GetYAxisLabel = "% of elution buffer"
            End Select
        Case "Pressure"
        Case Else
    End Select

End Function


'****************************************************************************************************
Private Sub ParseDataToProcess(ByRef DataToProcess As VBA.Collection, ByRef FilenamesToProcess As VBA.Collection)

    Dim i As Long
    Dim j As Long
    
    Dim tCurveType As String
    Dim tCurveNameBase As String
    Dim tCurveFilename As String
    Dim tCurveID As String
    Dim tXPath As String
    
    Dim tAxisType As AxisDataType
    
    'temp objects
    Dim tempSelection As IXMLDOMSelection
    Dim tempNodes As IXMLDOMNodeList
    Dim tempNode As IXMLDOMNode
    Dim tempCollection As VBA.Collection
    Dim tempChromatogram As clsChromatogram
    Dim tempGraph As clsGraphData
    
    Debug.Print ("Parsing XML...")
    
    Set FilenamesToProcess = New VBA.Collection

    'for each DataToProcess entry
    For j = 1 To DataToProcess.Count
    
        'extract the curve type and curve name
        tCurveType = DataToProcess.Item(j).Item(1)
        tCurveNameBase = DataToProcess.Item(j).Item(2)
                
        'define the XPath string to use for the curve type
        tXPath = "/Chromatogram/Curves/Curve[@CurveDataType=""" & tCurveType & """]"
        
        'get collection of all elements of the given curve type
        Set tempSelection = XML.SelectNodes(tXPath)
        
        'for each curve of this type in the xml file
        For i = 0 To tempSelection.Length - 1
        
            'look for curves with the right name
            Set tempNode = tempSelection.Item(i).SelectSingleNode("./Name")
            If Not tempNode Is Nothing Then
                tCurveID = tempNode.Text
                If tCurveID Like "*" & tCurveNameBase & "*" Then
                
                    'and extract the filename of the curve
                    Set tempNode = tempSelection.Item(i).SelectSingleNode("CurvePoints/CurvePoint/BinaryCurvePointsFileName")
                    If Not tempNode Is Nothing Then
                        tCurveFilename = tempNode.Text
                    
                        'add the filename to the FilenamesToProcess list
                        FilenamesToProcess.Add tCurveFilename, tCurveFilename
                        
                        'define a new element to add to the data collection of the importer
                        Set tempCollection = New Collection
                        tempCollection.Add DataToProcess.Item(j)
                        
                        'define the new chromatogram, and add it to the collection together with the label (DataToProcess element)
                        Set tempChromatogram = New clsChromatogram
                        tempCollection.Add tempChromatogram
                        
                        'also add the filename base, so it's easier to pick it later!
                        tempCollection.Add tCurveFilename
                                                
                        'add an empty graph to the chromatogram (to add metadata)
                        Set tempGraph = New clsGraphData
                        Set tempChromatogram.ChromatogramData = tempGraph
                        
                        'append the top-level container object for the datasets
                        pDataSets.Add tempCollection
                        
                        'TODO: add other data one can extract from here
                        Set tempNode = tempSelection.Item(i)
                        With tempNode
                        
                            'extract and parse X axis info, add it to graph
                            tAxisType.Label = "Elution Volume"
                            tAxisType.Unit = .SelectSingleNode("VolumeUnit").Text
                            Select Case tAxisType.Unit
                                Case "ml": tAxisType.Unit = "mL"
                                Case "l": tAxisType.Unit = "L"
                                Case "ul": tAxisType.Unit = "uL"
                                Case "nl": tAxisType.Unit = "nL"
                            End Select
                            tempGraph.XAxisInfo = tAxisType
                            
                            'extract and parse Y axis info, add it to graph
                            tAxisType.Label = GetYAxisLabel(DataToProcess.Item(j), tCurveID)
                            tAxisType.Unit = .SelectSingleNode("AmplitudeUnit").Text
                            tempGraph.YAxisInfo = tAxisType
                                
                        End With
                            
                    End If
                End If
            End If
                
        Next i
            
    Next j

End Sub


'****************************************************************************************************
Public Sub ImportUNI6ZipFile(ByVal InputFilename As String, Optional DataToProcess As VBA.Collection)
    
    
    Const DataFilenameEnding As String = "_True"
    
    'filenames
    Dim TempFolderBase As String
    Dim XMLFilename As String
    Dim DataFilenameBase As String
    Dim DataFilename As String
                
    'temp objects
    Dim tempSelection As IXMLDOMSelection
    Dim tempNodes As IXMLDOMNodeList
    Dim tempNode As IXMLDOMNode
    Dim tempCollection As VBA.Collection
    Dim tempXML As MSXML2.DOMDocument60
    
    'temp vars
    Dim tCurveType As String
    Dim tCurveName As String
    Dim tXPath As String
    
    'descriptors
    Dim NumberOfDataCurves As Long
    
    'Objects
    Dim FilenamesToProcess As VBA.Collection
    
    'iterators
    Dim i As Long
    
        
    Debug.Print ("ImportUNI6: start import")
        
    '===Parsing inputs
    
        If Not FileSystem_FileExists(InputFilename) Then
            pFileValid = False
            Debug.Print ("Input file not found")
            RaiseEvent InvalidInput
            Exit Sub
            'ErrorReport , "Input file not found!"
        Else
            Debug.Print ("Input file found on disk")
            pFileValid = True
        End If
        
        pFullFileName = InputFilename
        
        'if no data was selected to be processed, process only UV
        If DataToProcess Is Nothing Then
        
            Set DataToProcess = New VBA.Collection
            
            Set tempCollection = New VBA.Collection
                tempCollection.Add "UV", "UV"
                tempCollection.Add "UV", "UV"
                
            DataToProcess.Add tempCollection
            
        End If
        
        
    '===Process the inputs:
    
        'TODO:
        'check if it is a .zip file
        If UCase(FileSystem_GetExtension(InputFilename)) <> "ZIP" Then
            pFileValid = False
            RaiseEvent InvalidInput
            Exit Sub
        End If
        
        Debug.Print ("Unzipping result data file...")
        
        'unzip the main zip file
        TempFolderBase = FileSystem_Unzip(InputFilename, , False)
        'store the path to the temp files
        pTempPath = TempFolderBase & SystemSeparator
        
        'TODO:
        'check if necessary files are present (Chrom.#.Xml, Chrom.#_##_TRUE, ...)
        
        'find the main XML file
        XMLFilename = Dir(TempFolderBase & SystemSeparator & "Chrom*.Xml")
        
        'load it into memory
        XML.LoadXML (ReadTextFile(TempFolderBase & SystemSeparator & XMLFilename))
                    
        'parse Chrom.#.Xml file
        Set tempNodes = XML.SelectNodes("/Chromatogram/Curves")
        NumberOfDataCurves = tempNodes.Length
        
        'TODO: extract time of experiment, filename, person?
        
        'define chromatograms and extract their metadata, define which files need to be processed
        Call ParseDataToProcess(DataToProcess, FilenamesToProcess)
        
        'rectify all the non-canonical zip files inside (Chrom.#_##_TRUE are zip files, but zero-padded heavily)
        'files are named e.g. Chrom.1_10_True
        DataFilenameBase = TempFolderBase & SystemSeparator & FileSystem_FormatFilename(XMLFilename, True, False, False) & "_"
        Call RectifyZipFiles(DataFilenameBase, FilenamesToProcess)
        
        'unpack the sub zip files to separate subfolders
        Call UnzipDataFiles(DataFilenameBase)
        
        'extract chromatogram data from files
        Call ImportDatasets
        
        'extract injections
        'extract fractions
    
        'TODO:
        'extract comments / markers



    '===dereference objects
        Set tempNodes = Nothing
        Set tempNode = Nothing
        Set tempCollection = Nothing
        Set DataToProcess = Nothing
        Set FilenamesToProcess = Nothing

End Sub

'****************************************************************************************************
Sub RectifyZipFiles(ByVal BaseFilename As String, ByVal FilenamesToProcess As VBA.Collection)
    
    Const zipTerminalLength As Long = 22
    
    Dim zipEndMarker As String
    
    Dim FilePath As String
    Dim FolderPath As String
    
    Dim FileContents As String
    
    Dim MarkerIndex As Long
    Dim ZipLength As Long
            
    zipEndMarker = HexToString("50 4B 05 06", " ")
    
    FolderPath = FileSystem_GetPath(BaseFilename)
    
    'find the first filename
    FilePath = FolderPath & Dir(BaseFilename & "*_True")
    
    'rectify each zip filename, loop through all of them
    If FilePath = "" Then
        ErrorReport , "No data files present!"
    End If
    
    Do While FileSystem_GetFilename(FilePath) <> ""
        
        'if this is one of the elements to be processed
        If IsElementOf(FileSystem_GetFilename(FilePath), FilenamesToProcess) Then
            
            'load file
            FileContents = ReadBinaryFile(FilePath)
            
            'find where the terminal zip magic ID is
            MarkerIndex = InStrRev(FileContents, zipEndMarker, , vbBinaryCompare)
            
            'define how long the zip portion is
            ZipLength = MarkerIndex + zipTerminalLength - 1
            
            'truncate the file contents
            FileContents = Left(FileContents, ZipLength)
            
            'create truncated, proper-form zip file
            Call CreateEmptyFile(FilePath & ".zip") 'to ensure file is empty, binary write doesn't clear the file!
            Call WriteBinaryFile(FileContents, FilePath & ".zip")
            
            Debug.Print (FileSystem_GetFilename(FilePath) & " rectified.")
            
        End If
        
        'get next File
        FilePath = FolderPath & Dir
        
    
    Loop

End Sub

'****************************************************************************************************
Sub UnzipDataFiles(ByVal BaseFilename As String)
    
    Dim FilePath As String
    Dim FolderPath As String
    Dim OutputPath As String
    
    Dim ProcessThisFile As Boolean
    
    Dim i As Long
    
    Dim FilesCollection As VBA.Collection
        
    FolderPath = FileSystem_GetPath(BaseFilename)
    
    'find the first filename
    FilePath = FolderPath & Dir(BaseFilename & "*_True.zip")
    
    'unzip each zip file, loop through all of them
    If FilePath = "" Then
        ErrorReport , "No data files present!"
    End If
    
    Set FilesCollection = New VBA.Collection
    
    'define a collection with all the zip file paths
    Do While FileSystem_GetFilename(FilePath) <> ""
                
        FilesCollection.Add (FilePath)
        FilePath = FolderPath & Dir
        
    Loop
    
    'unzip file
    
    For i = 1 To FilesCollection.Count
                
        FilePath = FilesCollection.Item(i)
        
        'TODO: selective processing
        ProcessThisFile = True
        
        If ProcessThisFile Then
            OutputPath = FileSystem_FormatFilename(FilePath, True, False, False) & conDataFolderExt & SystemSeparator
            Call FileSystem_Unzip(FilePath, OutputPath, False)
            Debug.Print (FileSystem_GetFilename(FilePath) & " unzipped.")
        End If
        
    Next i

End Sub

'****************************************************************************************************
Private Sub ImportDatasets()
'for each dataset (already set up), grab its data (already unzipped, etc)
    
    Const conVolumeFilename As String = "CoordinateData.Volumes"
    Const conSignalFilename As String = "CoordinateData.Amplitudes"
    
    Dim i As Long
    Dim FilenameBase As String
    Dim PathToData As String
    Dim FullDataFilename As String
            
    Dim Buffer As String
            
    Dim DataUnpacker As clsBinaryUnpacker
            
    Dim tDataItem As VBA.Collection
    Dim tChrom As clsChromatogram
    Dim tempData As clsDataList
    
    Dim tempArrayOut() As Double
    
    
    'initialize unpacker object
        Set DataUnpacker = New clsBinaryUnpacker
    
    'loop through all the pre-initialized dataset items
        For i = 1 To pDataSets.Count
            Debug.Print ("Importing dataset #" & i & "...")
            
            'grab a single item
            Set tDataItem = pDataSets.Item(i)
            
            'extract filename and define path to data
            FilenameBase = GetCurveFilenameBase(tDataItem)
            PathToData = pTempPath & FilenameBase & conDataFolderExt & SystemSeparator
            
            'grab the chromatogram
            Set tChrom = GetChromatogramObject(tDataItem)
            
            'extract volumes from the data file and add to chromatogram
                FullDataFilename = PathToData & conVolumeFilename
                Buffer = ReadBinaryFile(FullDataFilename)
                
                Set tempData = DataUnpacker.UnpackBinaryData( _
                                            FormatString:="<f", _
                                            SourceDataBuffer:=Buffer, _
                                            StartByteOffset:=conDataOffset, _
                                            IgnoreExtraBytes:=True _
                                            ).Item(1)
                                            
                Call ArrayCopy(tempData.DataArray, tempArrayOut)
                tChrom.ChromatogramData.XDataArray = tempArrayOut
                
            'extract signal amplitudes (Yaxis = UV, Cond, ...)
                FullDataFilename = PathToData & conSignalFilename
                Buffer = ReadBinaryFile(FullDataFilename)
                
                Set tempData = DataUnpacker.UnpackBinaryData( _
                                            FormatString:="<f", _
                                            SourceDataBuffer:=Buffer, _
                                            StartByteOffset:=conDataOffset, _
                                            IgnoreExtraBytes:=True _
                                            ).Item(1)
                                            
                Debug.Print ("Data successfully unpacked")
                                            
                Call ArrayCopy(tempData.DataArray, tempArrayOut)
                tChrom.ChromatogramData.YDataArray = tempArrayOut
                
            'set the imported data of the chromatogram to be the "original" state
            'this is not very well implemented, but is ATM necessary (2016-06-12)
            'because the chromatograms are initialized "empty", with GraphData without
            'a Data Array (in order to extract metadata)
                Call tChrom.CommitChangesToOriginal
                        
        Next i

    'dereference objects
        Set tDataItem = Nothing
        Set DataUnpacker = Nothing
        Set tempData = Nothing
        Set tChrom = Nothing

End Sub

'****************************************************************************************************
Function DefineFullType( _
    CurveType As String, _
    ByVal CurveName As String _
    ) As VBA.Collection
    
    Dim tempColl As VBA.Collection
    
    Set tempColl = New VBA.Collection
    
    Call tempColl.Add(CurveType)
    Call tempColl.Add(CurveName)
        
    Set DefineFullType = tempColl
        
    Set tempColl = Nothing

End Function



