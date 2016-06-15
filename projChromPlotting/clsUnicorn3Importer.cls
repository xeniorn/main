VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUnicorn3Importer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2016-05-09, for this and that purpose
'Last update 2016-05-18
'2016-06-12
'====================================================================================================
Option Explicit

Private Const conClassName As String = "clsUnicorn3Importer"

'This is an example Class, using all the basic elements I deem important for classes to have in general
'not all elements need to be used for each class, of course, but it must absolutely have initialization
'for object parameters if it has them, and termination of them to prevent memory leaks, an error reporting
'method so it allows for robust and modular error reporting from within the class, and ideally the clone method,
'which allows creating a new instance of the object that is identical to the parent

Private pFullFileName As String
Private pDataSets As Collection 'change it to a type-safe collection!
'Private pParentExperiment As IChromExp

Public Event InvalidInput()

Private Sub ErrorReport(Optional ErrorNumber As Long = 0, Optional ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub

'****************************************************************************************************
Private Function SelectData(ByVal DataType As eUnicornDataType, _
                    Optional ByVal CurveType As eUnicornCurveType = uniOther, _
                    Optional ByVal MetaName As String = "", _
                    Optional ByVal ChromName As String = "" _
                    ) As clsUnicorn3Data
   
    Dim i As Byte
    Dim ctempUnicornData As clsUnicorn3Data
    Dim FoundIt As Boolean
    
    Set ctempUnicornData = Nothing
    
    For i = 1 To pDataSets.Count
        
        Set ctempUnicornData = pDataSets.Item(i)
        
        FoundIt = False
        
        With ctempUnicornData
            
            If .DataType = DataType Then
                    
                Select Case DataType
                    
                    Case uniData, uniFraction, uniInjection, uniLogBook
                        If .ChromName = ChromName Or ChromName = "" Then 'dangerous, it's not necessarily unique without chromatogram name! TODO:
                            If .CurveType = CurveType Then
                                FoundIt = True
                            End If
                        End If
                    
                    'maybe add also curve check here
                    Case uniMeta
                        If .ChromName = ChromName Or ChromName = "" Then
                            If .Name = MetaName Then
                                FoundIt = True
                            End If
                        End If
            
                End Select
                
                If FoundIt Then Exit For
                
            End If
                
        End With
        
    Next i
    
    Set SelectData = ctempUnicornData
    
    Set ctempUnicornData = Nothing

End Function

'****************************************************************************************************
Public Function GetInjections() As clsGraphAnnotation
        
    Dim tempObj As clsUnicorn3Data
    
    Set tempObj = SelectData(uniInjection, uniOther)
    
    If Not (tempObj Is Nothing) Then
        Set GetInjections = tempObj.GetInjection
    End If
    
    Set tempObj = Nothing

End Function

'****************************************************************************************************
Public Function GetFractions() As clsGraphRegions
        
    Dim tempObj As clsUnicorn3Data
    
    Set tempObj = SelectData(uniFraction, uniOther)
    
    If Not (tempObj Is Nothing) Then
        Set GetFractions = tempObj.GetFractions
    End If
    
    Set tempObj = Nothing

End Function

'****************************************************************************************************
Public Function GetData(DataType As eUnicornCurveType) As clsChromatogram

    Dim tempObj As clsUnicorn3Data
    
    Set tempObj = SelectData(uniData, DataType)
    
    If Not (tempObj Is Nothing) Then
        Set GetData = tempObj.GetData(DataType)
    End If
    
    Set tempObj = Nothing
        
End Function

'****************************************************************************************************
Public Function GetUV() As clsChromatogram

    Dim tempObj As clsUnicorn3Data
    
    Set tempObj = SelectData(uniData, uniUV)
    
    If Not (tempObj Is Nothing) Then
        Set GetUV = tempObj.GetUV
    End If
    
    Set tempObj = Nothing
        
End Function

'****************************************************************************************************
Public Function GetResultName() As String

    'GetResultName = SelectData(uniLogBook, uniOther, "", "").ResultName
    GetResultName = FileSystem_GetFilename(pFullFileName, False)
    
End Function

'****************************************************************************************************
Public Property Get LogBook() As clsUnicorn3Data

    Set LogBook = SelectData(uniMeta, uniOther, "LogBook", "")

End Property

'****************************************************************************************************
Public Function GetDateTime() As String
   
    Const year_pos As Long = 18
    Const month_pos As Long = 15
    Const day_pos As Long = 12
    Const hour_pos As Long = 24
    Const minute_pos As Long = 27
        
    
    Dim SourceString As String
    
    Dim year As String
    Dim month As String
    Dim day As String
    Dim hour As String
    Dim minute As String
        
    SourceString = LogBook.GetTitle
    
    year = Mid(SourceString, year_pos, 4)
    month = Mid(SourceString, month_pos, 2)
    day = Mid(SourceString, day_pos, 2)
    hour = Mid(SourceString, hour_pos, 2)
    minute = Mid(SourceString, minute_pos, 2)
    
    GetDateTime = year & "-" & month & "-" & day & " " & hour & ":" & minute
    
End Function

'****************************************************************************************************
Public Sub ImportUNI3(InputFilename As String, DataToProcess As VBA.Collection)
'todo: change the DataToProcess to a type-safe collection!!!
    
    'iterators
    Dim i As Long, j As Long
    
    'temps
    Dim tempHeaderString As String
    Dim tempReadingOffset As Long
    
    Dim tempRead As String
    
    'Input
    Dim ResFile As String
    
    'Decriptors
    Dim FileLength As Long
    
    'Containers
    Dim Headers As Collection
    Dim DataSets As Collection
    Dim ctempUnicornData As clsUnicorn3Data
    
    'temp objects
    Dim ctempHeader As clsUnicorn3Header
        
    ResFile = ReadBinaryFile(InputFilename)
    
    If ResFile = vbNullString Then
        RaiseEvent InvalidInput
        Exit Sub
    End If
        
    'check if it is a res file (magic ID signature)
    
    tempRead = Mid(ResFile, conMagicID_Pos, conMagicID_L)
    
    'all files have a signature in the beginning
    If AsHex(tempRead) <> conMagicID_RES Then
        RaiseEvent InvalidInput
        Exit Sub
        'Call Err.Raise(1, , "File lacks a proper .res file Signature (Unicorn 5)")
    End If
    
    tempRead = Mid(ResFile, conFileSize_Pos, conFileSize_L)
    
    FileLength = BinaryStringToInt32(tempRead, LittleEndian:=False)
       
    
    'collection of headers
    Set Headers = New Collection
    
    Do
        'each following header has i incremented
        i = i + 1
        
        'starting position of headers + i times the total header length
        tempReadingOffset = conHeaderStart_Pos + (i - 1&) * conHeaderTotal_L
        
        'the full string of the header
        tempHeaderString = VBA.Strings.Space(conHeaderTotal_L)
        tempHeaderString = Mid(ResFile, tempReadingOffset, conHeaderTotal_L)
        
        'DEBUGGING, OLD
        'If InStr(1, ResFile, tempHeaderString, vbBinaryCompare) > 0 Then
        '    Debug.Print i & ": OK!"
        'Else
        '    Debug.Print i & ": ERROR!!!"
        'End If
        
        'if the read header is completely empty (last header reached)
        If StringCharCount(tempHeaderString, Chr(0)) = conHeaderTotal_L Then
            
            'stop importing headers
            Exit Do
        
        Else
            
            'create a new header
            Set ctempHeader = New clsUnicorn3Header
            
            Call WriteBinaryFile(ResFile)
            
            'define it
            ctempHeader.ImportHeader (tempHeaderString)
            
            'add it to the collection
            Call Headers.Add(ctempHeader, ctempHeader.Name)
            
            Set ctempHeader = Nothing
            
            tempHeaderString = ""
            
        End If
        
    Loop
    
    'Used in debugging:
    'Call tempOutput(Headers)
    
    Dim IShouldImport As Boolean
    
    Dim DatumToProcess As clsUnicorn3FullType
    
    Set DataSets = New Collection
    
    'create all datasets, import headers first (so if there is an error, we don't waste time importing data)
    'for each data type requested
    For i = 1 To DataToProcess.Count
    
        Set DatumToProcess = DataToProcess.Item(i)
        
        'search through all detected headers
        For j = 1 To Headers.Count
            
            Set ctempHeader = Headers.Item(j)
            
            IShouldImport = False
            
            'if the currently sampled header is a data header, search by curve type, otherwise search by general data type
            
            
            If ctempHeader.DataType = DatumToProcess.DataType Then
                If ctempHeader.CurveType = DatumToProcess.CurveType Then
                    IShouldImport = True
                End If
            End If
                        
            If IShouldImport Then
                Set ctempUnicornData = New clsUnicorn3Data
                Call ctempUnicornData.ImportHeader(ctempHeader)
                DataSets.Add ctempUnicornData
            End If
            
                
            'if it is a class with a single possible dataset, just exit
            'DON'T THIS IS SO WRONG OMG
            'If Not (DataToProcess.Item(i) = uniMeta) _
            '    And Not _
            '(ctempHeader.DataType = uniData And DataToProcess.Item(i) = uniUV) _
            'Then Exit For
            
        Next j
        
    Next i
    
    'import data into datasets
    For i = 1 To DataSets.Count
        
        Set ctempUnicornData = DataSets.Item(i)
        ctempUnicornData.ImportFromRaw (ResFile)
    
    Next i
    
    Set pDataSets = DataSets
    
    pFullFileName = InputFilename
    
    Set ctempHeader = Nothing
    Set Headers = Nothing
    Set ctempUnicornData = Nothing
    Set DataSets = Nothing

End Sub


Function ConstructHeaderArray() As String()

    Dim ctempHeader As clsUnicorn3Header
    
    Dim i As Long
    
    Dim tempOutArray(1 To 1, 1 To 6)
    
    'Dim outrange As Excel.Range
        
    ReDim tempOutArray(1 To Headers.Count, 1 To 6)
        
    For i = 1 To Headers.Count
        
        Set ctempHeader = Headers.Item(i)
        tempOutArray(i, 1) = ctempHeader.MagicID
        tempOutArray(i, 2) = CStr(ctempHeader.Name)
        tempOutArray(i, 3) = ctempHeader.DataSize
        tempOutArray(i, 4) = ctempHeader.DataOffsetToNext
        tempOutArray(i, 5) = ctempHeader.DataAddress
        tempOutArray(i, 6) = ctempHeader.OffsetMetaToData
        
    Next i
    
    'outrange.Value = tempOutArray

    ConstructHeaderArray = tempOutArray
    
End Function

'****************************************************************************************************
Function DefineFullType(t As eUnicornDataType, C As eUnicornCurveType) As clsUnicorn3FullType

    Dim tempFullType As clsUnicorn3FullType
    
    Set tempFullType = New clsUnicorn3FullType
    
    tempFullType.CurveType = C
    tempFullType.DataType = t
    
    Set DefineFullType = tempFullType
    
    Set tempFullType = Nothing

End Function
