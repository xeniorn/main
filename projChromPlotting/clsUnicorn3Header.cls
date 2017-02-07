VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUnicorn3Header"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'2017-01-27 fix curve type recognition and magic number constants to work with re-saved chromatograms!

Option Explicit

'[Var Declaration]
Private pMagicID As String
Private pName As String
Private pDataSize As Long
Private pDataOffsetToNext As Long
Private pDataAddress As Long
Private pOffsetMetaToData As Long

Public ResName As String
Public ChromName As String
Public CurveName As String

Private pDataType As eUnicornDataType
Private pCurveType As eUnicornCurveType



'[Property Gets and Lets and Sets]

Public Property Get CurveType() As eUnicornCurveType
    CurveType = pCurveType
End Property

Public Property Let CurveType(CT As eUnicornCurveType)
    pCurveType = CT
End Property

Private Property Let DataType(ByVal InpVar As eUnicornDataType)
    pDataType = InpVar
End Property

Public Property Get DataType() As eUnicornDataType
    DataType = pDataType
End Property

Public Property Get MagicID() As String
    MagicID = pMagicID
End Property

Public Property Let MagicID(ByVal InpVarMagicID As String)
    pMagicID = InpVarMagicID
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal InpVarName As String)
    pName = InpVarName
End Property

Public Property Get DataSize() As Long
    DataSize = pDataSize
End Property

Public Property Let DataSize(ByVal InpVarDataSize As Long)
    pDataSize = InpVarDataSize
End Property

Public Property Get DataOffsetToNext() As Long
    DataOffsetToNext = pDataOffsetToNext
End Property

Public Property Let DataOffsetToNext(ByVal InpVarDataOffsetToNext As Long)
    pDataOffsetToNext = InpVarDataOffsetToNext
End Property

Public Property Get DataAddress() As Long
    DataAddress = pDataAddress
End Property

Public Property Let DataAddress(ByVal InpVarDataAddress As Long)
    pDataAddress = InpVarDataAddress
End Property

Public Property Get OffsetMetaToData() As Long
    OffsetMetaToData = pOffsetMetaToData
End Property

Public Property Let OffsetMetaToData(ByVal InpVarOffsetMetaToData As Long)
    pOffsetMetaToData = InpVarOffsetMetaToData
End Property

Private Sub DefineCurveType()

    Const conUV As String = "UV*"
    Const conCond As String = "Cond"
    Const conFlow As String = "*Flow*"
    Const conPressure As String = "*Press*"
    Const conConc As String = "Conc"
    Const conpH As String = "pH"
    Const conTemp As String = "Temp"
    
    Dim colon As Long
    Dim Underscore As Long
    
    'Names are of the format [Result]:[Chromatogram]_[Curve]
    'Result and Chromatogram can't have : or _ (nor space)
    '
    colon = InStr(1, Name, ":")
    Underscore = InStr(colon, Name, "_")
        
    ResName = Left(Name, colon - 1)
    ChromName = Mid(Name, colon + 1, Underscore - colon - 1)
    CurveName = Right(Name, Len(Name) - Underscore)
        
    
    If CurveName Like conUV Then
        
        CurveType = uniUV
    
    ElseIf CurveName Like conPressure Then
    
        CurveType = uniPressure
    
    Else
    
        Select Case CurveName
            Case conCond
                CurveType = uniCond
            Case conFlow
                CurveType = uniFlow
            Case conConc
                CurveType = uniConc
            Case conpH
                CurveType = uniOther
            Case conTemp
                CurveType = uniOther
            Case Else
                CurveType = uniOther
        End Select
            
    End If

End Sub

Private Sub DefineType()

    Dim RegEx As RegExp
    
    Set RegEx = New RegExp
    
    'initialize
    DataType = uniMeta
    
    'test if it's some other data type
    
    RegEx.Pattern = conHeaderNotConst_DataCurve_RES
    If RegEx.test(MagicID) Then DataType = uniData
    
    RegEx.Pattern = conHeaderNotConst_ImportedCurve_RES
    If RegEx.test(MagicID) Then DataType = uniData
    
    RegEx.Pattern = conHeaderNotConst_ChromatogramLogBook_RES
    If RegEx.test(MagicID) Then DataType = uniLogBook
    
    RegEx.Pattern = conHeaderInject_RES
    If RegEx.test(MagicID) Then DataType = uniInjection
    
    RegEx.Pattern = conHeaderFractions_RES
    If RegEx.test(MagicID) Then DataType = uniFraction
    
    Set RegEx = Nothing
    
    
End Sub

Public Sub ImportMagicID(ByVal InpVarImportMagicID As String)

    Const conByteLength As Long = conHeaderMagicID_L
    Const conName As String = "Header Magic ID"
    
    If Len(InpVarImportMagicID) <> conByteLength Then
        Call Err.Raise(1001, "clsUnicorn3Header", _
            conName & " must be " & _
            conByteLength & " bytes long! " & vbCrLf & _
            "(input: " & AsHex(InpVarImportMagicID) & ")")
    End If
        
    MagicID = AsHex(InpVarImportMagicID)
    
    DefineType
    
End Sub

Public Sub ImportName(ByVal InpVarImportName As String)
    
    Const conByteLength As Long = conHeaderName_L
    Const conName As String = "Header Name"
    
    Dim i As Long
    
    If Len(InpVarImportName) <> conByteLength Then
        Call Err.Raise(1001, "clsUnicorn3Header", _
            conName & " must be " & _
            conByteLength & " bytes long! " & vbCrLf & _
            "(input: " & Left(AsHex(InpVarImportName), 3 * 8) & "(...))")
    End If
    
    i = 0
    
    Do While Asc(Mid(InpVarImportName, conByteLength - i, 1)) = 0 And i < (conByteLength - 1)
        i = i + 1
    Loop
           
    Name = Left(InpVarImportName, conByteLength - i)
    
    If Len(Name) = 0 Then Call ErrorReport(, conName & " cannot be blank!")
    
End Sub

Public Sub ImportDataSize(ByVal InpVarImportDataSize As String)
   
       Const conByteLength As Long = conDataSize_L
    Const conName As String = "Data Size"
    
    If Len(InpVarImportDataSize) <> conByteLength Then
        Call Err.Raise(1001, "clsUnicorn3Header", _
            conName & " must be " & _
            conByteLength & " bytes long! " & vbCrLf & _
            "(input: " & AsHex(InpVarImportDataSize) & ")")
    End If
        
    DataSize = BinaryStringToInt32(InpVarImportDataSize, LittleEndian:=False)
    
    If DataSize < 0 Then Call ErrorReport(, conName & " cannot be less than zero!")
   
End Sub

Public Sub ImportDataOffsetToNext(ByVal InpVarImportDataOffsetToNext As String)
    
    Const conByteLength As Long = conDataOffsetToNext_L
    Const conName As String = "Data Offset To Next"
    
    If Len(InpVarImportDataOffsetToNext) <> conByteLength Then
        Call Err.Raise(1001, "clsUnicorn3Header", _
            conName & " must be " & _
            conByteLength & " bytes long! " & vbCrLf & _
            "(input: " & AsHex(InpVarImportDataOffsetToNext) & ")")
    End If
        
    DataOffsetToNext = BinaryStringToInt32(InpVarImportDataOffsetToNext, LittleEndian:=False)
            
    If DataOffsetToNext < 0 Then Call ErrorReport(, conName & " cannot be less than zero!")
            
End Sub

Public Sub ImportDataAddress(ByVal InpVarImportDataAddress As String)

    Const conByteLength As Long = conDataAddress_L
    Const conName As String = "Data Address"
    
    If Len(InpVarImportDataAddress) <> conByteLength Then
        Call Err.Raise(1001, "clsUnicorn3Header", _
            conName & " must be " & _
            conByteLength & " bytes long! " & vbCrLf & _
            "(input: " & AsHex(InpVarImportDataAddress) & ")")
    End If
        
    DataAddress = BinaryStringToInt32(InpVarImportDataAddress, LittleEndian:=False)
    
    If DataAddress < 0 Then Call ErrorReport(, conName & " cannot be less than zero!")
    
End Sub

Public Sub ImportOffsetMetaToData(ByVal InpVarImportOffsetMetaToData As String)

    Const conByteLength As Long = conMetaToData_L
    Const conName As String = "Meta-to-Data Offset"
    
    If Len(InpVarImportOffsetMetaToData) <> conByteLength Then
        Call Err.Raise(1001, "clsUnicorn3Header", _
            conName & " must be " & _
            conByteLength & " bytes long! " & vbCrLf & _
            "(input: " & AsHex(InpVarImportOffsetMetaToData) & ")")
    End If
        
    OffsetMetaToData = BinaryStringToInt32(InpVarImportOffsetMetaToData, LittleEndian:=False)
    
    If OffsetMetaToData < 0 Then Call ErrorReport(, conName & " cannot be less than zero!")
    
End Sub

Public Sub ImportHeader(ByVal RawHeader As String)
    
    Dim HeaderLength As Long
        
    HeaderLength = conHeaderTotal_L
    
    If Len(RawHeader) <> HeaderLength Then
        Call ErrorReport(, "Header length must be exactly " & HeaderLength & " bytes long!")
    Else
    
        Call ImportMagicID(Mid(RawHeader, _
                                conHeaderMagicID_IntPos, _
                                conHeaderMagicID_L))
                                
        Call ImportName(Mid(RawHeader, _
                                conHeaderName_IntPos, _
                                conHeaderName_L))
                                
        Call ImportDataSize(Mid(RawHeader, _
                                conDataSize_IntPos, _
                                conDataSize_L))
                                
        Call ImportDataOffsetToNext(Mid(RawHeader, _
                                conDataOffsetToNext_IntPos, _
                                conDataOffsetToNext_L))
                                
        Call ImportDataAddress(Mid(RawHeader, _
                                conDataAddress_IntPos, _
                                conDataAddress_L))
                                
        Call ImportOffsetMetaToData(Mid(RawHeader, _
                                conMetaToData_IntPos, _
                                conMetaToData_L))
        
        If DataType = uniData Then DefineCurveType
                                
    End If
        
        

End Sub


Private Sub ErrorReport(Optional ByVal ErrorNumber As Long, Optional ByVal StringMessage As String)
    
    Const conDefaultErrorN As Long = 1001
    Const conDefaultError As String = "An undocumented error has occured."
    Const conSource As String = "clsUnicorn3Header"
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    If Len(StringMessage) = 0 Then StringMessage = conDefaultError
    
    Call Err.Raise(ErrorNumber, conSource, StringMessage)

End Sub
