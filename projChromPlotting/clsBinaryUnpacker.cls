VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBinaryUnpacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2016-06-08, for this and that purpose
'Last update 2016-06-10
'====================================================================================================
Option Explicit

Const conClassName As String = "clsBinaryUnpacker"

Private Const sizeLong As Long = 4
Private Const sizeInteger As Long = 2
Private Const sizeString As Long = 1
Private Const sizeDouble As Long = 8
Private Const sizeSingle As Long = 4
Private Const SizeByte As Long = 1

Private pStructureCollection As VBA.Collection
Private DescriptionCollection As VBA.Collection

Private IsLittleEndian As Boolean

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    Set pStructureCollection = Nothing
End Sub

Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub


Public Function UnpackBinaryData( _
    ByVal FormatString As String, _
    ByVal SourceDataBuffer As String, _
    Optional ByVal StartByteOffset As Long = 0, _
    Optional ByVal BytesToProcess As Long = 0, _
    Optional ByVal IgnoreExtraBytes As Boolean = False _
    ) As VBA.Collection
   
    Dim tempDataList As clsDataList
    
    Dim InputDataLength As Long
    Dim StructureLengthSum As Long
    Dim NumberOfStructures As Long
    
    Dim tempType As String
    Dim tempSubString As String
    
    Dim i As Long, j As Long
    
    Set DescriptionCollection = ParseFormatString(FormatString)
    
    If Len(SourceDataBuffer) = 0 Then
        ErrorReport , "InputData cannot be empty"
    End If
    
    'correct the SourceDataBuffer so that it starts from desired beginning, truncate rest
    SourceDataBuffer = Right(SourceDataBuffer, Len(SourceDataBuffer) - StartByteOffset)
    InputDataLength = Len(SourceDataBuffer)
    
    'calculate structure tuple size
        StructureLengthSum = 0
        For i = 1 To DescriptionCollection.Count
            Select Case DescriptionCollection.Item(i)
                Case "BOOL"
                    StructureLengthSum = StructureLengthSum + 1
                Case "INT16", "UINT16"
                    StructureLengthSum = StructureLengthSum + sizeInteger
                Case "INT32", "UINT32", "SINGLE"
                    StructureLengthSum = StructureLengthSum + sizeLong
                Case "DOUBLE"
                    StructureLengthSum = StructureLengthSum + sizeDouble
                Case Else
                    If DescriptionCollection.Item(i) Like "STRING?*" Then
                        StructureLengthSum = StructureLengthSum + _
                            CLng(Right(DescriptionCollection.Item(i), Len(DescriptionCollection.Item(i)) - 6))
                    Else
                        ErrorReport , "This should have never happened"
                    End If
            End Select
        Next i
        
    'calculate the max number of structures that fit inside
        NumberOfStructures = InputDataLength \ StructureLengthSum
        
    'if extra bytes are not ignored, all data must be consumed (therefore byte length needs to be multiple of single tuple length)
        If Not IgnoreExtraBytes Then
            
            If InputDataLength Mod StructureLengthSum <> 0 Then
                ErrorReport , "Chosen binary input length is not a multiple of the defined structure"
            End If
        
    'otherwise, data must be truncated so that extra characters are removed, due to how it is programmed
        Else
        
            SourceDataBuffer = Left(SourceDataBuffer, NumberOfStructures * StructureLengthSum)
        
        End If
    
        
        
    Set pStructureCollection = New VBA.Collection
    
    'add DataList objects to the collection into which the structure will be decomposed
        For i = 1 To DescriptionCollection.Count
            
            Set tempDataList = New clsDataList
            tempType = DescriptionCollection.Item(i)
            
            Select Case tempType
                Case "INT32"
                    tempDataList.SetType (vbLong)
                Case "SINGLE"
                    tempDataList.SetType (vbSingle)
                Case "DOUBLE"
                    tempDataList.SetType (vbDouble)
                Case "BOOL"
                    tempDataList.SetType (vbBoolean)
                Case Else
                    If tempType Like "STRING?*" Then
                        tempDataList.SetType (vbString)
                    Else
                        ErrorReport , "Option not valid or not yet implemented: " & tempType
                    End If
            End Select
            
            tempDataList.SetSize (NumberOfStructures)
            
            pStructureCollection.Add tempDataList
            
        Next i
    
    'extract the actual data to individual items in the collection!
        For i = 1 To NumberOfStructures
            tempSubString = Mid(SourceDataBuffer, 1 + (i - 1) * StructureLengthSum, StructureLengthSum)
            Call AppendStructure(tempSubString, i)
        Next i
    
    Set UnpackBinaryData = pStructureCollection
        
End Function




Private Sub AppendStructure( _
    ByVal SourceString As String, _
    ByVal TargetIndex As Long)
        
    Dim i As Long
    Dim tempDataList As clsDataList
    Dim tempVarType As String
    Dim Buffer As String
    Dim CurrentPosition As Long
    Dim tempStringSize As Long
    
    CurrentPosition = 1
    IsLittleEndian = True
    
    For i = 1 To DescriptionCollection.Count
        
        Set tempDataList = pStructureCollection.Item(i)
        tempVarType = DescriptionCollection.Item(i)
        
        Select Case tempVarType
            Case "INT32"
                Buffer = Mid(SourceString, CurrentPosition, sizeLong)
                CurrentPosition = CurrentPosition + sizeLong
                tempDataList.Data(TargetIndex) = BinaryStringToInt32(Buffer, IsLittleEndian)
            Case "DOUBLE"
                Buffer = Mid(SourceString, CurrentPosition, sizeDouble)
                CurrentPosition = CurrentPosition + sizeDouble
                tempDataList.Data(TargetIndex) = BinaryStringToDouble(Buffer, IsLittleEndian)
            Case "SINGLE"
                Buffer = Mid(SourceString, CurrentPosition, sizeSingle)
                CurrentPosition = CurrentPosition + sizeSingle
                tempDataList.Data(TargetIndex) = BinaryStringToSingle(Buffer, IsLittleEndian)
            Case Else
                If tempVarType Like "STRING?*" Then
                    tempStringSize = CLng(Right(tempVarType, Len(tempVarType) - 6))
                    Buffer = Mid(SourceString, CurrentPosition, tempStringSize)
                    CurrentPosition = CurrentPosition + tempStringSize
                    tempDataList.Data(TargetIndex) = Buffer
                Else
                    ErrorReport
                End If
        End Select
    
    Next i

End Sub

Private Function ParseFormatString( _
    ByVal InputString As String _
    ) As VBA.Collection

'like python struct:
'c char (string of length 1)
'b signed char
'B unsigned char
'? Boolean
'h short (Integer)
'H unsigned short
'i,l int (Long)
'I,L unsigned int
'q long long
'Q unsigned long long
'f float (Single)
'd double (Double)
's,p char[] (String)
'
'e.g. 164s2d is a structure composed of string(164)-double-double triplets

    Const AllowedPattern As String = "[cbB?hHiIlLqQfdsp0-9]"
    Const ImplementedPattern As String = "[hHiIlLfdsp0-9]"
    Const ImplementedPatternLastPosition As String = "[hHiIlLfdsp]"
    Const Prefixes As String = "[@=<>!]"
    
    Dim tempCollection As VBA.Collection
    Dim i As Long
    Dim currentIndex As Long
    Dim tempNumberString As String
    Dim currentChar As String
    Dim tempCount As Long
    Dim tempVarType As String
    
    If Not (Left(InputString, 1) Like Prefixes) Then
        InputString = "@" & InputString
    End If
    
    'decide on the endinanness of data
    Select Case Left(InputString, 1)
        Case "@", "=", "<"
            IsLittleEndian = True
        Case ">", "!"
            IsLittleEndian = False
        Case Else
            ErrorReport , "Unrecognized prefix"
    End Select
        
    
    'must be at least 2 after parsing (prefix + at least 1 data)
    If Len(InputString) < 2 Then
        ErrorReport ("Invalid format string")
    End If
    
    For i = 2 To Len(InputString)
        If Not (Mid(InputString, i, 1) Like AllowedPattern) Then
            Debug.Print ("unsupported character in format string: " & Mid(InputString, i, 1))
        ElseIf Not (Mid(InputString, i, 1)) Like ImplementedPattern Then
            Debug.Print ("unimplemented character in format string: " & Mid(InputString, i, 1))
        End If
    Next i
    
    If Not (Right(InputString, 1) Like ImplementedPatternLastPosition) Then
        Call ErrorReport(, "Format String cannot have a number as the last element")
    End If
    
    currentIndex = 1 'so that the prefix is skipped
    Set tempCollection = New VBA.Collection
    
    'for the whole input entry...
    Do While currentIndex < Len(InputString)
    
        currentIndex = currentIndex + 1
        currentChar = Mid(InputString, currentIndex, 1)
        tempNumberString = ""
        
        'determine the numeric prefix
        Do While IsNumeric(currentChar)
            tempNumberString = tempNumberString & currentChar
            currentIndex = currentIndex + 1
            currentChar = Mid(InputString, currentIndex, 1)
        Loop
        
        'if no numeric prefix, assume 1, otherwise use the numeric prefix
        If tempNumberString = "" Then
            tempCount = 1
        Else
            tempCount = CLng(tempNumberString)
        End If
        
        'if the structure member is a string (variable length)
        If (currentChar Like "[ps]") Then
            tempCollection.Add ("STRING" & tempCount)
        Else
            'add a separate entry the number of counted times of the variable type
            Select Case currentChar
                    Case "?": tempVarType = "BOOL"
                    Case "h": tempVarType = "INT16"
                    Case "H": tempVarType = "UINT16"
                    Case "i", "l": tempVarType = "INT32"
                    Case "I", "L": tempVarType = "UINT32"
                    Case "f": tempVarType = "SINGLE"
                    Case "d": tempVarType = "DOUBLE"
                    Case Else
                        Debug.Print ("this should never have happened")
            End Select
            
            For i = 1 To tempCount
                tempCollection.Add tempVarType
            Next i
        End If
        
    Loop 'for each structure entry...
    
    Set ParseFormatString = tempCollection
    
    Set tempCollection = Nothing

End Function


