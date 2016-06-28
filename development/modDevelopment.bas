Attribute VB_Name = "modDevelopment"
Option Explicit

Sub Datatest()

Dim inp
Dim inp1()
Dim out
Dim out1
Dim i

inp = Array(0.24, 0.3, 0.47, -0.33, -0.27, 0.44, 0.11, 0.27, 0.09, -0.48, 9.5, -0.32, -0.33, -0.01, -0.21, -0.44, -0.07, -0.06, 4.5, -0.42, -0.17, 0.36, -0.38, 0.44, -0.3, 0.16, -0.2, 9.5, 0.01, 9.5, 0.07, 9.5, 0.24, 19.5, -0.29, -0.1, -0.08, -0.23, 0.22, 0.42, -0.02, 0.49, 19.5, -0.3, 0.12, -0#, 0.28, -0.13, -0.26, 0.47, -0.22, -0.33, 0.14, 4.5, -0.24, -0.02, -0.12, 0.26, 0.02, 0.29)

ReDim inp1(1 To UBound(inp) + 1)

For i = 1 To UBound(inp) + 1
    inp1(i) = inp(i - 1)
Next i

out = SmoothData(inp1, 5)

For i = LBound(out) To UBound(out)
    out1 = out1 & out(i) & "#"
Next i
    

End Sub

Sub testa()

Dim a As String

Dim b
Dim C

a = ">prvi" & vbCrLf & "MAGNSTV" & vbCrLf
a = a & ">drugi" & vbCrLf & "SAL-AKA" & vbCrLf

b = mFastaToArrayOfFasta(a, , True)


C = CalculateConservationScore(b)

End Sub

Function InputFileToString(Filename As String) As String

Dim strFileContent As String
Dim iFile As Long: iFile = FreeFile

Open Filename For Input As #iFile
    strFileContent = Input(LOF(iFile), iFile)
Close #iFile

InputFileToString = strFileContent

End Function

Sub CalcConservation()

Dim strFilename As String
Dim InputString As String

Dim i As Long

Dim tempOutput As String

strFilename = InputBox("InputFileName", "FileName", "E:\TEMP 2015\Computer Backup 20151012\Desktop\Mysterin\Mysterins_sel4.a2m_mouseontop.fa")

InputString = InputFileToString(strFilename)

Dim ConservationScore As Variant

ConservationScore = CalculateConservationScore(mFastaToArrayOfFasta(InputFileToString(strFilename), "UPPER", True), 1, 100)

For i = LBound(ConservationScore) To UBound(ConservationScore)
    tempOutput = tempOutput & ConservationScore(i) & "#"
Next i

Call ExportDataToTextFile(tempOutput, "C:\temp\20160412test.txt")


End Sub


Function SmoothData(DataArray As Variant, WindowSize As Long) As Variant

    Dim DataLength As Long
    Dim tempOutput() As Variant
    Dim tempIndex As Long
    Dim i As Long, j As Long
    Dim TempSum As Double
        
        
    DataLength = 1 + UBound(DataArray) - LBound(DataArray)
    
    ReDim tempOutput(1 To DataLength)
        
    
            
    For i = 1 To DataLength - WindowSize
        
        tempIndex = i + WindowSize \ 2
        TempSum = 0
        
        For j = i To i + WindowSize - 1
            TempSum = TempSum + DataArray(j)
        Next j
        
        tempOutput(tempIndex) = TempSum / WindowSize
        
    Next i
        
        
    For i = 1 To WindowSize \ 2
        tempOutput(i) = 0
        tempOutput(DataLength - i + 1) = 0
    Next i
    
    SmoothData = tempOutput

End Function

Sub testaaa()

Dim a

a = GetMaxLetterCount("aaa")

End Sub

Function GetMaxLetterCount(InputString As String) As Long

    Dim i As Byte
    Dim Char As String
    Dim tout As String
    Dim tempCount As Long
    Dim CharCount As Long
    
    tempCount = 0
    
    For i = 65 To 90
    
        Char = Chr(i)
        CharCount = StringCharCount(InputString, Char)
        If CharCount > tempCount Then tempCount = CharCount
        
    Next i
    
    GetMaxLetterCount = tempCount
        

End Function


Function CalculateConservationScore(FASTAArray As Variant, _
                                    Optional ReferenceSequence As Long = 1, _
                                    Optional Smoothing As Long = 1) _
                                    As Variant
    
'this calculation is too basic. Try to find documentation about AACon from JalView, which uses
'physicochemical properties
    
    Const conSpacer = "-"
    
    Dim SequenceNumber As Long
    Dim SequenceLength As Long
    Dim ReferenceLength As Long
    Dim tempString As String
    
    Dim i As Long, j As Long
    
    Dim ConservationArray() As Long
    Dim ReferenceSeqArray() As String
    Dim OutputArray() As Long
    Dim tempSmooth
        
    SequenceNumber = 1 + UBound(FASTAArray, 1) - LBound(FASTAArray, 1)
    SequenceLength = Len(FASTAArray(1, 2))
    ReferenceLength = Len(Replace(FASTAArray(ReferenceSequence, 2), conSpacer, ""))
    
    ReDim ConservationArray(1 To SequenceLength)
    ReDim ReferenceSeqArray(1 To SequenceLength)
    
    For i = 1 To SequenceLength
        
        ReferenceSeqArray(i) = Mid(FASTAArray(ReferenceSequence, 2), i, 1)
        tempString = ""
        
        For j = 1 To SequenceNumber
            tempString = tempString & Mid(FASTAArray(j, 2), i, 1)
        Next j
        
        ConservationArray(i) = GetMaxLetterCount(tempString)
        
    Next i
    
    ReDim OutputArray(1 To ReferenceLength)
    
    If Smoothing > 1 Then
        
        tempSmooth = SmoothData(ConservationArray, Smoothing)
    
    Else
    
        tempSmooth = ConservationArray
    
    End If
        
    j = 0
    For i = 1 To SequenceLength
        If ReferenceSeqArray(i) <> conSpacer Then
            j = j + 1
            OutputArray(j) = tempSmooth(i)
        End If
    Next i
            
           
    CalculateConservationScore = OutputArray
    
    
End Function



'****************************************************************************************************
Function mFastaToArrayOfFasta(FASTASequence As String, _
                                Optional SequenceCase As String = "UPPER", _
                                Optional Alignment As Boolean = False) As Variant

'====================================================================================================
'Takes a FASTA or mFASTA text as input and extracts the headers and the sequences to a 2D array
'Juraj Ahel, 2016-04-12, for general purposes
'Last update 2016-04-12
'====================================================================================================

    Dim SStart As Long, SEnd As Long, HStart As Long, HEnd As Long
    Dim SequenceNumber As Long
    Dim i As Long, j As Long
    
    Dim LineTerminator As String: LineTerminator = Chr(10)
    Dim DoubleLineTerminator As String: DoubleLineTerminator = LineTerminator & LineTerminator
    
    Dim ForbiddenSymbols As Variant
    
    Dim ErrorMessage As String
    
    Dim tempOutput() As String
    
    'Symbols that will be removed from sequence data - they are ok in headers'
    
    'Normal Fasta
    'ForbiddenSymbols = Array( _
    '                         Chr(9), Chr(124), LineTerminator, " ", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
    '                         "-", "*", ":", ";", "'", """", "#", "@", "&", "/", "\", "^", "_", "+", "?", "!", "$", _
    '                         "%", "=", "[", "]", "(", ")", "{", "}" _
    '                        )
    
    'AlignedFasta - do not remove "-"
    
    If Alignment Then
        ForbiddenSymbols = Array( _
                                 Chr(9), Chr(124), LineTerminator, " ", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                                 "*", ":", ";", "'", """", "#", "@", "&", "/", "\", "^", "_", "+", "?", "!", "$", _
                                 "%", "=", "[", "]", "(", ")", "{", "}" _
                                )
    Else
        ForbiddenSymbols = Array( _
                                 Chr(9), Chr(124), LineTerminator, " ", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                                 "*", ":", ";", "'", """", "#", "@", "&", "/", "\", "^", "_", "+", "?", "!", "$", _
                                 "%", "=", "[", "]", "(", ")", "{", "}", _
                                 "-")
    End If
    
                                
    SequenceNumber = StringCharCount(FASTASequence, ">")
    
    ReDim tempOutput(1 To SequenceNumber, 1 To 2)
    
    FASTASequence = Replace(FASTASequence, Chr(13), LineTerminator)
    FASTASequence = Replace(FASTASequence, Chr(10), LineTerminator)
    FASTASequence = FASTASequence & LineTerminator & ">" 'to allow termination for the final sequence
    
    Do While StringCharCount(FASTASequence, DoubleLineTerminator) > 0
        FASTASequence = Replace(FASTASequence, DoubleLineTerminator, LineTerminator)
    Loop
    
    HStart = 1: HEnd = 1
    SStart = 1: SEnd = 1
        
    For i = 1 To SequenceNumber
        
        HStart = InStr(SEnd, FASTASequence, ">", vbBinaryCompare) + 1
        HEnd = InStr(HStart, FASTASequence, LineTerminator, vbBinaryCompare) - 1
        SStart = HEnd + 2
        SEnd = InStr(SStart, FASTASequence, ">", vbBinaryCompare) - 2
        
        If HEnd > HStart Then tempOutput(i, 1) = Mid(FASTASequence, HStart, HEnd - HStart + 1) Else tempOutput(i, 1) = "[EMPTY_HEADER]"
        If SEnd > SStart Then tempOutput(i, 2) = Mid(FASTASequence, SStart, SEnd - SStart + 1) Else tempOutput(i, 2) = ""
        
        For j = LBound(ForbiddenSymbols) To UBound(ForbiddenSymbols)
            tempOutput(i, 2) = Replace(tempOutput(i, 2), ForbiddenSymbols(j), "")
        Next j
        
    Next i
        
    'Change case, as per settings. UPPERCASE is the default.
    
    Select Case SequenceCase
        Case "UPPER"
            For i = 1 To SequenceNumber: tempOutput(i, 2) = UCase(tempOutput(i, 2)): Next i
        Case "lower"
            For i = 1 To SequenceNumber: tempOutput(i, 2) = LCase(tempOutput(i, 2)): Next i
        Case "Preserve"
        Case Else
            For i = 1 To SequenceNumber: tempOutput(i, 2) = UCase(tempOutput(i, 2)): Next i
    End Select
    
    'in alignments, total length of each sequence must be the same
    If Alignment Then
        
        j = Len(tempOutput(1, 2))
        
        For i = 2 To SequenceNumber
            
            If Len(tempOutput(i, 2)) <> j Then
                ErrorMessage = "Sequence #" _
                                & i _
                                & " not equal in length to #1 (" _
                                & Len(tempOutput(i, 2)) _
                                & " vs " _
                                & j _
                                & "). Check input file!"
                Call Err.Raise(13, "mFastaToArrayOfFasta", ErrorMessage)
            End If
                        
        Next i
        
    End If
    
   
    mFastaToArrayOfFasta = tempOutput
    
End Function

'****************************************************************************************************
Sub ExportStringToTXT(InputString)

'====================================================================================================
'Exports a separate text file for each cell in selection
'Still needs to be added modular file naming, now it's always "Fragment #.txt"
'
'Juraj Ahel, 2015-02-10, for Gibson assembly and general purposes
'Last update 2015-02-10
'====================================================================================================

Dim FilePath As String, OutputFile As String
Dim DataSource As Range

FilePath = "C:\Excel_outputs\"



    OutputFile = FilePath & "OutputTemp.txt"
    Call ExportDataToTextFile(CStr(InputString), OutputFile)


End Sub
'****************************************************************************************************
Function ReindexFromAlignment(inp1, Inp2)

Const AllowedLetters = "[ACDEFGHIKLMNPQRSTVWY]"

Dim len1, len2
Dim i
Dim a1(), a2(), a3()
len1 = Len(inp1)

ReDim a1(1 To len1)
ReDim a2(1 To len1)

c1 = 1
c2 = 1

For i = 1 To len1

    If Mid(inp1, i, 1) Like AllowedLetters Then
        a1(i) = c1
        c1 = c1 + 1
    Else
        a1(i) = ""
    End If
    
    If Mid(Inp2, i, 1) Like AllowedLetters Then
        a2(i) = c2
        c2 = c2 + 1
    Else
        a2(i) = ""
    End If
    
Next i

For i = 1 To len1
    If a1(i) Like "[1-9][0-9]*" Then
        OP = OP & vbCrLf & a1(i) & vbTab
        If a2(i) Like "##*" Then
            OP = OP & a2(i)
        Else
            OP = OP & "%"
        End If
    End If
Next i

aaa = Len(OP)
ReindexFromAlignment = OP
Call ExportStringToTXT(OP)


End Function

