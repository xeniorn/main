Attribute VB_Name = "XmodImportFasta"
'****************************************************************************************************
Sub ImportFasta()
Attribute ImportFasta.VB_ProcData.VB_Invoke_Func = "F\n14"

'====================================================================================================
'Takes a FASTA or mFASTA text as input (pasted in) and extracts the headers and the sequences
'and prints them out below the selected cell (starting from the cell down, 1 cell 1 sequence)
'headers, if output, are in the next column
'Juraj Ahel, 2015-02-15, for general purposes
'Last update 2015-02-15
'====================================================================================================
'Need to make this InputForm be created within the macro (or using a side macro)
'So that this code can easily be copied to other workbooks without re-creating
'the user form manually

'also add functionality where you can pick a file instead of copying the text in!

'also add a choice on which side the headers are output!

Dim FASTASequence As String
Dim Sequences() As String, Headers() As String
Dim SStart As Long, SEnd As Long, HStart As Long, HEnd As Long
Dim SequenceNumber As Long
Dim i As Long

Dim LineTerminator As String: LineTerminator = Chr(10)
Dim DoubleLineTerminator As String: DoubleLineTerminator = LineTerminator & LineTerminator

Dim OutputRangeSequences As Range, OutputRangeHeaders As Range
Dim ForbiddenSymbols As Variant

Dim DefaultInput As String

Dim OutputHeaders As Boolean, OutputSequences As Boolean
Dim SequenceCase As String

'Symbols that will be removed from sequence data - they are ok in headers'

'Normal Fasta
'ForbiddenSymbols = Array( _
'                         Chr(9), Chr(124), LineTerminator, " ", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
'                         "-", "*", ":", ";", "'", """", "#", "@", "&", "/", "\", "^", "_", "+", "?", "!", "$", _
'                         "%", "=", "[", "]", "(", ")", "{", "}" _
'                        )

'AlignedFasta - do not remove "-"
ForbiddenSymbols = Array( _
                         Chr(9), Chr(124), LineTerminator, " ", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                         "*", ":", ";", "'", """", "#", "@", "&", "/", "\", "^", "_", "+", "?", "!", "$", _
                         "%", "=", "[", "]", "(", ")", "{", "}" _
                        )

DefaultInput = ">Seq1                                  " & vbCrLf & _
               "ATCTGACGAGCGAGCGTAGCTAGTCGATGCTACCG    " & vbCrLf & _
               "AGCGCACGTATTCCCGCGATCGCGATATTTGCGTCA   " & vbCrLf & _
               "ACACGGGTTTTTGCCCCAATCGCCGTCGATATCGC    " & vbCrLf & _
               ">Seq2                                  " & vbCrLf & _
               "ATCTGACGAGCGAGCGTAGCTAGTCGATGCTACCG    " & vbCrLf & _
               "AGCGCACGTATTCCCGCGATCGCGATATTTGCGTCA   " & vbCrLf & _
               "ACACGGGTTTTTGCCCCAATCGCCGTCGATATCGC    " & vbCrLf

'Need to make this InputForm be created within the macro (or using a side macro)
'So that this code can easily be copied to other workbooks without re-creating
'the user form manually

'Invoke Input Form:
FASTAInputForm.TextBox.Text = DefaultInput
FASTAInputForm.Show

'Take values from the input form:
FASTASequence = FASTAInputForm.TextBox.Text

'Copy this from the user form code
Const CaptionA = "Output Sequences Only"
Const CaptionB = "Output Sequences and Headers"
Const CaptionC = "Output Headers Only"

Const CaptionD = "Output UPPERCASE"
Const CaptionE = "Output lowercase"
Const CaptionF = "Preserve case"

Select Case FASTAInputForm.B_OutputType.Caption
    Case CaptionA: OutputSequences = True: OutputHeaders = False
    Case CaptionB: OutputSequences = True: OutputHeaders = True
    Case CaptionC: OutputSequences = False: OutputHeaders = True
    Case Else: OutputSequences = True: OutputHeaders = False
End Select

Select Case FASTAInputForm.B_OutputCase.Caption
    Case CaptionD: SequenceCase = "UPPER"
    Case CaptionE: SequenceCase = "lower"
    Case CaptionF: SequenceCase = "Preserve"
    Case Else: SequenceCase = "UPPER"
End Select
        
'So far it was just hidden, now unload it (destroy its instance), for cleanliness
Unload FASTAInputForm

'TODO: Check if it is proper FASTA here...'

SequenceNumber = StringCharCount(FASTASequence, ">")
ReDim Sequences(1 To SequenceNumber, 1 To 1)
ReDim Headers(1 To SequenceNumber, 1 To 1)

'Warn user that cell contents will be relpaced:'            '#Make a check for this, whether the cells are actually empty -
                                                            'don't scare the user if it's not necessary
ProceedCheck = MsgBox( _
                      "Input will replace the contents of the cells at and/or under the current selection." & vbCrLf & _
                      "Any contents will be lost. Proceed?", _
                      vbYesNo, _
                      "Confirm replacement:" _
                     )

If ProceedCheck = vbNo Then
    MsgBox ("Input aborted.")
Else
    
    
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
        
        If HEnd > HStart Then Headers(i, 1) = Mid(FASTASequence, HStart, HEnd - HStart + 1) Else Headers(i, 1) = "[EMPTY_HEADER]"
        If SEnd > SStart Then Sequences(i, 1) = Mid(FASTASequence, SStart, SEnd - SStart + 1) Else Sequences(i, 1) = ""
        
        For j = LBound(ForbiddenSymbols) To UBound(ForbiddenSymbols)
            Sequences(i, 1) = Replace(Sequences(i, 1), ForbiddenSymbols(j), "")
        Next j
        
    Next i
        
    'Change case, as per settings. UPPERCASE is the default.
    
    Select Case SequenceCase
        Case "UPPER"
            For i = 1 To SequenceNumber: Sequences(i, 1) = UCase(Sequences(i, 1)): Next i
        Case "lower"
            For i = 1 To SequenceNumber: Sequences(i, 1) = LCase(Sequences(i, 1)): Next i
        Case "Preserve"
        Case Else
            For i = 1 To SequenceNumber: Sequences(i, 1) = UCase(Sequences(i, 1)): Next i
    End Select
    
    'Output the sequences to worksheet
    If OutputSequences Then
        Set OutputRangeSequences = Selection.Cells(1, 1).Resize(SequenceNumber, 1)
        OutputRangeSequences.Value = Sequences
    End If
    
    'OutputHeadersToggle = MsgBox("Output headers as well?", vbYesNo + vbDefaultButton2)
    'If OutputHeadersToggle = vbYes Then Outputheaders = True Else Outputheaders = False
    
    'Output the headers to worksheet
    If OutputHeaders Then
        Set OutputRangeHeaders = Selection.Cells(1, 1).Resize(SequenceNumber, 1).Offset(0, -1)
        OutputRangeHeaders.Value = Headers
    End If
        
End If


    
End Sub
