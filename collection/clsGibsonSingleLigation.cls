VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGibsonSingleLigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2017-01-19, for this and that purpose
'Last update 2017-01-19
'====================================================================================================
Option Explicit

    Private Const conClassName As String = "clsGibsonSingleLigation"
    
    Private Const condefRNAFoldPath As String = "C:\ViennaRNA\"
    Private Const condefRNAFoldCommand = "RNAfold.exe"

'[Const Declaration]
    Private Const conTargetOverlapLength As Long = 15
    Private Const conTargetOverlapTm As Double = 48
    Private Const conMaxDistanceFromCenter As Long = 50
    Private Const conMaxOverlapLength As Long = 25

'[Var Declaration]
    Private RNAFoldPath As String
    Private RNAFoldCommand As String
    Private TempFilenamesRoot As String

    Private pFirstSequence As String
    Private pInsertedSequence As String
    Private pSecondSequence As String

    Private pForbiddenRegions As String
    
    Private pTargetOverlapLength As Long
    Private pTargetOverlapTm As Double
    Private pMaxDistanceFromCenter As Long
    Private pMaxOverlapLength As Long
    
    Private pOverlap As String
    Private pdG As Double
    Private pTm As Double
    
    Private IsCalculated As Boolean

'[Object Initialization]
Private Sub Class_Initialize()

    TargetOverlapLength = conTargetOverlapLength
    TargetOverlapTm = conTargetOverlapTm
    MaxDistanceFromCenter = conMaxDistanceFromCenter
    MaxOverlapLength = conMaxOverlapLength
    
    RNAFoldPath = condefRNAFoldPath
    RNAFoldCommand = condefRNAFoldCommand
        
    TempFilenamesRoot = FileSystem_GetTempFolder(True) & "JA_Gibson_"
        
    ResetCalculation

End Sub

'[Object Dereferencing]
Private Sub Class_Terminate()


End Sub

Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    'If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    If Len(ErrorString) = 0 Then
        Err.Raise vbError + ErrorNumber, conClassName, ErrorString
    Else
        Err.Raise vbError + ErrorNumber, conClassName
    End If

End Sub

'[Property Gets and Lets and Sets]
Public Property Get FirstSequence() As String
    FirstSequence = pFirstSequence
End Property

Public Property Let FirstSequence(ByVal inp_FirstSequence As String)
    pFirstSequence = inp_FirstSequence
    ResetCalculation
End Property

Public Property Get InsertedSequence() As String
    InsertedSequence = pInsertedSequence
End Property

Public Property Let InsertedSequence(ByVal inp_InsertedSequence As String)
    pInsertedSequence = inp_InsertedSequence
    ResetCalculation
End Property

Public Property Get SecondSequence() As String
    SecondSequence = pSecondSequence
End Property

Public Property Let SecondSequence(ByVal inp_SecondSequence As String)
    pSecondSequence = inp_SecondSequence
    ResetCalculation
End Property

Public Property Get ForbiddenRegions() As String
    ForbiddenRegions = pForbiddenRegions
End Property

Public Property Let ForbiddenRegions(ByVal inp_ForbiddenRegions As String)
    pForbiddenRegions = inp_ForbiddenRegions
    ResetCalculation
End Property

Public Property Get TargetOverlapLength() As Long
    TargetOverlapLength = pTargetOverlapLength
End Property

Public Property Let TargetOverlapLength(ByVal inp_TargetOverlapLength As Long)
    pTargetOverlapLength = inp_TargetOverlapLength
    ResetCalculation
End Property

Public Property Get TargetOverlapTm() As Double
    TargetOverlapTm = pTargetOverlapTm
End Property

Public Property Let TargetOverlapTm(ByVal inp_TargetOverlapTm As Double)
    pTargetOverlapTm = inp_TargetOverlapTm
    ResetCalculation
End Property

Public Property Get MaxDistanceFromCenter() As Long
    MaxDistanceFromCenter = pMaxDistanceFromCenter
End Property

Public Property Let MaxDistanceFromCenter(ByVal inp_MaxDistanceFromCenter As Long)
    pMaxDistanceFromCenter = inp_MaxDistanceFromCenter
    ResetCalculation
End Property

Public Property Get MaxOverlapLength() As Long
    MaxOverlapLength = pMaxOverlapLength
End Property

Public Property Let MaxOverlapLength(ByVal inp_MaxOverlapLength As Long)
    pMaxOverlapLength = inp_MaxOverlapLength
    ResetCalculation
End Property


Public Property Get FinalSequence() As String

    FinalSequence = FirstSequence & InsertedSequence & SecondSequence

End Property

Public Property Get Overlap() As String
    If Not IsCalculated Then
        PerformCalculation
    End If
    Overlap = pOverlap
End Property

Public Property Get dG() As Double
    If Not IsCalculated Then
        PerformCalculation
    End If
    dG = pdG
End Property

Public Property Get Tm() As Double
    If Not IsCalculated Then
        PerformCalculation
    End If
    Tm = pTm
End Property

'***************************************************************************************************
Private Sub ResetCalculation()
'====================================================================================================
'Undoes the calculations - e.g. if any parameter is changed
'Juraj Ahel, 2017-01-19
'====================================================================================================
'2017-01-19


    pTm = 0
    pdG = 0
    pOverlap = 0
    
    IsCalculated = False

End Sub

'***************************************************************************************************
Private Sub PerformCalculation()
'====================================================================================================
'Calculates the results if they haven't been calculated yet
'Juraj Ahel, 2017-01-19
'====================================================================================================
'2017-01-19

    Dim tColl As VBA.Collection
    
    Set tColl = OptimizeOverlaps

    'Collection of
        'OVERLAP as String      'overlap sequence
        'LEN as Long            'length of overlap
        'DG as Double           'Gibbs energy for SS formation of overlap
        'TM as Double           'melting temperature of overlap
        
    With tColl
        pOverlap = .Item("OVERLAP")
        pdG = .Item("DG")
        pTm = .Item("TM")
    End With
    
    IsCalculated = True
    
    Set tColl = Nothing

End Sub




'***************************************************************************************************
Private Function ConstructOverlaps() As VBA.Collection
'====================================================================================================
'Constructs the collection of overlaps to be used for Gibson
'Juraj Ahel, 2016-12-20, for Gibson assembly

'====================================================================================================
'2017-01-19 convert to class function
'2017-01-24 fix race condition that was coming up because output files were overlapping if the function was called too fast (naming based on YYYYMMDDhhmmss)

    Dim i As Long
    Dim MiddleIndex As Long
    
    Dim Len1 As Long, Len2 As Long, Len3 As Long
    Dim SearchStart As Long, SearchEnd As Long
    
    Dim CurrentLength As Long
    Dim CurrentStart As Long
    Dim CurrentEnd As Long
    Dim FinalLength As Long
        
    Dim OverlapAllowed As Boolean
    Dim tOverlap As String
    Dim FinalSequence As String
    
    Dim tempTm As Double
        
    Dim PutativeOverlaps As VBA.Collection
    Dim tColl As VBA.Collection
    
    Len1 = Len(FirstSequence)
    Len2 = Len(InsertedSequence)
    Len3 = Len(SecondSequence)
    
    FinalSequence = FirstSequence & InsertedSequence & SecondSequence
    
    FinalLength = Len1 + Len2 + Len3
    
    MiddleIndex = 1 + Len1 + Len2 \ 2
    
    Set PutativeOverlaps = New VBA.Collection
    'Sequence, start index'
    
    FinalSequence = FirstSequence & InsertedSequence & SecondSequence
    
    SearchStart = MiddleIndex - MaxDistanceFromCenter
    If SearchStart < 1 Then SearchStart = 1
    
    'test all allowed lengths
    For CurrentLength = TargetOverlapLength To MaxOverlapLength
        
        
        SearchEnd = MiddleIndex + MaxDistanceFromCenter - CurrentLength
        If SearchEnd > FinalLength - CurrentLength Then SearchEnd = FinalLength - CurrentLength
        
        'all allowed positions for a given length
        For CurrentStart = SearchStart To SearchEnd
            CurrentEnd = CurrentStart + CurrentLength - 1
        
            OverlapAllowed = True
            
            'if it isn't in any of the forbidden regions
                If OverlapAllowed Then
                    If InStr(1, ForbiddenRegions, "1") <> 0 Then
                        If RangeOverlaps(CurrentStart, CurrentEnd, 1, Len1) Then
                            OverlapAllowed = False
                        End If
                    End If
                End If
                
                If OverlapAllowed Then
                    If InStr(1, ForbiddenRegions, "2") <> 0 Then
                        If RangeOverlaps(CurrentStart, CurrentEnd, Len1 + 1, Len1 + Len2) Then
                            OverlapAllowed = False
                        End If
                    End If
                End If
                
                If OverlapAllowed Then
                    If InStr(1, ForbiddenRegions, "3") <> 0 Then
                        If RangeOverlaps(CurrentStart, CurrentEnd, Len1 + Len2 + 1, Len1 + Len2 + Len3) Then
                            OverlapAllowed = False
                        End If
                    End If
                End If
            
            
            If OverlapAllowed Then
                tOverlap = SubSequenceSelect(FinalSequence, CurrentStart, CurrentEnd)
                ' if the Tm is allowed
                tempTm = OligoTm(tOverlap)
                If tempTm >= TargetOverlapTm Then
                    ' and it doesn't have multiple annealing sites
                    If StringCharCount_IncludeOverlap(FinalSequence, tOverlap) > 1 Then
                        'Debug.Print ("Multiple annealing sites for " & tOverlap)
                    Else
                        Set tColl = New VBA.Collection
                        With tColl
                            .Add tOverlap
                            .Add CurrentStart
                            .Add tempTm
                        End With
                        PutativeOverlaps.Add tColl
                    End If
                End If
            End If
            
        Next CurrentStart
        
    Next CurrentLength
    
    Set ConstructOverlaps = PutativeOverlaps
    
    Set tColl = Nothing
    Set PutativeOverlaps = Nothing

End Function


'****************************************************************************************************
Private Function OptimizeOverlaps() As VBA.Collection
    
'====================================================================================================
'Calculates the optimal Gibson overlap for annealing sequences 1 and 3, with inserting sequence 2 in between
'Juraj Ahel, 2016-12-20, for Gibson assembly

'output:
'Collection of
    'OVERLAP as String      'overlap sequence
    'DG as Double           'Gibbs energy for SS formation of overlap
    'TM as Double           'melting temperature of overlap
    '1 as String            'first fragment to be annealed
    '2 as String            'second fragment to be annealed

'====================================================================================================
'2017-01-19 adapt for class module


    Dim i As Long, j As Long, k As Long

    Dim SequencesFound As Long
    Dim tOverlap As String
    
    Dim PutativeOverlaps As VBA.Collection
    
    Dim TempInput As String
    Dim TempOutput As String
    Dim TempFilenameBase As String
    
    Dim OutputLines() As String
    Dim tempdGString As String
    Dim tempdG As Double
    Dim RegEx As New RegExp
    Dim SortedCollection As VBA.Collection
    Dim tColl As VBA.Collection
    
    Dim Buffer As String
    
        
    ' Construct collection of overlaps
    Set PutativeOverlaps = ConstructOverlaps()
        
    TempFilenameBase = TempFilenamesRoot & TempTimeStampName
    'TempFilenameBase = FileSystem_GetTempFolder(True) & "JA_Gibson_test"
    
    
    i = 0
    TempInput = TempFilenameBase & "_in"
    
    Do While FileSystem_FileExists(TempInput)
        i = i + 1
        If i > 1 Then
            TempInput = Left(TempInput, Len(TempInput) - 5)
        End If
        TempInput = TempInput & "_" & Format(i, "0000")
        
    Loop
    
    i = 0
    TempOutput = TempFilenameBase & "_out"
    
    Do While FileSystem_FileExists(TempOutput & ".fold")
        i = i + 1
        If i > 1 Then
            TempOutput = Left(TempOutput, Len(TempInput) - 5)
        End If
        TempOutput = TempOutput & "_" & Format(i, "0000")
        
    Loop
        
    Call CreateEmptyFile(TempInput)
    Call CreateEmptyFile(TempOutput & ".fold")
    
    'Debug.Print (TempOutput)
    
    'calculate secondary structure energies (ViennaRNA/RNAfold)
        
        Buffer = ""
        
        For i = 1 To PutativeOverlaps.Count
            'Debug.Print (PutativeOverlaps.Item(i).Item(1))
            Buffer = Buffer & PutativeOverlaps.Item(i).Item(1) & vbCrLf
        Next i
        Call WriteTextFile(Buffer, TempInput, True)
        
        'make sure the output file is empty (otherwise stuff gets appended to it, making all kinds of weird hidden bugs and crap)
        Debug.Assert (Len(ReadTextFile(TempOutput & ".fold")) = 0)
        
        Call CallProgram( _
            ProgramCommand:=RNAFoldCommand, _
            ProgramPath:=RNAFoldPath, _
            ArgList:="--noGU --noPS -T 10 --infile=" & TempInput & " --outfile=" & TempOutput, _
            WaitUntilFinished:=True, _
            WindowMode:=0 _
            )
         
        TempOutput = TempOutput & ".fold" 'ViennaRNA 2.3 appends the extension automatically'
    
    'extract results
        With RegEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = ".*\( *(-?[0-9]+\.[0-9]+)\)"
        End With
                
        OutputLines = Split(ReadTextFile(TempOutput), vbCrLf)
        Set SortedCollection = New VBA.Collection
        
    'sort results into a collection
        For i = 1 To PutativeOverlaps.Count
            tempdG = val(RegEx.Replace(OutputLines(2 * i - 1), "$1"))
            Set tColl = PutativeOverlaps.Item(i)
            tColl.Add tempdG
                           
            If i = 1 Then
                SortedCollection.Add tColl
            Else
                j = 1
                Do While j <= SortedCollection.Count
                    If tempdG < SortedCollection.Item(j).Item(4) Then
                        j = j + 1
                    Else
                        Exit Do
                    End If
                Loop
                
                If j > SortedCollection.Count Then
                    SortedCollection.Add tColl
                Else
                    SortedCollection.Add tColl, Before:=j
                End If
                
            End If
            
        Next i
        
    Set OptimizeOverlaps = ExtractBestOverlaps(SortedCollection)
    
    For i = 1 To OptimizeOverlaps.Count
        'Debug.Print (OptimizeOverlaps.Item(i))
    Next i
    
    Set SortedCollection = Nothing
    Set PutativeOverlaps = Nothing
    Set tColl = Nothing
    

End Function


Private Function CalculateNicenessScore(ByRef GibsonOverlap As VBA.Collection) As Double
    
    Const conOverlapParameters As Long = 4
    
    Dim Seq As String
    Dim Index As Long
    Dim dG As Double
    Dim S As Double
    Dim GC As Double
    Dim Tm As Double
    
    Dim MiddleIndex As Long
        
    Dim deltaL As Long
    Dim deltaTm As Double
    Dim deltaIndex As Long
    Dim GCExcess As Double
        
    Dim OutBuff As String
        
    'Each GibsonOverlap is a collection [Sequence][TargetIndex]
    If Not GibsonOverlap Is Nothing Then
        With GibsonOverlap
            If .Count = conOverlapParameters Then
                If VarType(.Item(1)) = vbString And VarType(.Item(2)) = vbLong And VarType(.Item(3) = vbDouble) And VarType(.Item(4) = vbDouble) Then
                    Seq = .Item(1)
                    Index = .Item(2) + Len(Seq) / 2
                    Tm = .Item(3)
                    dG = .Item(4)
                End If
            End If
        End With
    End If
            
    MiddleIndex = Len(FirstSequence) + Len(InsertedSequence) \ 2
    
    'terminating in G/C is good
    If Right(Seq, 1) Like "[GC]" Then
        S = S + 5
    End If
    If Left(Seq, 1) Like "[GC]" Then
        S = S + 5
    End If
    
    'high GC is bad, low less bad
    GC = DNAGCContent(Seq)
    GCExcess = GC - 0.5
    Select Case GCExcess
        Case Is > 0.4
            S = S - 500 * GCExcess
        Case Is > 0.3
            S = S - 200 * GCExcess
        Case Is > 0.2
            S = S - 150 * GCExcess
        Case Is > 0.1
            S = S - 100 * GCExcess
        Case Is < -0.4
            S = S + 200 * GCExcess
        Case Is < -0.2
            S = S + 100 * GCExcess
        Case Is < -0.1
            S = S + 25 * GCExcess
    End Select
        
        
    'the shorter the better
    deltaL = (Len(Seq) - TargetOverlapLength)
    Select Case deltaL
        Case Is > TargetOverlapLength
            S = S - 5 * deltaL
        Case Is > 10
            S = S - 2 * deltaL
        Case Is > 5
            S = S - 1 * deltaL
        Case Else
            S = S - 0.5 * deltaL
    End Select
        
    'Tm should not be too high (too low is guaranteed from before, but double check)
    deltaTm = Tm - TargetOverlapTm
    Select Case deltaTm
        Case Is < 0
            S = S - 10 * (TargetOverlapTm - Tm)
        Case Is > 10
            S = S - 10 * (Tm - TargetOverlapTm)
        Case Is > 5
            S = S - 5 * (Tm - TargetOverlapTm)
        Case Else
            S = S - 1 * (Tm - TargetOverlapTm)
    End Select
    
    'dg should be as little as possible
    S = S - 5 * dG ^ 2
        
    'should be as close to middle as possible
    deltaIndex = Abs(Index - MiddleIndex)
    Select Case deltaIndex
        Case Is > 50
            S = S - 1 * deltaIndex
        Case Is > 25
            S = S - 0.75 * deltaIndex
        Case Is > 10
            S = S - 0.5 * deltaIndex
        Case Is > 5
            S = S - 0.25 * deltaIndex
        Case Else
    End Select
    
    
    'Debug.Print (Round(S, 0) & " " & Seq & " di:" & deltaIndex & " dG:" & Round(dG, 1) & " dTm:" & Round(deltaTm, 0) & " GCE:" & Round(GCExcess, 2))
    
    OutBuff = OutBuff & Round(S, 0)
    OutBuff = OutBuff & vbTab & Seq
    OutBuff = OutBuff & vbTab & "di:" & deltaIndex
    OutBuff = OutBuff & vbTab & "dG:" & Round(dG, 1)
    OutBuff = OutBuff & vbTab & "dTm:" & Round(deltaTm, 0)
    OutBuff = OutBuff & vbTab & "GCE:" & Round(GCExcess, 2)
    
    'Call WriteTextFile(OutBuff, "C:\temp\debuginfo.log", Append:=True)
    
    CalculateNicenessScore = S
    
End Function

'****************************************************************************************************
Private Function ExtractBestOverlaps(ByRef SortedCollection As VBA.Collection) As VBA.Collection
'====================================================================================================
'Extract the best overlap from the sorted collection, based on length and dG
'Juraj Ahel, 2016-12-20, for Gibson assembly
'====================================================================================================
'2017-01-19 adapt for class module
    
    
    Dim OverlapLength As Long
    Dim MinLength As Long
    Dim MindG As Double
    
    Dim BestScore As Double
    
    Dim i As Long
    Dim ChosenIndex As Long
    
    Dim tColl As VBA.Collection
    Dim SortColl2 As VBA.Collection
    
    Set tColl = SortedCollection.Item(1)
    Set SortColl2 = tColl
    
    BestScore = CalculateNicenessScore(tColl)
    ChosenIndex = 1
    
    'grab the shortest overlap with minimal dG
    For i = 2 To SortedCollection.Count
        Set tColl = SortedCollection.Item(i)
        If CalculateNicenessScore(tColl) >= BestScore Then
            Set SortColl2 = tColl
            ChosenIndex = i
        End If
    Next i
    
    OverlapLength = Len(SortColl2.Item(1))
    
    Set tColl = New VBA.Collection
    
    Debug.Print ("Seq" & ": " & SortColl2.Item(1) & " " & SortColl2.Item(4) & " " & BestScore)
    
    With tColl
        .Add SortColl2.Item(1), "OVERLAP"
        .Add Len(SortColl2.Item(1)), "LEN"
        .Add SortColl2.Item(4), "DG"
        .Add SortColl2.Item(3), "TM"
    End With
    
    Set ExtractBestOverlaps = tColl
    
    Set tColl = Nothing
    Set SortColl2 = Nothing
                
End Function

