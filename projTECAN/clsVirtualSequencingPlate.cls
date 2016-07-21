VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVirtualSequencingPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2015-11-xx
'Last update 2016-03-07
'====================================================================================================

Option Explicit

Private Const PlateXDim As Integer = 12
Private Const PlateYDim As Integer = 8

Public Enum enPlateStates
    eRestricted
    eTaken
    eFree
    eTemplatePipetted
    ePrimerPipetted
    eAllPipetted
End Enum

Private Enum pDirections
    eTopLeft
    eTopRight
    eBottomLeft
    eBottomRight
End Enum

'Plate definition
Private pXDim As Integer
Private pYDim As Integer
Private pTotalNumber As Integer
Private pPlateNumber As Integer

'Plate temporary state
Private pStateLinear() As enPlateStates

'Summary vars
Private TotalSeqReactionNumber As Integer

'Data collections
Private pSeqReactionList() As clsSeqReaction
Private pSequencingSamples As clsSeqSamples
Private pPrimers As clsPrimers

Private Sub Class_Initialize()

    'If pSeqReactionList() Is Nothing Then
    '    Set pPrimers = New Collection
    '    Set pPrimersKeys = New Collection
    'End If

End Sub

Private Property Let NumberOfPlates(N As Integer)

pPlateNumber = N

End Property

Public Property Get NumberOfPlates() As Integer

NumberOfPlates = pPlateNumber

End Property

Public Property Get RowsCount() As Integer
    RowsCount = pXDim
End Property

Public Property Get ColumnsCount() As Integer
    ColumnsCount = pYDim
End Property

Public Property Get Well(x As Integer, y As Integer) As clsSeqReaction
    Set Well = pSeqReactionList(XYToLinear(x, y))
End Property

Public Property Get WellState(x As Integer, y As Integer) As enPlateStates
    WellState = pStateLinear(XYToLinear(x, y))
End Property

Public Property Let WellState(x As Integer, y As Integer, State As enPlateStates)
    pStateLinear(XYToLinear(x, y)) = State
End Property


Public Sub ResetPipettingState()

    Dim i As Integer
    
    For i = 1 To TotalSeqReactionNumber
        Select Case pStateLinear(i)
            Case eTaken, eRestricted, eFree
                'Do nothing
            Case ePrimerPipetted, eTemplatePipetted, eAllPipetted
                pStateLinear(i) = eTaken
        End Select
    Next i

End Sub

Private Sub Define(x As Integer, y As Integer)
    
    Dim i As Integer
    
    pXDim = x
    pYDim = y
    pTotalNumber = pXDim * pYDim
    ReDim pState(1 To pXDim, 1 To pYDim)
    ReDim pStateLinear(1 To pTotalNumber)
    ReDim pSeqReactionList(1 To pTotalNumber)
    
    For i = 1 To pTotalNumber
        pStateLinear(i) = eFree
        Set pSeqReactionList(i) = New clsSeqReaction
    Next i
    
End Sub

Private Function XYToLinear(x As Integer, y As Integer) As Integer
    XYToLinear = (x - 1) * pYDim + y
End Function

Private Sub LinearToXY(pos As Integer, x As Integer, y As Integer)
    x = 1 + (pos - 1) \ pYDim
    y = 1 + (pos - 1) Mod pYDim
End Sub

Public Function IsWellUsed(XPos As Integer, YPos As Integer) As Boolean

    IsWellUsed = (pStateLinear(XYToLinear(XPos, YPos)) = eTaken)

End Function

Private Function IsFree(pos As Integer) As Boolean
    If pStateLinear(pos) = eFree Then IsFree = True Else IsFree = False
End Function

Private Function FindFree(Optional Direction As pDirections = eTopLeft) As Integer
'Finds the first free well starting from either corner of the plate
'Giving priority to columns over rows

    Dim Xstart As Integer, Ystart As Integer
    Dim Xstep As Integer, Ystep As Integer
    Dim i As Integer, j As Integer
    Dim TempPosition As Integer
    
    'still need to handle this, so far it handles only proper enumerated directions
    If VarType(Direction) = vbString Then Direction = UCase(Direction)
    
    Select Case Direction
        Case eTopLeft
            Xstart = 1
            Ystart = 1
            Xstep = 1
            Ystep = 1
         Case eTopRight
            Xstart = pXDim
            Ystart = 1
            Xstep = -1
            Ystep = 1
        Case eBottomLeft
            Xstart = 1
            Ystart = pYDim
            Xstep = 1
            Ystep = -1
        Case eBottomRight
            Xstart = pXDim
            Ystart = pYDim
            Xstep = -1
            Ystep = -1
    End Select
        
    For i = Xstart To pXDim - Xstart + 1 Step Xstep
        For j = Ystart To pYDim - Ystart + 1 Step Ystep
            TempPosition = XYToLinear(i, j)
            If IsFree(TempPosition) Then
                FindFree = TempPosition
                GoTo 999
            End If
        Next j
    Next i
    
    'if didn't find
    ErrorReport (1)
        
999 End Function

Private Function ImportRange(InputRange As Range) As Variant
    
    Dim InputArray() As Variant
    
    
    TotalSeqReactionNumber = InputRange.Rows.Count
    ReDim InputArray(1 To TotalSeqReactionNumber, 1 To 2)
    InputArray = InputRange.Value
    
    ImportRange = InputArray

End Function

Public Sub ImportSequencingList()
    
    Const IAmDebugging = True
    
    'input variables (temp?)
    Dim InputRange As Range
    Dim InputArray() As Variant
        
    'data collection
    Dim SequencingSamples As clsSeqSamples
    Dim Primers As clsPrimers
    Dim PrimerCounts As clsTypeCollection
    
    Set SequencingSamples = New clsSeqSamples
    Set Primers = New clsPrimers
    Set PrimerCounts = New clsTypeCollection
        
    'temp exchange vars and iterators
    Dim ctempSeqSample As clsSeqSample
    Dim ctempPrimer As clsPrimer
    Dim ctempDNA As clsDNA
    Dim tempDNAText As String, tempPrimerText As String
    Dim tempCollectionMember As Variant
    Dim i As Integer, j As Integer
        
    'Output variables
    Dim sortedPrimers As New clsPrimers
    Dim sortedPrimerCounts As New clsTypeCollection
    
    'control variables
    Dim ElementNotIn As Boolean
    
    If IAmDebugging Then
        Set InputRange = Range("TESTINPUTR")
    Else
        Set InputRange = Selection
    End If
    
    InputArray = ImportRange(InputRange)
        
    'Put all the primers, DNAs, and Seq Reactions into appropriate collections
    For i = 1 To TotalSeqReactionNumber
        
        tempDNAText = InputArray(i, 1)
        tempPrimerText = InputArray(i, 2)
        
        
        'if this sample wasn't added yet, add it
        ElementNotIn = Not (IsElementOf(tempDNAText, SequencingSamples.SeqSamples))
        If ElementNotIn Then
            
            'adding the template part
            Set ctempDNA = New clsDNA
            Set ctempSeqSample = New clsSeqSample
            ctempDNA.Name = tempDNAText
            Set ctempSeqSample.Template = ctempDNA
            SequencingSamples.Add ctempSeqSample
        
        Else
        
            Set ctempSeqSample = SequencingSamples.SeqSample(tempDNAText)
            
        End If
        
            
        ElementNotIn = Not (IsElementOf(tempPrimerText, Primers.Primers))
        If ElementNotIn Then
            
            'adding the primer to the list of primer used in this experiment
            Set ctempPrimer = New clsPrimer
            ctempPrimer.Name = tempPrimerText
            Primers.AddPrimer ctempPrimer
            PrimerCounts.Add 1, tempPrimerText
                    
        Else
            
            'Increase count of an element inside the collection
            Set ctempPrimer = Primers.Primer(tempPrimerText)
            tempCollectionMember = PrimerCounts.Item(tempPrimerText)
            PrimerCounts.Remove tempPrimerText
            PrimerCounts.Add tempCollectionMember + 1, tempPrimerText
        
        End If
        
        
        'add the annotations -> primer inside the seq sample, seq sample inside the primer
        ctempPrimer.SeqSamples.Add ctempSeqSample
        ctempSeqSample.Primers.AddPrimer ctempPrimer
        
        '#OLD
        'ElementNotIn = Not IsElementOf(tempPrimerText, SequencingSamples.SeqSample(tempDNAText).Primers.Primers)
        'If ElementNotIn Then
        
        '    Set ctempPrimer = New clsPrimer
        '    ctempPrimer.Name = tempPrimerText
        '    SequencingSamples.SeqSample(tempDNAText).AddPrimer ctempPrimer
            'Maybe also add the reverse annotation... Template in Primer!
            
        'End If
        
    
        'If (Not IsElementOf(tempDNAText, Primers.Primer(tempPrimerText).SeqSamples)) Then
        '
        '    Set ctempSeqSample = New clsSeqSample
        '    ctempPrimer.Name = tempPrimerText
        '    SequencingSamples.SeqSample(tempDNAText).AddPrimer ctempPrimer
        '    '###Maybe also add the reverse annotation... Template in Primer!
       
'        End If
        
    Next i
    
    '#OLD
    'This is not necessary - it could have been done directly in the initial interation
    'since objects can be passed by reference, and their values don't get "copied".
    'That is, if a clsDNA object is set as a template for the clsSeqSample object, and
    'later a name is added to this clsDNA object, it will automatically be updated "inside"
    'the clsSeqSample object, as it's not really inside, it's just linked, like with a pointer!
    'This also means there is no significant overhead from having "rich" objects that include links
    'to other objects, as it is just the pointer!
    'So, at some point this should be rewritten
    'For i = 1 To SequencingSamples.Count
    '    Set ctempSeqSample = SequencingSamples.SeqSamples(i)
    '    For j = 1 To ctempSeqSample.Primers.Count
    '        Set ctempPrimer = ctempSeqSample.Primers.Primers(j)
    '        ctempPrimer.SeqSamples.Add ctempSeqSample
    '        Primers.Primer(ctempPrimer.Name).SeqSamples.Add ctempSeqSample
    '    Next j
    'Next i
    
    'Sort primers by counts!
    Set sortedPrimers = New clsPrimers
    Set sortedPrimerCounts = New clsTypeCollection
    
    Dim tempCount As Integer
    Dim tempKey As String
    
    For i = 1 To Primers.Primers.Count
        
        tempCollectionMember = CInt(CollectionMaxElement(PrimerCounts))
        
        tempCount = PrimerCounts.Item(tempCollectionMember)
        tempKey = PrimerCounts.Keys(tempCollectionMember)
        
        sortedPrimerCounts.Add tempCount, tempKey
        sortedPrimers.AddPrimer Primers.Primer(tempKey)
        
        PrimerCounts.Remove (tempCollectionMember)
        
    Next i
    
    Set PrimerCounts = Nothing
    Set Primers = Nothing
    Set ctempSeqSample = Nothing
    Set ctempPrimer = Nothing
    
    NumberOfPlates = TotalSeqReactionNumber \ PlateXDim * PlateYDim + (1 + ((TotalSeqReactionNumber Mod PlateXDim * PlateYDim) = 0))
    Define PlateXDim * NumberOfPlates, PlateYDim '###add support for different plates
    
    
    If pXDim * pYDim - TotalSeqReactionNumber > 0 Then BlockExcessTubes pXDim * pYDim - TotalSeqReactionNumber
    
    FillWithPrimerList sortedPrimers, sortedPrimerCounts
    
    Set SequencingSamples = Nothing
    Set sortedPrimers = Nothing
    Set sortedPrimerCounts = Nothing

End Sub

Private Sub BlockExcessTubes(TubesToBlock As Integer)
'labels the excess tubes as "unused", for easier handling later
'to ensure that the empty tubes are in the end of the plate

Dim pos As Integer
pos = XYToLinear(pXDim, pYDim)

Do While TubesToBlock > 0
    pStateLinear(pos) = eRestricted
    mUp pos
    TubesToBlock = TubesToBlock - 1
Loop

End Sub



Private Function CollectionMaxElement(InputCollection As clsTypeCollection) As Integer
    'Finds an element of the collection (of integers) with the highest value
    'if multiple are equal, the element that occurs earliest is taken
    
    Dim i As Integer
    Dim temp As Integer
    temp = InputCollection.Item(1)
    CollectionMaxElement = 1
    
    For i = 1 To InputCollection.Count
        If InputCollection.Item(i) > temp Then
            temp = InputCollection.Item(i)
            CollectionMaxElement = i
        End If
    Next i

End Function

Private Sub FillWithPrimerList(Primers As clsPrimers, Counts As clsTypeCollection)

    Dim i As Integer

    For i = 1 To Primers.Primers.Count
        FillWithPrimer Primers.Primer(i), Counts.Item(i)
    Next i
    
    Set Primers = Nothing
    Set Counts = Nothing
    
    'tempOutput
    
End Sub

Public Sub tempOutput()
    
    'Debug.Print outrange.Name
    
    Dim i As Integer, j As Integer
    Dim testout As Worksheet
    Set testout = Workbooks(1).Worksheets("test")
    testout.Cells.Clear

    Dim outrange As Range
    Dim outarray()
    Set outrange = Range("testoutput").Offset(1, 10).Resize(pYDim, pXDim)
    ReDim outarray(1 To pYDim, 1 To pXDim)
    
    For i = 1 To pYDim
        For j = 1 To pXDim
            outarray(i, j) = pSeqReactionList(XYToLinear(j, i)).Primer.Name
        Next j
    Next i
    
    outrange.Value = outarray
    
    Set outrange = outrange.Offset(10, 0)
    
    For i = 1 To pYDim
        For j = 1 To pXDim
            outarray(i, j) = pSeqReactionList(XYToLinear(j, i)).Template.Name
        Next j
    Next i
    
    outrange.Value = outarray

End Sub


Private Sub FillWithPrimer(Primer As clsPrimer, Count As Integer)
'Adds a virtual primer to the virtual plate in a particular order
'I thought to be the best (as many full lines as possible)
'It also adds the appropriate DNA samples to the virtual plate
'to positions that correspond to primers, so far sorted in a way
'that's not really controlled (2016-01-07)
'This is iterated for each primer to be added by the parent sub
    
    
    Dim startpos As Integer, startX As Integer, startY As Integer
    Dim tempCount As Integer
    Dim EndControl As Boolean
    Dim DebugText As String
    Dim ctempTemplate As clsDNA
    
    tempCount = 0
    startpos = 0
    startX = 0
    startY = 0
    EndControl = False
    
    Do While tempCount < Count
        
        startpos = FindFree(eTopLeft)
       
        EndControl = False
        
        Do Until EndControl
            
            tempCount = tempCount + 1
            
            LinearToXY startpos, startX, startY
            Set pSeqReactionList(startpos).Primer = Primer
            Set ctempTemplate = Primer.SeqSamples.SeqSamples(tempCount).Template
            Set pSeqReactionList(startpos).Template = ctempTemplate
                
            pStateLinear(startpos) = eTaken
            
            
            EndControl = (startX = pXDim) Or tempCount = Count
            If Not EndControl Then mRight startpos
            EndControl = EndControl Or pStateLinear(startpos) <> eFree
            
        Loop
        
    Loop
    
Set ctempTemplate = Nothing

End Sub

'****************************************************************************************************
'====================================================================================================
'Virtualized movements across the plate!
'Juraj Ahel, 2015-11-23
'Last update 2015-11-23
'====================================================================================================
Private Sub PlateWellMovement(Direction As String, pos As Integer)
'maybe put all the movement commands here, and make mLeft, etc be just aliases that call this sub
'
'    Dim x As Integer, y As Integer
'    Dim xter As Integer, yter As Integer
'    Dim mov As Integer
'    Dim xdir As Boolean
'
'    LinearToXY pos, x, y
'
'    If x <> xter Then
'        x = x + xmov
'    Else
'        If y <> yter Then
'        y = y + xmov
'
'    Select Case Direction
'
'        Case "Left"
'            If x <> 1 Then x = x - 1 else if y<>1 Then y=
'
'            If x = 1 Then
'                If y = 1 Then ErrorReport (2) Else y = y - 1
'                x = pXDim
'            Else
'                x = x - 1
'            End If
'
'        Case "Right"
'            If x = pXDim Then
'                If y = pYDim Then ErrorReport (2) Else y = y + 1
'                x = 1
'            Else
'                x = x + 1
'            End If
'
'        Case "Up"
'            If y = 1 Then
'                If x = 1 Then ErrorReport (2) Else x = x - 1
'                y = pYDim
'            Else
'                y = y - 1
'            End If
'
'        Case "Down"
'            If y = pYDim Then
'                If x = pXDim Then ErrorReport (2) Else x = x + 1
'                y = 1
'            Else
'                y = y + 1
'            End If
'        Case Else
'            ErrorReport (3)
'    End Select
'
'    pos = XYToLinear(x, y)

End Sub

Private Sub mLeft(pos As Integer)
    Dim x As Integer, y As Integer
    LinearToXY pos, x, y
    If x = 1 Then
        If y = 1 Then ErrorReport (2) Else y = y - 1
        x = pXDim
    Else
        x = x - 1
    End If
    pos = XYToLinear(x, y)
End Sub

Private Sub mRight(pos As Integer)
    Dim x As Integer, y As Integer
    LinearToXY pos, x, y
    If x = pXDim Then
        If y = pYDim Then ErrorReport (2) Else y = y + 1
        x = 1
    Else
        x = x + 1
    End If
    pos = XYToLinear(x, y)
End Sub

Private Sub mUp(pos As Integer)
    Dim x As Integer, y As Integer
    LinearToXY pos, x, y
    If y = 1 Then
        If x = 1 Then ErrorReport (2) Else x = x - 1
        y = pYDim
    Else
        y = y - 1
    End If
    pos = XYToLinear(x, y)
End Sub

Private Sub mDown(pos As Integer)
    Dim x As Integer, y As Integer
    LinearToXY pos, x, y
    If y = pYDim Then
        If x = pXDim Then ErrorReport (2) Else x = x + 1
        y = 1
    Else
        y = y + 1
    End If
    pos = XYToLinear(x, y)
End Sub

'****************************************************************************************************
Public Function ExtractPlate(Number As Integer) As clsWorktableSeqPlate
'====================================================================================================
'Create an actual plate from the virtual plate, to be used in automatic pipetting with TECAN EVO
'Juraj Ahel, 2016-03-07
'Last update 2016-03-07
'====================================================================================================

    Dim i As Integer
    Dim TotalNumberPerPlate As Integer
    Dim Offset As Integer
    Dim tempcWorktableSeqPlate As clsWorktableSeqPlate
                
    If Number < 1 Then
        ErrorReport 4
    ElseIf Number > TotalSeqReactionNumber Then
        ErrorReport 5
    End If
                
    Set tempcWorktableSeqPlate = New clsWorktableSeqPlate
    
    tempcWorktableSeqPlate.Define
    
    TotalNumberPerPlate = PlateXDim * PlateYDim
    
    Offset = (Number - 1) * TotalNumberPerPlate
    
    Do
        i = i + 1
        Set tempcWorktableSeqPlate.Well(i) = pSeqReactionList(Offset + i)
        tempcWorktableSeqPlate.State(i) = eTaken
    Loop Until (i = TotalNumberPerPlate) Or ((Offset + i) >= TotalSeqReactionNumber)
    
    
    Set ExtractPlate = tempcWorktableSeqPlate
    Set tempcWorktableSeqPlate = Nothing

End Function


'****************************************************************************************************
'====================================================================================================
'Catching all the errors.
'Juraj Ahel, 2015-11-23
'Last update 2016-03-07
'====================================================================================================

Private Sub ErrorReport(ErrorCode As Integer)

    Dim ErrorString As String

    Select Case ErrorCode
        Case 1
            ErrorString = "No free tubes found in plate."
        Case 2
            ErrorString = "Tried to move outside of plate."
        Case 3
            ErrorString = "Tried to do a well movement in nonexisting direction."
        Case 4
            ErrorString = "Plate number cannot be less than 1."
        Case 5
            ErrorString = "Tried to select plate index larger than total number of sequencing plates."
    End Select
    
    R ErrorString
    ErrorReportGlobal ErrorCode, ErrorString
    
    
End Sub

Private Sub Class_Terminate()

    Set pSequencingSamples = Nothing
    Set pPrimers = Nothing

    Dim i As Integer
    For i = 1 To pTotalNumber
        Set pSeqReactionList(i) = Nothing
    Next i

End Sub

