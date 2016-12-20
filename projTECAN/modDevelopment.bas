Attribute VB_Name = "modDevelopment"
Option Explicit

Private TargetPrimerVolume As Double
Private TargetSampleVolume As Double

Private Const tempConstParameter As String = "OptimizePipetting"
'Private Const tempConstParameter As String = "PreserveOrder"


'****************************************************************************************************
Sub DoTheThingWrapper()
Attribute DoTheThingWrapper.VB_ProcData.VB_Invoke_Func = "W\n14"

'====================================================================================================
'Sets up the automatic dispensing from a virtual plate
'And generates the actual commands
'Juraj Ahel, 2016-01-19
'Last Update 2016-03-15
'====================================================================================================
'2016-08-01 import cleaner/waste
'2016-08-02 add functionality needed to pipette other stuff at different volumes

    Dim Worktable As clsWorktableSetup
    Dim tempcContainer As clsWorktableContainer
    Dim tempsPlate As clsVirtualSequencingPlate
        
    'define the robot's worktable
        Set Worktable = New clsWorktableSetup
        
    'define pipetting parameters
        TargetSampleVolume = 3
        TargetPrimerVolume = 200
    
    'import cleaning stuff
        Call Worktable.ImportComponent("CLEANERSHALLOW", Grid:=10, Site:=0)
        Call Worktable.ImportComponent("WASTE", Grid:=10, Site:=1)
        Call Worktable.ImportComponent("CLEANERDEEP", Grid:=10, Site:=2)
    
    'import the source containers (temp!!!)
        'samples!
        Set tempcContainer = New clsWorktableContainer
            tempcContainer.Define PlateType:="96", Grid:=12, Site:=0
            tempcContainer.ImportRange Range("Container1")
    
        Worktable.Containers.Add tempcContainer
        
        'Primers!
        Set tempcContainer = New clsWorktableContainer
            tempcContainer.Define PlateType:="96", Grid:=12, Site:=1
            tempcContainer.ImportRange Range("Container2")
    
        Worktable.Containers.Add tempcContainer
    
    'define the virtual sequencing plate / grid
        Set tempsPlate = New clsVirtualSequencingPlate
            tempsPlate.ImportSequencingList
            Worktable.DefineSequencingPlatesFromVirtual tempsPlate
    
    'carry out the automation
        DoTheThing Worktable, tempsPlate
        
    
        
    
    'clean up
        Set Worktable = Nothing
        Set tempcContainer = Nothing
        Set tempsPlate = Nothing

End Sub
'****************************************************************************************************
Sub DoTheThingImportVirtualPlateData( _
    ByRef VirtualSequencingPlate As clsVirtualSequencingPlate, _
    ByRef SamplesArray As Variant, _
    ByRef PrimersArray As Variant)

'====================================================================================================
'Imports the data from the class object to the arrays needed for pipetting
'Juraj Ahel, 2016-03-08
'Last update 2016-03-15
'====================================================================================================

    Const conIgnorEl = ""
    
    Dim i As Integer
    Dim j As Integer
    
    Dim RowsCount As Integer
    Dim ColumnsCount As Integer
    
    RowsCount = VirtualSequencingPlate.RowsCount
    ColumnsCount = VirtualSequencingPlate.ColumnsCount
    
    ReDim SamplesArray(1 To RowsCount, 1 To ColumnsCount)
    ReDim PrimersArray(1 To RowsCount, 1 To ColumnsCount)
    
    For i = 1 To RowsCount
        For j = 1 To ColumnsCount
            With VirtualSequencingPlate
                If .IsWellUsed(j, i) Then
                    SamplesArray(i, j) = .Well(j, i).Template.Name
                    PrimersArray(i, j) = .Well(j, i).Primer.Name
                Else
                    SamplesArray(i, j) = conIgnorEl
                    PrimersArray(i, j) = conIgnorEl
                End If
            End With
        Next j
    Next i


End Sub

'****************************************************************************************************
Sub DoTheThing( _
    Worktable As clsWorktableSetup, _
    VirtualSequencingPlate As clsVirtualSequencingPlate)
'====================================================================================================
'Sets up the automatic dispensing from a virtual plate
'And generates the actual commands
'Juraj Ahel, 2016-01-11
'Last Update 2016-03-16
'====================================================================================================

    Const IAmDebugging = True
    
    'constants
    Const conIgnorEl As String = ""
    
    Dim IgnoreList As Collection
    
    'objects
    Dim Commands As clsCommandSequence
    Dim LIHA As clsLIHA

    'temp and iterators
    Dim i As Integer, j As Integer
    
    'flow control
    Dim GotoNextStep As Boolean
    Dim AllIsPipetted As Boolean
    Dim CommandType As String
    
    'containers
    Dim SamplesPlate(), PrimersPlate() 'as Variants
    Dim tempVirtualPlate()
    
    'descriptors
    Dim TotalNumber As Integer
            
    '[INITIALIZATION]
    Set Commands = New clsCommandSequence 'here I will store the command set
    
    Set LIHA = New clsLIHA
        With LIHA
            Set .Worktable = Worktable
            .TargetSampleVolume = TargetSampleVolume
            .TargetPrimerVolume = TargetPrimerVolume
        End With
    
    Set IgnoreList = New Collection
    IgnoreList.Add conIgnorEl
    
    
    '[!!!!]Get the input data into an array
    DoTheThingImportVirtualPlateData VirtualSequencingPlate, SamplesPlate, PrimersPlate
    
    
    'start with pipetting samples
        tempVirtualPlate = SamplesPlate
        CommandType = "Samples"
        
        TotalNumber = (UBound(tempVirtualPlate, 1) - LBound(tempVirtualPlate, 1) + 1) _
                    * (UBound(tempVirtualPlate, 2) - LBound(tempVirtualPlate, 2) + 1)
            
    '[MAIN]
    Do
    
        DoTheThingDecideOnNextSamples Worktable:=Worktable, LIHA:=LIHA, VirtualPlate:=tempVirtualPlate, CommandType:=CommandType
        
        '### TODO: also add comment between rounds / etc
        
        'get aspiration commands
        
        Debug.Print ("AspirateCommands: ")
        
        Commands.Append (LIHA.AspirateCommand(CommandType))
        
        Debug.Print ("DispenseCommands: ")
        
        Commands.Append (LIHA.DispenseCommand(CommandType))
                
        'if the virtual plate is empty (all primers or samples were pipetted)
        If MatrixElementCount("", tempVirtualPlate, CountAllExceptIgnored:=True, IgnoreList:=IgnoreList, Recursion:=1) = 0 Then
            Select Case UCase(CommandType)
                Case "SAMPLES", "S"
                    'if I was pipetting samples, go to primers
                    CommandType = "Primers"
                    tempVirtualPlate = PrimersPlate
                Case "PRIMERS", "P"
                    'if I was pipetting primers, I am now finished
                    AllIsPipetted = True
                Case Else
                    ErrorReportGlobal 5080, "Wrong CommandType detected, cannot continue.", "modDevelopment:DoTheThing"
            End Select
        End If
        
        
        'check if pipettes are empty (it's an error if they are not, as it should only aspirate as much as it can dispense!!!
        
        If Not LIHA.Free Then
            Call ErrorReportGlobal(5081, "LIHA is not free after pipetting everything, aspirate - dispense volume mismatch!")
        End If
                
        Debug.Print ("WashCommands: ")
        Commands.Append (LIHA.WashCommand("SEQUENCING"))
        
        
    Loop Until AllIsPipetted
    
    R Commands.Output
       
    Debug.Print ("Setup successful!")
        
    'get the output table
        Call DoTheThingOutput( _
                Worktable:=Worktable, _
                VirtualSequencingPlate:=VirtualSequencingPlate, _
                Commands:=Commands)

        
        
    '2c if no, calculate the total number of occurences of the max element of each row and
    '   make the master row eat all of them
    '3  insert commands for aspiration of the target elements
    '4  define an aspirated LIHA
    '5  column by column, detect if a particular tip needs to dispense - fill a matrix with those
    '6  after this first pass
    
    ' IF LESS THAN 8 REQUIRED ROWS, WHAT THEN?
    
    
    
    
    
    '99  remove the elements that were covered so far from the table and repeat the procedure
    '   with the restricted table (either recursion [might be an overkill] or do.. while not empty
    
    
    
    '###TODO aspirate from all by counts
    
    'OutRan.Value = SamplesPlate
    
    '###OUTRAN is output and forwarded on to the next instance of the iterator for recursive solution

    '[DESTRUCTION]
        Set Commands = Nothing
        Set LIHA = Nothing
        Set IgnoreList = Nothing
        Set Worktable = Nothing
    
End Sub


'****************************************************************************************************
Sub DoTheThingDecideOnNextSamples( _
    ByRef Worktable As clsWorktableSetup, _
    ByRef LIHA As clsLIHA, _
    ByRef VirtualPlate As Variant, _
    ByVal CommandType As String)
    
'====================================================================================================
'
'Juraj Ahel, 2016-01-16
'Last Update 2016-03-16
'====================================================================================================
'2016-08-02 make it work properly with volumes for primers

    Const IAmDebugging = True
    
    'constants
    Const conIgnorEl As String = ""
    Const conPlateXDim As Integer = 12
    Const conPlateYDim As Integer = 8
    
    Dim rcount As Integer, ccount As Integer
    
    Dim i As Integer
    Dim j1 As Integer
    Dim j2 As Integer
    Dim k As Integer
    
    Dim tempCount As Integer
    Dim tempTotalCount As Integer
    Dim tempIndex As Integer
    Dim tempMax As String
    
    Dim TotalNumber As Long
    
    Dim IgnoreList As Collection
    Dim TargetVolume As Double
    Dim MaxAllowedSamplesAtOnce As Integer
    
    Select Case UCase(CommandType)
        Case "SAMPLES", "S"
            TargetVolume = TargetSampleVolume
        Case "PRIMERS", "P"
            TargetVolume = TargetPrimerVolume
        Case Else
            Call ErrorReportGlobal(1, "Unsupported pipetting command type: " & CommandType)
    End Select
    
    MaxAllowedSamplesAtOnce = LIHA.MaxAllowedAspirationVolume \ TargetVolume
    
    Set IgnoreList = New Collection
    IgnoreList.Add conIgnorEl
    
    tempCount = 0
    tempTotalCount = 0
    
    rcount = UBound(VirtualPlate, 1) - LBound(VirtualPlate, 1) + 1
    ccount = UBound(VirtualPlate, 2) - LBound(VirtualPlate, 2) + 1
    
        
    'Count the total number of nonpipetted reactions in the virtual plate
    TotalNumber = MatrixElementCount(Element:="", _
                                    Matrix:=VirtualPlate, _
                                    IgnoreList:=IgnoreList, _
                                    CountAllExceptIgnored:=True, _
                                    Recursion:=1)

    Do While Not (LIHA.Full) And Not (tempTotalCount = TotalNumber)
        
        For i = 1 To rcount
            
            tempIndex = LIHA.FirstUnusedPipette
            
            If tempIndex <> 0 Then 'if LIHA isn't full
                
                'find the most abundant element in current row
                tempMax = MatrixMaxElement(Matrix:=VirtualPlate, OnlyRowN:=i, IgnoreList:=IgnoreList)
                
                If Not IsMemberOf(tempMax, IgnoreList) Then 'if it's a valid name (e.g. not "")
                
                    LIHA.Pipette(tempIndex).Sample.Name = tempMax
                    tempCount = MatrixElementCount(tempMax, VirtualPlate)
                    
                    If tempCount <= MaxAllowedSamplesAtOnce Then
                        
                        'All of the samples in all lanes with this name will be pipetted, so remove it from further steps:
                        MatrixElementReplace tempMax, VirtualPlate, conIgnorEl
                        
                    Else
                    
                        'remove only part of the samples
                        tempCount = MaxAllowedSamplesAtOnce
                        j1 = 0
                        j2 = 0
                        k = 0
                        
                        For j2 = LBound(VirtualPlate, 2) To UBound(VirtualPlate, 2)
                            For j1 = LBound(VirtualPlate, 1) To UBound(VirtualPlate, 1)
                                If VirtualPlate(j1, j2) = tempMax Then
                                    VirtualPlate(j1, j2) = conIgnorEl
                                    k = k + 1
                                    If k >= tempCount Then Exit For
                                End If
                            Next j1
                            If k >= tempCount Then Exit For
                        Next j2
                        
                        If k <> tempCount Then
                            Call ErrorReportGlobal(5091, "A counting mismatch has occurred")
                        End If
                    
                    End If
                    
                    LIHA.Pipette(tempIndex).Volume = tempCount * TargetVolume
                    tempTotalCount = tempTotalCount + tempCount
                    
                End If
            
            End If
        
        Next i
    
    Loop

End Sub


'****************************************************************************************************
Sub DoTheThingOutput( _
    ByRef Worktable As clsWorktableSetup, _
    ByRef VirtualSequencingPlate As clsVirtualSequencingPlate, _
    ByRef Commands As clsCommandSequence)

    Dim tSeqPlate As clsWorktableSeqPlate
    
    Dim OutputReference As Excel.Range
    Dim tOutputRange As Excel.Range
    Dim TempOffset As Long
    Dim TempX As Long
    Dim TempY As Long
    
    Dim i As Long
    Dim j As Long
    
'====================================================================================================
'
'Juraj Ahel, 2016-08-01
'
'====================================================================================================
    
    Dim tempArray(1 To 96, 1 To 2) As String
    
    Set OutputReference = Excel.Range("Outputs")
    
    TempOffset = 1
    
    For i = 1 To Worktable.SequencingPlates.Count
        
        Set tSeqPlate = Worktable.SequencingPlates.Item(i)
            
        With tSeqPlate
            TempX = .maxColumn
            TempY = .maxRow
            Set tOutputRange = OutputReference.Offset(TempOffset, 4).Resize(TempY, TempX)
            tOutputRange.Value = .TemplatesArray
            tOutputRange.Borders.LineStyle = xlContinuous
            
            Set tOutputRange = tOutputRange.Offset(0, TempX + 2)
            tOutputRange.Value = .PrimersArray
            tOutputRange.Borders.LineStyle = xlContinuous
            
            For j = 1 To 96
                tempArray(j, 1) = j
                If Not .Well(j) Is Nothing Then
                    tempArray(j, 2) = .Well(j).Template.Name & "_" & .Well(j).Primer.Name
                Else
                    tempArray(j, 2) = ""
                End If
            Next j
            
            Set tOutputRange = OutputReference.Offset(97 * (i - 1), 0).Resize(96, 2)
            tOutputRange.Value = tempArray
            
        End With
            
        TempOffset = TempOffset + TempY + 2
            
    Next i
        
    Debug.Print (Commands.Output)
        
    
    Set tSeqPlate = Nothing
    Set OutputReference = Nothing
    Set tOutputRange = Nothing
    

End Sub

'****************************************************************************************************
Sub testat()

Dim a As clsLIHA

b = VarType(a)

Set a = New clsLIHA


b = VarType(a)


End Sub

