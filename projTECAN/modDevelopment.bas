Attribute VB_Name = "modDevelopment"
Option Explicit

Sub testoooo()

Dim aa As clsLIHA

Set aa = New clsLIHA

For i = 1 To 8
    aa.Pipette(i).Volume = 5
Next i

c = aa.Pipette(3).Count

End Sub
'****************************************************************************************************
Sub DoTheThingWrapper()

'====================================================================================================
'Sets up the automatic dispensing from a virtual plate
'And generates the actual commands
'Juraj Ahel, 2016-01-19
'Last Update 2016-03-15
'====================================================================================================

    Dim Worktable As clsWorktableSetup
    Dim tempcContainer As clsWorktableContainer
    Dim tempsPlate As clsVirtualSequencingPlate
    
    Set Worktable = New clsWorktableSetup
    
    Set tempcContainer = New clsWorktableContainer
    tempcContainer.Define PlateType:="96", Grid:=12, Site:=0
    tempcContainer.ImportRange Range("Container1")
    
    Worktable.Containers.Add tempcContainer
    
    Set tempsPlate = New clsVirtualSequencingPlate
    tempsPlate.ImportSequencingList
    Worktable.DefineSequencingPlatesFromVirtual tempsPlate
    
    DoTheThing Worktable, tempsPlate
    
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
                If .IsWellUsed(i, j) Then
                    SamplesArray(i, j) = .Well(i, j).Template.Name
                    PrimersArray(i, j) = .Well(i, j).Primer.Name
                Else
                    SamplesArray(i, j) = conIgnorEl
                    PrimersArray(i, j) = conIgnorEl
                End If
            End With
        Next j
    Next i


End Sub
'****************************************************************************************************
Sub DoTheThingDecideOnNextSamples( _
    ByRef Worktable As clsWorktableSetup, _
    ByRef LIHA As clsLIHA, _
    ByRef VirtualPlate As Variant)
    
'====================================================================================================
'
'Juraj Ahel, 2016-01-16
'Last Update 2016-03-16
'====================================================================================================
    
    Const IAmDebugging = True
    
    'constants
    Const conIgnorEl As String = ""
    Const conPrimerVolume As Double = 4
    Const conSampleVolume As Double = 3
    Const conPlateXDim As Integer = 12
    Const conPlateYDim As Integer = 8
    
    Dim rcount As Integer, ccount As Integer
    
    Dim i As Integer
    
    Dim tempCount As Integer
    Dim tempTotalCount As Integer
    Dim tempIndex As Integer
    Dim tempMax As String
    
    Dim IgnoreList As Collection
    Dim PrimerVolume As Double
    Dim SampleVolume As Double
    
    PrimerVolume = conPrimerVolume
    SampleVolume = conSampleVolume
    
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
                    tempTotalCount = tempTotalCount + tempCount
                    LIHA.Pipette(tempIndex).Volume = tempCount * SampleVolume
                    'All of the samples in all lanes with this name will be pipetted, so remove it from further steps:
                    MatrixElementReplace tempMax, VirtualPlate, conIgnorEl
                    
                End If
            
            End If
        
        Next i
    
    Loop

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
    Set LIHA.Worktable = Worktable
    
    '[!!!!]Get the input data into an array
    DoTheThingImportVirtualPlateData VirtualSequencingPlate, SamplesPlate, PrimersPlate
    
    TotalNumber = (UBound(tempVirtualPlate, 1) - LBound(tempVirtualPlate, 1) + 1) _
                * (UBound(tempVirtualPlate, 2) - LBound(tempVirtualPlate, 2) + 1)
            
    tempVirtualPlate = SamplesPlate
    CommandType = "Samples"
            
    '[MAIN]
    Do
    
        DoTheThingDecideOnNextSamples Worktable:=Worktable, LIHA:=LIHA, VirtualPlate:=tempVirtualPlate
        
        '### TODO: also add comment between rounds / etc
        
        'get aspiration commands
        
        Commands.Append (LIHA.AspirateCommand(CommandType:="CommandType"))
        
        Commands.Append (LIHA.DispenseCommand(CommandType:="CommandType"))
        
        
    
        R Commands.Output
        
        'if it's empty
        If MatrixElementCount("", tempVirtualPlate, CountAllExceptIgnored:=True, Recursion:=1) = 0 Then
            Select Case UCase(CommandType)
                Case "SAMPLES", "S"
                    'if I was pipetting samples, go to primers
                    CommandType = "Primers"
                Case "PRIMERS", "P"
                    'if I was pipetting primers, I am now finished
                    AllIsPipetted = True
                Case Else
                    ErrorReportGlobal 5080, "Wrong CommandType detected, cannot continue.", "modDevelopment:DoTheThing"
            End Select
        End If
        
        '###TODO Issue wash command
        'Commands.Append (LIHA.WashCommand("SEQUENCING")
        
        'Purge the virtual pipette
        LIHA.Purge
        
    Loop Until AllIsPipetted
        
        
        
       
        
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


End Sub
'****************************************************************************************************
Sub testat()

Dim a As clsLIHA

b = VarType(a)

Set a = New clsLIHA


b = VarType(a)


End Sub

