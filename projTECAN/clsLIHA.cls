VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLIHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const conDefaultLiquidClass As String = "Water"
Private Const conDefaultSpacing As Integer = 1

'Public Enum ePipettingStates
    
'programming constants
Private Const conQ As String = """"

'system constants
Private Const conDefaultPipetteNumber As Integer = 8 'setup of our TECAN EVO instrument (2016-03-09)
Private Const conMaxPipetteNumber As Integer = 12   'defined by TECAN EVO instrument
Private Const conMaximumSpacing As Integer = 4      'defined by TECAN EVO instrument

'system definition
Private pPipetteNumber As Integer

'data variables
Private pWorktable As clsWorktableSetup
Private pPipettes() As clsPipette
Private pCommands As clsCommandSequence
Private pVolumes() As Double

'state variables
Private pPipetteUsed() As Boolean


Private Sub Class_Initialize()

    Dim i As Integer
    
    PipetteNumber = conDefaultPipetteNumber
    
    ReDim pPipetteUsed(1 To PipetteNumber)
    ReDim pVolumes(1 To PipetteNumber)
    ReDim pPipettes(1 To PipetteNumber)
    
    For i = 1 To PipetteNumber
        pPipetteUsed(i) = False
        pVolumes(i) = 0
        Set pPipettes(i) = New clsPipette
    Next i

End Sub

Private Sub Class_Terminate()

    Dim i As Integer
    
    For i = 1 To PipetteNumber
        Set pPipettes(i) = Nothing
    Next i
    
    Set pWorktable = Nothing
    Set pCommands = Nothing

End Sub

Public Property Get PipetteNumber() As Integer
    PipetteNumber = pPipetteNumber
End Property

Private Property Let PipetteNumber(N As Integer)
    pPipetteNumber = N
End Property

Public Property Get Pipette(i) As clsPipette
    Set Pipette = pPipettes(i)
End Property

Private Property Get Commands() As clsCommandSequence
    Set Commands = pCommands
End Property

Private Property Set Commands(c As clsCommandSequence)
    Set pCommands = c
End Property

Private Sub AddCommandString(InputStr As String)
    
    If Commands Is Nothing Then Set Commands = New clsCommandSequence
    
    Commands.Append InputStr
    
End Sub

Private Sub AddPipettingStep(InputStep As clsPipettingCommand)
    AddCommandString InputStep.FullCommandText
End Sub

'****************************************************************************************************
Public Property Get FirstUnusedPipette(Optional AfterPipetteNo As Integer = 0) As Integer
'====================================================================================================
'
'Juraj Ahel, 2016-01-xx
'Last update 2016-01-09
'====================================================================================================
    
    Dim i As Byte
    Dim tempIndex As Byte
        
    Select Case AfterPipetteNo
        Case Is < 0
            AfterPipetteNo = 0
        Case Is > 256
            ErrorReportGlobal 13, "clsLIHA: AfterPipetteNo larger than Byte."
        Case Is > PipetteNumber
            ErrorReportGlobal 13, "clsLIHA: AfterPipetteNo is larger than total number of pipettes."
        Case Else
            'do nothing
    End Select
    
    tempIndex = 0
    i = AfterPipetteNo + 1
        
    Do While i <= PipetteNumber
        If Not (Pipette(i).Used) Then
            tempIndex = i
            GoTo 99
        End If
        i = i + 1
    Loop
    
99    FirstUnusedPipette = tempIndex
    
End Property

'****************************************************************************************************
Public Property Get FirstUsedPipette(Optional AfterPipetteNo As Integer = 0) As Integer
'====================================================================================================
'
'Juraj Ahel, 2016-01-xx
'Last update 2016-03-09
'====================================================================================================

    Dim i As Byte
    Dim tempIndex As Byte
    
    Select Case AfterPipetteNo
        Case Is < 0
            AfterPipetteNo = 0
        Case Is > 256
            ErrorReportGlobal 13, "clsLIHA: AfterPipetteNo larger than Byte."
        Case Is > PipetteNumber
            ErrorReportGlobal 13, "clsLIHA: AfterPipetteNo is larger than total number of pipettes."
        Case Else
            'do nothing
    End Select
    
    tempIndex = 0
    i = AfterPipetteNo + 1
        
    Do While i <= PipetteNumber
        If Pipette(i).Used Then
            tempIndex = i
            GoTo 99
        End If
        i = i + 1
    Loop
    
99    FirstUsedPipette = tempIndex
'0 means it wasn't found
    
End Property

Public Property Get Full() As Boolean

    If FirstUnusedPipette = 0 Then
        Full = True
    Else
        Full = False
    End If
    
End Property

Public Property Get Free() As Boolean

    Dim i As Integer
    Free = True
        
    For i = 1 To PipetteNumber
        Free = Free And Not (Pipette(i).Used)
    Next i
    
End Property

Public Property Get Worktable() As clsWorktableSetup

    Set Worktable = pWorktable

End Property

Public Property Set Worktable(c As clsWorktableSetup)

    Set pWorktable = c

End Property

Private Function ConstructVolumeArray(InpVolumes() As Variant)

'sets up the string that defines the volumes to be used in an aspirate / dispense command

End Function
Private Function DefineVolumesString(InpVolumes As Variant) As String
    
    Const conEmpty = "0"
    Const conDelimiter = ","
    'zero without quotes passed on to the final parameter
    'defined by the Gemini Worklist set scripting rules
    'otherwise, nonzero volumes have quotes around them!
    
    Dim i As Integer, j As Integer
    Dim VolumeArray(1 To conMaxPipetteNumber) As Double
    Dim StringArray(1 To conMaxPipetteNumber) As String
    
    'initialization
    j = 0
    
    For i = 1 To conMaxPipetteNumber
        VolumeArray(i) = 0
    Next i
    
    'construction of the intermediate - mostly unnecessary but oh well
    For i = LBound(InpVolumes) To UBound(InpVolumes)
        j = j + 1
        VolumeArray(j) = CDbl(InpVolumes(i))
    Next i
    
    'construction of the string intermediate
    For i = 1 To conMaxPipetteNumber
        If VolumeArray(i) > 0 Then
            StringArray(i) = conQ & CStr(Round(VolumeArray(i), 2)) & conQ
        Else
            StringArray(i) = conEmpty
        End If
    Next i
    
    'assembly of the final command
    DefineVolumesString = Join(StringArray, conDelimiter)

End Function

'****************************************************************************************************
Private Function LiquidCommand(CommandType As String, _
                        Worktable As clsWorktableSetup, _
                        tipMask As String, _
                        liquidClass As String, _
                        VolumesString As String, _
                        Grid As Integer, _
                        Site As Integer, _
                        Spacing As Integer, _
                        WellSelectionString As String _
                        ) As String


'====================================================================================================
'
'Juraj Ahel, 2016-03-09
'Last update 2016-03-09
'====================================================================================================
                    
Const conParametersCount As Integer = 8
Const connoOfLoopOptions As Integer = 0
Const conSeparator As String = ","
                    
Dim Parameters(1 To 8) As String
Dim Parameter0 As String

Select Case UCase(CommandType)
    Case "ASPIRATE", "A", "1"
        Parameter0 = "Aspirate"
    Case "DISPENSE", "D", "2"
        Parameter0 = "Dispense"
    Case Else
        '#TODO Give error
End Select

Parameters(1) = CStr(tipMask)
Parameters(2) = conQ & liquidClass & conQ
Parameters(3) = VolumesString
Parameters(4) = CStr(Grid)
Parameters(5) = CStr(Site)
Parameters(6) = CStr(Spacing)
Parameters(7) = conQ & WellSelectionString & conQ
Parameters(8) = CStr(connoOfLoopOptions)

LiquidCommand = Parameter0 & "(" & Join(Parameters, conSeparator) & ");"
                    
                    
End Function

'****************************************************************************************************
Public Sub RefreshPipettes()
'====================================================================================================
'Cleans up the pipettes' states
'Juraj Ahel, 2016-03-09
'Last update 2016-03-09
'====================================================================================================

    Dim i As Integer
    
    For i = 1 To PipetteNumber
        With Pipette(i)
            If .Sample Is Nothing Then
                .Count = 0
                .Volume = 0
            Else
                If .Count = 0 Then
                    Set .Sample = Nothing
                    .Volume = 0
                End If
            End If
        End With
    Next i
    
End Sub

Private Sub ResetArray(InpArray As Variant, Optional targetValue As Variant = 0)

Dim i As Integer

For i = LBound(InpArray) To UBound(InpArray)
    InpArray(i) = targetValue
Next i

End Sub
'****************************************************************************************************
Public Function AspirateCommand(Optional CommandType As String = "Default") As String

'====================================================================================================
'
'Juraj Ahel, 2016-01-xx
'Last update 2016-03-18
'====================================================================================================

'1 iterate through pipettes, check for each pipette whether it's used
'2 find its required sample
'3 locate the required sample in the Worktable, getting Grid, site, and well
'  and construct the string that defines the well(s) used
'3b check for uniqueness, just to be sure, give critical error if not unique
'4? try to suck more than one in parallel if by chance possible
'5 assemble the pipette command

Const conDefaultLiquidClass As String = "Water"
Const conDefaultSpacing As Integer = 1

Dim i As Integer

Dim tipMask As String
Dim liquidClass As String
Dim VolumesString As String
Dim Grid As Integer
Dim Site As Integer
Dim Spacing As Integer
Dim WellSelectionString As String
Dim tempString As String

Dim tempLocation As clsWorktableSampleLocation
Dim tempPipetting As clsPipettingCommand

Dim Volumes(1 To conMaxPipetteNumber) As Double

liquidClass = conDefaultLiquidClass

ResetArray Volumes, 0

For i = 1 To PipetteNumber
    
    If Pipette(i).Used Then
        Worktable.Locate Pipette(i).Sample
    End If
    
Next i

tempString = ""

For i = 1 To PipetteNumber

    With Pipette(i)
    
        If .Used Then
        
            Set tempLocation = .Sample.Location
            
            If Not (tempLocation Is Nothing) Then
                
                tipMask = BitFlag(CStr(i), conMaxPipetteNumber)
                
                ResetArray Volumes, 0
                Volumes(i) = .Volume
                VolumesString = DefineVolumesString(Volumes)
                
                With .Sample.Location
                    Grid = .Grid
                    Site = .Site
                End With
                
                Spacing = conDefaultSpacing
                
                WellSelectionString = tempLocation.Container.DefineWellSelectionString(.Sample)
                                    
                Set tempPipetting = New clsPipettingCommand
                                    
                tempPipetting.Define "Aspirate", tipMask, liquidClass, VolumesString, Grid, Site, Spacing, WellSelectionString
                                                    
                'tempString = tempString & _
                '    LiquidCommand("Aspirate", Worktable, tipMask, liquidClass, VolumesString, Grid, Site, Spacing, WellSelectionString) & _
                '    vbCrLf
                
                tempString = tempString & tempPipetting.FullCommandText & vbCrLf
                
                AddPipettingStep tempPipetting
                
                Set tempPipetting = Nothing
                    
            End If
        End If
        
        Set tempLocation = Nothing
        
    End With
    
Next i
          
AspirateCommand = Commands.Output

Commands.Purge

Set tempLocation = Nothing
Set tempPipetting = Nothing

End Function

'****************************************************************************************************
Private Function FindSampleInPlateColumn( _
            ByRef Sample As clsSample, _
            ByRef InputContainer As clsWorktableSeqPlate, _
            ByVal CurrentColumn As Integer, _
            Optional ByVal StartingRow As Integer = 1 _
            ) As Integer
'====================================================================================================
'
'Juraj Ahel, 2016-03-09
'Last update 2016-03-14
'====================================================================================================
    
    Const conMaxRow = 8
    
    Dim i As Integer
    
    Dim FinalRowNo As Integer
    
    Dim CurrentRow As Integer
    Dim CurrentSample As String
    Dim tempIndex As Integer
    
    Dim tempFound As Boolean
          
    'FinalRowNo = conMaxRow 'replace later with info from InputContainer
    FinalRowNo = InputContainer.maxRow
    
    If Sample Is Nothing Then 'just in case, catch null input
        FindSampleInPlateColumn = 0
        GoTo 999
    End If
    
    'Parsing inputs
    Select Case StartingRow
        Case Is <= 0
            StartingRow = 1
        Case Is > PipetteNumber
            FindSampleInPlateColumn = 0
            GoTo 999
    End Select
           
    CurrentSample = Sample.Name
    CurrentRow = StartingRow
    tempFound = False
    
    'search for it
    Do While Not tempFound And CurrentRow <= FinalRowNo
        
        With InputContainer 'go row by row in the container and try to find the sample
        
            tempIndex = .GetIndexFromXY(CurrentColumn, CurrentRow)
            
            Select Case .State(tempIndex)
                Case eTaken, ePrimerPipetted 'if sample hasn't been pipetted yet
                    If (.Well(tempIndex).Template.Name = CurrentSample) Then
                        tempFound = True
                    End If
                    
                Case Else
                    'do nothing
            End Select
            
            If Not tempFound Then
                CurrentRow = CurrentRow + 1
            End If
                  
        End With
        
    Loop
    
    If tempFound Then
        FindSampleInPlateColumn = CurrentRow
    Else '0 means it ain't found
        FindSampleInPlateColumn = 0
    End If

999 End Function

'****************************************************************************************************
Private Function FindBestSpacing( _
            MainPipette As Integer, _
            InputContainer As clsWorktableSeqPlate, _
            CurrentColumn As Integer, _
            MainRow As Integer _
            ) As Integer

'====================================================================================================
'find the spacing which allows to pipette more then 1 sample at 1 (sometimes possible)
'Juraj Ahel, 2016-03-09
'Last update 2016-03-16
'====================================================================================================
    
    '[DECLARATION]
    
    'temp vars
    Dim tempSampleNo As Integer
    Dim tempSampleLocation As Integer
    Dim tempResult As Integer
                
    'real vars
    Dim PipetteDistance As Integer
    Dim SampleSpacing As Integer
                
    'flow control
    Dim IFoundAnotherSampleDownstream As Boolean
    
    'iterators
    Dim i As Integer
        
    '[INITIALIZATION]
    IFoundAnotherSampleDownstream = False
    tempSampleLocation = 0
    'tempSampleNo = MainPipette + 1
    tempSampleNo = FirstUsedPipette(AfterPipetteNo:=MainPipette)
    SampleSpacing = 1
    PipetteDistance = tempSampleNo - MainPipette
            
    'until a sample is found or bust
    Do While (tempSampleNo <= PipetteNumber) _
        And (tempSampleNo > 0) _
        And (Not IFoundAnotherSampleDownstream)
        
        tempSampleLocation = FindSampleInPlateColumn( _
                                Sample:=Pipette(tempSampleNo).Sample, _
                                InputContainer:=InputContainer, _
                                CurrentColumn:=CurrentColumn, _
                                StartingRow:=MainRow + PipetteDistance)
        
        If tempSampleLocation <> 0 Then 'if the current sample was found downstream
            IFoundAnotherSampleDownstream = True
        Else                            'otherwise try the next one
            tempSampleNo = FirstUsedPipette(AfterPipetteNo:=tempSampleNo)
            PipetteDistance = tempSampleNo - MainPipette
        End If
        
        If IFoundAnotherSampleDownstream Then
            'test if it's physically achievable (spacing * pipette distance) = sample distance
            SampleSpacing = tempSampleLocation - MainRow
            
            tempResult = 0
            For i = 1 To conMaximumSpacing
                If PipetteDistance * i = SampleSpacing Then
                    tempResult = i
                End If
            Next i
            
            If tempResult = 0 Then
                IFoundAnotherSampleDownstream = False 'I found it, but it can't be done
            End If
        End If
                
    Loop
    
    If IFoundAnotherSampleDownstream Then
        FindBestSpacing = tempResult
    Else
        FindBestSpacing = 1 'might as well leave it at default. Later I could change it to "last used spacing" also
    End If
    
End Function

'****************************************************************************************************
Private Sub DispenseCommandHere( _
            ByRef SeqPlate As clsWorktableSeqPlate, _
            ByVal ReferencePipetteNo As Integer, _
            ByVal ReferenceRow As Integer, _
            ByVal TargetColumn As Integer, _
            ByVal Spacing As Integer)

'====================================================================================================
'Cross references the samples in the LIHA pipettes and the chosen SeqPlate, aligned on ReferencePipetteNo
'in the LIHA and ReferenceRow in the plate, using fixed spacing and column.
'Sets up pipetting for all matching samples
'Juraj Ahel, 2016-03-15
'Last update 2016-03-15
'====================================================================================================
    
    Const conSampleVolume = 3
    Const conPrimerVolume = 4
    
    Dim PipetteToPlateOffset As Integer
    
    Dim i As Integer, j As Integer
    Dim tempRow As Integer
    
    Dim maxRow As Integer
    Dim maxColumn As Integer
    
    Dim tempFlag As String
    Dim VolumeToPipette As Double
    
    Dim Volumes(1 To conMaxPipetteNumber) As Double
    Dim Plate2D() As Byte
    
    Dim tempSampleName As String
    Dim tempcPipetting As clsPipettingCommand
    
    
    VolumeToPipette = conSampleVolume
    ResetArray Volumes, 0
    
    PipetteToPlateOffset = (ReferenceRow - 1) - (Spacing * (ReferencePipetteNo - 1))
    maxRow = SeqPlate.maxRow
    maxColumn = SeqPlate.maxColumn
    
    ReDim Plate2D(1 To maxColumn, 1 To maxRow)
    
    For i = 1 To maxColumn
        For j = 1 To maxRow
            Plate2D(i, j) = 0
        Next j
    Next i
    
    For i = 1 To PipetteNumber
        
        If Pipette(i).Used Then 'if the pipette is even used
            
            tempRow = 1 + PipetteToPlateOffset + (i - 1) * Spacing
            
            If tempRow > 0 And tempRow <= maxRow Then 'if we are even in the plate
                                               
                If Not (SeqPlate.WellXY(TargetColumn, tempRow) Is Nothing) Then
                    
                    tempSampleName = SeqPlate.WellXY(TargetColumn, tempRow).Template.Name
                    
                    If tempSampleName = Pipette(i).Sample.Name Then
                    
                        tempFlag = tempFlag & i
                        Volumes(i) = VolumeToPipette
                                            
                        Pipette(i).DecreaseVolume VolumeToPipette 'empty the virtual pipette correspondingly
                        SeqPlate.ToggleTemplatePipetted tempRow, TargetColumn 'update the state of the sample, maybe problematic as it is byref...
                        
                        Plate2D(TargetColumn, tempRow) = 1 'the bit field needed to construct the targeting string for gemini
                                                
                    End If
                    
                End If
                
            End If
            
        End If
        
    Next i
    
    Set tempcPipetting = New clsPipettingCommand
    
    With tempcPipetting
        .CommandType = "D"
        .tipMask = BitFlag(tempFlag, conMaxPipetteNumber)
        .liquidClass = conDefaultLiquidClass
        .VolumesString = DefineVolumesString(Volumes)
        .Spacing = Spacing
        .Grid = SeqPlate.Grid
        .Site = SeqPlate.Site
        .WellSelectionString = GetWellString(Plate2D)
    End With
    
    AddPipettingStep tempcPipetting
    
    Set tempcPipetting = Nothing
    
End Sub

'****************************************************************************************************
Public Sub Purge(Optional InputParameter As Variant = "0")
'====================================================================================================
'cleans out the pipettes
'Juraj Ahel, 2016-03-16
'Last update 2016-03-16
'====================================================================================================
'0 means purge all

    Dim i As Integer

    Select Case UCase(InputParameter)
        Case "0"
            'wash all
    End Select
    
    '###TEMP
    
    For i = 1 To PipetteNumber
        Pipette(i).CleanMe
    Next i

End Sub

'****************************************************************************************************
Public Function WashCommand( _
        Optional PipetteSelection As Variant = "ALL", _
        Optional WashingOption As String = "DEFAULT" _
        ) As String
'====================================================================================================
'defines the Gemini worklist command for Washing
'Juraj Ahel, 2016-03-16
'Last update 2016-03-16
'====================================================================================================

    Select Case UCase(WashingOption)
        Case "DEFAULT"
    End Select


End Function


'****************************************************************************************************
Public Function DispenseCommand() As String
'====================================================================================================
'
'Juraj Ahel, 2016-03-09
'Last update 2016-03-15
'====================================================================================================
    
    Const conMaxRow As Integer = 8
    
    Dim i As Integer
    
    Dim Volumes(1 To conMaxPipetteNumber) As Double
    
    Dim SequencingPlatesNumber As Integer
    Dim FinalPlates() As clsWorktableSeqPlate
        
    'temporary variables
    Dim MainPipetteIndex As Integer
    Dim Spacing As Integer
     
    Dim CurrentColumn As Integer
    Dim CurrentContainer As Integer
    Dim CurrentRow As Integer
    Dim CurrentSample As String
    
    Dim StartRow As Integer
    Dim StartColumn As Integer
    Dim StartContainer As Integer
        
    Dim FinalRow As Integer
    Dim FinalColumn As Integer
    Dim FinalContainer As Integer
    
    'flow control
    Dim TryNextFlow As Boolean
        
    '[INITIALIZATION]
    
    'Loads in the Sequencing Plates into the array
    SequencingPlatesNumber = Worktable.SequencingPlates.Count
    ReDim FinalPlates(1 To SequencingPlatesNumber)
    
    For i = 1 To SequencingPlatesNumber
        Set FinalPlates(i) = Worktable.SequencingPlates.Item(i)
    Next i
    
    ResetArray Volumes, 0
    
    'walk the first nonempty pipette through Sequencing plates and see if there is something to be pipetted
        
    StartRow = 1
    StartColumn = 1
    StartContainer = 1
    
    '***********************************************
    '[INITIALIZATION] set first sample
    
    MainPipetteIndex = FirstUsedPipette(0)
    CurrentSample = Pipette(MainPipetteIndex).Sample.Name
    CurrentContainer = StartContainer
    CurrentColumn = StartColumn
    CurrentRow = StartRow
    
    FinalContainer = UBound(FinalPlates)
    FinalRow = FinalPlates(CurrentContainer).maxRow
    FinalColumn = FinalPlates(CurrentContainer).maxColumn
    
    Do 'repeat until the pipette is empty [0]
        
        Do 'repeat while I can still find this sample in this column
            
            CurrentRow = FindSampleInPlateColumn(Sample:=Pipette(MainPipetteIndex).Sample, _
                                                InputContainer:=FinalPlates(CurrentContainer), _
                                                CurrentColumn:=CurrentColumn, _
                                                StartingRow:=CurrentRow)
            
            If CurrentRow <> 0 Then
                
                If CurrentRow <> FinalRow Then
                    Spacing = FindBestSpacing(MainPipette:=MainPipetteIndex, _
                                        InputContainer:=FinalPlates(CurrentContainer), _
                                        CurrentColumn:=CurrentColumn, _
                                        MainRow:=CurrentRow)
                Else
                    Spacing = conDefaultSpacing
                End If
                
                'construct the actual command and update all the object states
                DispenseCommandHere _
                    SeqPlate:=FinalPlates(CurrentContainer), _
                    ReferencePipetteNo:=MainPipetteIndex, _
                    ReferenceRow:=CurrentRow, _
                    TargetColumn:=CurrentColumn, _
                    Spacing:=Spacing
                
            End If
                    
        Loop Until CurrentRow = 0 Or Free
            
        If Free Then 'if the LIHA is emptied, don't check for samples further
            TryNextFlow = False
        Else
            TryNextFlow = True
        End If
        
        'try next pipette in same column
        If TryNextFlow And MainPipetteIndex < PipetteNumber Then
            
            MainPipetteIndex = FirstUsedPipette(AfterPipetteNo:=MainPipetteIndex)
            
            If MainPipetteIndex <> 0 Then
            
                TryNextFlow = False
                        
            Else
            
                TryNextFlow = True
                MainPipetteIndex = FirstUsedPipette
                
            End If
        End If
                
        'try next column in same container
        If TryNextFlow Then
            If CurrentColumn < FinalColumn Then
            
                CurrentColumn = CurrentColumn + 1
                TryNextFlow = False
                
            Else
            
                CurrentColumn = StartColumn
                TryNextFlow = True
            End If
        End If
        
        'try next container
        If TryNextFlow Then
            If CurrentContainer < FinalContainer Then
            
                CurrentContainer = CurrentContainer + 1
                FinalRow = FinalPlates(CurrentContainer).maxRow
                FinalColumn = FinalPlates(CurrentContainer).maxColumn
                TryNextFlow = False
                
            Else 'if there is still a sample to pipette and no target, something is wrong
            
                ErrorReportGlobal 70, "clsLIHA: sample in pipette, but not found in target plate!"
            End If
        End If
            
    Loop While Not Free '[0]
    
    DispenseCommand = Commands.Output
    
    Commands.Purge
    
    For i = LBound(FinalPlates) To UBound(FinalPlates)
        Set FinalPlates(i) = Nothing
    Next i
    


End Function



