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
Private Const conTemplateLiquidClass As String = "Seq_Template"
Private Const conPrimerLiquidClass As String = "Seq_Primer"
Private Const conDefaultSpacing As Integer = 1

Private Const conDefaultSampleVolume As Double = 3
Private Const conDefaultPrimerVolume As Double = 4
    
'programming constants
Private Const conQ As String = """"

'system constants
Private Const conDefaultPipetteNumber As Integer = 8 'setup of our TECAN EVO instrument (2016-03-09)
Private Const conMaxPipetteNumber As Integer = 12   'defined by TECAN EVO instrument
Private Const conMaximumSpacing As Integer = 4      'defined by TECAN EVO instrument
Private Const conMaximumAspirationVolume As Double = 950      'defined by TECAN EVO instrument diluters

'system definition
Private pPipetteNumber As Integer

'data variables
Private pWorktable As clsWorktableSetup
Private pPipettes() As clsPipette
Private pCommands As clsCommandSequence
Private pVolumes() As Double

Public TargetSampleVolume As Double
Public TargetPrimerVolume As Double

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
    
    TargetSampleVolume = conDefaultSampleVolume
    TargetPrimerVolume = conDefaultPrimerVolume

End Sub

Private Sub Class_Terminate()

    Dim i As Integer
    
    For i = 1 To PipetteNumber
        Set pPipettes(i) = Nothing
    Next i
    
    Set pWorktable = Nothing
    Set pCommands = Nothing

End Sub

Public Property Get MaxAllowedAspirationVolume() As Double
    MaxAllowedAspirationVolume = conMaximumAspirationVolume
End Property

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
Private Function LiquidCommandString(CommandType As String, _
                        TipMask As String, _
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
                    
Dim Parameters(1 To conParametersCount) As String
Dim Parameter0 As String

Select Case UCase(CommandType)
    Case "ASPIRATE", "A", "1"
        Parameter0 = "Aspirate"
    Case "DISPENSE", "D", "2"
        Parameter0 = "Dispense"
    Case Else
        '#TODO Give error
End Select

Parameters(1) = CStr(TipMask)
Parameters(2) = conQ & liquidClass & conQ
Parameters(3) = VolumesString
Parameters(4) = CStr(Grid)
Parameters(5) = CStr(Site)
Parameters(6) = CStr(Spacing)
Parameters(7) = conQ & WellSelectionString & conQ
Parameters(8) = CStr(connoOfLoopOptions)

LiquidCommandString = Parameter0 & "(" & Join(Parameters, conSeparator) & ");"
                    
                    
End Function

'****************************************************************************************************
Private Function WashCommandString( _
                            ByVal TipMask As String, _
                            ByVal WasteGrid As Integer, _
                            ByVal WasteSite As Integer, _
                            ByVal CleanerGrid As Integer, _
                            ByVal CleanerSite As Integer, _
                            ByVal WasteVolume As Integer, _
                            ByVal WasteTimeDelay As Integer, _
                            ByVal CleanerVolume As Integer, _
                            ByVal CleanerTimeDelay As Integer, _
                            ByVal AirgapVolume As Integer, _
                            ByVal AirgapSpeed As Integer, _
                            ByVal RetractSpeed As Integer, _
                            ByVal FastWash As Byte, _
                            ByVal PinchValve As Byte, _
                            ByVal ActiveTipFrequency As Integer _
                            ) As String

'====================================================================================================
'
'Juraj Ahel, 2016-08-01
'Last update 2016-08-01
'====================================================================================================
                    
    Const conParametersCount As Integer = 15
    Const conSeparator As String = ","
                        
    Dim Parameters(1 To conParametersCount) As String
    Dim Parameter0 As String
    
    Parameter0 = "Wash"
    
    Parameters(1) = CStr(TipMask)
    Parameters(2) = CStr(WasteGrid)
    Parameters(3) = CStr(WasteSite)
    Parameters(4) = CStr(CleanerGrid)
    Parameters(5) = CStr(CleanerSite)
    Parameters(6) = conQ & CStr(WasteVolume) & conQ
    Parameters(7) = CStr(WasteTimeDelay)
    Parameters(8) = conQ & CStr(CleanerVolume) & conQ
    Parameters(9) = CStr(CleanerTimeDelay)
    Parameters(10) = CStr(AirgapVolume)
    Parameters(11) = CStr(AirgapSpeed)
    Parameters(12) = CStr(RetractSpeed)
    Parameters(13) = CStr(FastWash)
    Parameters(14) = CStr(PinchValve)
    Parameters(15) = CStr(ActiveTipFrequency)
    
    WashCommandString = Parameter0 & "(" & Join(Parameters, conSeparator) & ");"
                        
                    
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

Public Function AspirateCommand(ByVal CommandType As String) As String

'====================================================================================================
'
'Juraj Ahel, 2016-01-xx
'Last update 2016-03-18
'====================================================================================================
'2016-08-01 big update, first version that works
'2016-08-03 made multiple pipettes aspirate at the same time whenever possible, so far restricted on spacing = 1
    
    '1 iterate through pipettes, check for each pipette whether it's used
    '2 find its required sample
    '3 locate the required sample in the Worktable, getting Grid, site, and well
    '  and construct the string that defines the well(s) used
    '3b check for uniqueness, just to be sure, give critical error if not unique
    '4? try to suck more than one in parallel if by chance possible
    '5 assemble the pipette command
    
    Dim i As Integer
    Dim j As Integer
    
    Dim TipMask As String
    Dim liquidClass As String
    Dim VolumesString As String
    Dim Grid As Integer
    Dim Site As Integer
    Dim Spacing As Integer
    Dim WellSelectionString As String
    'Dim tempString As String
    
    Dim tempLocation As clsWorktableSampleLocation
    Dim tempPipetting As clsPipettingCommand
    
    Dim Volumes(1 To conMaxPipetteNumber) As Double
    Dim IShouldPipetteThis(1 To conMaxPipetteNumber) As Boolean
    
    Dim tSamplesCollection As VBA.Collection
    Dim tempPipetteString As String
    
    Dim LastColumn As Integer
    Dim LastRow As Integer
    Dim lastContainer As clsWorktableContainer
    
    'liquidClass = conDefaultLiquidClass
    
    Select Case UCase$(CommandType)
        Case "SAMPLES", "S"
            liquidClass = conTemplateLiquidClass
        Case "PRIMERS", "P"
            liquidClass = conPrimerLiquidClass
    End Select
            
    'initialize variables
        ResetArray Volumes, 0
        ResetArray IShouldPipetteThis, False
        
        Set lastContainer = Nothing
        LastColumn = -1
        LastRow = -1
        
        Spacing = conDefaultSpacing
        
        For i = 1 To PipetteNumber
        
            If Pipette(i).Used Then
                Worktable.Locate Pipette(i).Sample
                IShouldPipetteThis(i) = True
            End If
        
        Next i
    
    'carry out the pipetting
        For j = 1 To PipetteNumber
        
            If IShouldPipetteThis(j) Then
            
                ResetArray Volumes, 0
                tempPipetteString = ""
                Set tSamplesCollection = New VBA.Collection
            
                For i = j To PipetteNumber
                    With Pipette(i)
                        If .Used Then
                            
                            Set tempLocation = .Sample.Location
                            
                            If Not (tempLocation Is Nothing) Then
                                
                                'if it's the first tested location, or the next position fits
                                If (i = j) _
                                    Or ((tempLocation.Container Is lastContainer) _
                                    And (tempLocation.WellColumn = LastColumn) _
                                    And (tempLocation.WellRow = j + Spacing * (i - j))) Then
                                                            
                                    If .Volume <= conMaximumAspirationVolume Then
                                        Volumes(i) = .Volume
                                    Else
                                        Call ErrorReportGlobal(5090, "Tried to aspirate more than allowed volume (" & .Volume & " ul)!")
                                    End If
                                    
                                    tSamplesCollection.Add .Sample
                                                                
                                    tempPipetteString = tempPipetteString & CStr(i)
                                    IShouldPipetteThis(i) = False
                                    
                                    Set lastContainer = tempLocation.Container
                                    LastColumn = tempLocation.WellColumn
                                    LastRow = tempLocation.WellRow
                                
                                End If
                                    
                            End If
                            
                        End If
                    End With
                Next i
                
                Set tempLocation = tSamplesCollection.Item(1).Location
                            
                With tempLocation
                    Grid = .Grid
                    Site = .Site
                    WellSelectionString = .Container.DefineWellSelectionString(tSamplesCollection)
                End With
                
                TipMask = BitFlag(tempPipetteString, conMaxPipetteNumber)
                VolumesString = DefineVolumesString(Volumes)
                                    
                Set tempPipetting = New clsPipettingCommand
                tempPipetting.Define "Aspirate", TipMask, liquidClass, VolumesString, Grid, Site, Spacing, WellSelectionString
                
                AddPipettingStep tempPipetting
                
                Set tempLocation = Nothing
                Set tempPipetting = Nothing
                Set tSamplesCollection = Nothing
                
                Set lastContainer = Nothing
                LastColumn = -1
                LastRow = -1
                
            End If
            
        Next j
              
    'output the results
        If Not Commands Is Nothing Then
            AspirateCommand = Commands.Output
            Commands.Purge
        End If
    
    'clean up
        Set tSamplesCollection = Nothing
        Set tempLocation = Nothing
        Set tempPipetting = Nothing
        Set lastContainer = Nothing

End Function

'****************************************************************************************************
Private Function FindSampleInPlateColumn( _
            ByRef Sample As clsSample, _
            ByRef TargetContainer As clsWorktableSeqPlate, _
            ByVal CurrentColumn As Integer, _
            ByVal StartingRow As Integer, _
            ByVal CommandType As String _
            ) As Integer
'====================================================================================================
'
'Juraj Ahel, 2016-03-09
'Last update 2016-03-14
'====================================================================================================
'2016-08-01 add support for pipetting primers as well
    
    Const conMaxRow = 8
    
    Dim i As Integer
    
    Dim FinalRowNo As Integer
    
    Dim CurrentRow As Integer
    Dim CurrentSample As String
    Dim tempIndex As Integer
    
    Dim tempFound As Boolean
    Dim UseWell As Boolean
    
    Dim tempSample As String
          
    'FinalRowNo = conMaxRow 'replace later with info from TargetContainer
    FinalRowNo = TargetContainer.maxRow
    
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
        
        With TargetContainer 'go row by row in the container and try to find the sample
        
            tempIndex = .GetIndexFromXY(CurrentColumn, CurrentRow)
            
            UseWell = False
                            
            'check whether the current well is something that needs to be checked
                Select Case .State(tempIndex)
                
                    Case eTaken
                        UseWell = True
                        
                    Case ePrimerPipetted
                        Select Case UCase(CommandType)
                            Case "SAMPLES", "S"
                                UseWell = True
                        End Select
                        
                    Case eTemplatePipetted
                        Select Case UCase(CommandType)
                            Case "PRIMERS", "P"
                                UseWell = True
                        End Select
                        
                End Select
            
            'if the well needs to be checked, check if it matches!
                If UseWell Then
                    
                    Select Case UCase$(CommandType)
                        Case "SAMPLES", "S"
                            tempSample = .Well(tempIndex).Template.Name
                        Case "PRIMERS", "P"
                            tempSample = .Well(tempIndex).Primer.Name
                    End Select
                    
                    If (tempSample = CurrentSample) Then
                        tempFound = True
                    End If
                    
                End If
            
            'otherwise, check next well in row
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
            MainRow As Integer, _
            CommandType As String _
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
                                TargetContainer:=InputContainer, _
                                CurrentColumn:=CurrentColumn, _
                                StartingRow:=MainRow + PipetteDistance, _
                                CommandType:=CommandType)
        
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
                tempSampleNo = 0 'not sure if this is optimal in EVERY case but should be fine
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
            ByVal Spacing As Integer, _
            ByVal CommandType As String)

'====================================================================================================
'Cross references the samples in the LIHA pipettes and the chosen SeqPlate, aligned on ReferencePipetteNo
'in the LIHA and ReferenceRow in the plate, using fixed spacing and column.
'Sets up pipetting for all matching samples
'Juraj Ahel, 2016-03-15
'Last update 2016-03-15
'====================================================================================================
'2016-08-02 big update, made it work
    
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
    
    Dim tempState As enPlateStates
    
    Select Case UCase$(CommandType)
        Case "SAMPLES", "S"
            VolumeToPipette = TargetSampleVolume
        Case "PRIMERS", "P"
            VolumeToPipette = TargetPrimerVolume
    End Select
        
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
                                               
                If Not (SeqPlate.WellXY(TargetColumn, tempRow) Is Nothing) Then 'if something is supposed to be in this well at all
                
                    tempState = SeqPlate.State(SeqPlate.GetIndexFromXY(TargetColumn, tempRow))
                
                    Select Case UCase$(CommandType)
                        Case "SAMPLES", "S"
                            If tempState = eTaken Or tempState = ePrimerPipetted Then
                                tempSampleName = SeqPlate.WellXY(TargetColumn, tempRow).Template.Name
                            Else
                                GoTo NextIteration
                            End If
                                
                        Case "PRIMERS", "P"
                            If tempState = eTaken Or tempState = eTemplatePipetted Then
                                tempSampleName = SeqPlate.WellXY(TargetColumn, tempRow).Primer.Name
                            Else
                                GoTo NextIteration
                            End If
                    End Select
                    
                    If tempSampleName = Pipette(i).Sample.Name Then 'if sample from THIS pipette is supposed to be in thiswell
                    
                        tempFlag = tempFlag & i
                        Volumes(i) = VolumeToPipette
                                            
                        Pipette(i).DecreaseVolume VolumeToPipette 'empty the virtual pipette correspondingly
                        
                        Select Case UCase$(CommandType)
                            Case "SAMPLES", "S"
                                SeqPlate.ToggleTemplatePipetted tempRow, TargetColumn 'update the state of the sample, maybe problematic as it is byref...
                            Case "PRIMERS", "P"
                                SeqPlate.TogglePrimerPipetted tempRow, TargetColumn 'update the state of the sample, maybe problematic as it is byref...
                        End Select
                        
                        Plate2D(TargetColumn, tempRow) = 1 'the bit field needed to construct the targeting string for gemini
                                                
                    End If
                    
                End If
                
            End If
            
        End If
        
NextIteration:
    Next i
    
    Set tempcPipetting = New clsPipettingCommand
    
    With tempcPipetting
        .CommandType = "D"
        .TipMask = BitFlag(tempFlag, conMaxPipetteNumber)

        Select Case UCase$(CommandType)
            Case "SAMPLES", "S"
                .liquidClass = conTemplateLiquidClass
            Case "PRIMERS", "P"
                .liquidClass = conPrimerLiquidClass
        End Select
        
        .VolumesString = DefineVolumesString(Volumes)
        .Spacing = Spacing
        .Grid = SeqPlate.Grid
        .Site = SeqPlate.Site
        .WellSelectionString = GetWellString(Plate2D)
    End With
    
    AddPipettingStep tempcPipetting
    
    Set tempcPipetting = Nothing
    
End Sub

Private Sub Pipette_CleanedWhileNotEmpty(SampleName As String, Volume As Double)

    Debug.Print ("Pipette was cleaned while there was still some sample in it! (" & SampleName & ": " & Volume & " uL)")

End Sub

'****************************************************************************************************
Public Sub Purge(Optional InputParameter As Variant = "0")
'====================================================================================================
'cleans out the pipettes
'Juraj Ahel, 2016-03-16
'Last update 2016-03-16
'====================================================================================================
'2016-08-01 clean up a bit, add debug logging
'0 means purge all

    Dim i As Integer
    Dim tempPipette As clsPipette

    Select Case UCase(InputParameter)
        Case "0"
            'wash all
    End Select
    
    '###TEMP
    
    For i = 1 To PipetteNumber
    
        Set tempPipette = Pipette(i)
        
        With tempPipette
            If .IsDirty Then
                .CleanMe
                Debug.Print ("Pipette " & i & " virtually purged.")
            End If
        End With
        
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
        
    'constants
        Const conDefaultDelay As Integer = 500 '500 ms delay
        Const conDefaultAirGap As Integer = 20
        Const conDefaultAirGapSpeed As Integer = 70
        Const conDefaultRetractSpeed As Integer = 30
        Const conDefaultActiveFreq As Integer = 1000
    
    'Gemini variables
        Dim TipMask As Integer
        
        Dim WasteGrid As Integer
        Dim WasteSite As Integer
        Dim CleanerGrid As Integer
        Dim CleanerSite As Integer
        
        Dim WasteVolume As Integer
        Dim WasteTimeDelay As Integer
        Dim CleanerVolume As Integer
        Dim CleanerTimeDelay As Integer
        
        Dim AirgapVolume As Integer
        Dim AirgapSpeed As Integer
        Dim RetractSpeed As Integer
        
        Dim FastWash As Byte
        Dim PinchValve As Byte
        
        Dim ActiveTipFrequency As Integer
        
    'programming variables
        Dim tString As String
        Dim i As Integer
    
    'define values that will almost certainly be constant
        WasteTimeDelay = conDefaultDelay
        CleanerTimeDelay = conDefaultDelay
        AirgapVolume = conDefaultAirGap
        AirgapSpeed = conDefaultAirGapSpeed
        RetractSpeed = conDefaultRetractSpeed
        ActiveTipFrequency = conDefaultActiveFreq
    
    
    With Worktable
    
        WasteGrid = .GetGrid("WASTE")
        WasteSite = .GetSite("WASTE")
            
        Select Case UCase(WashingOption)
        
            Case "DEFAULT"
                CleanerGrid = .GetGrid("CLEANERDEEP")
                CleanerSite = .GetSite("CLEANERDEEP")
                
            Case Else
                CleanerGrid = .GetGrid("CLEANERDEEP")
                CleanerSite = .GetSite("CLEANERDEEP")
                
        End Select
    
    End With
    
    PinchValve = 0
    
    'select only dirty pipettes
        For i = 1 To PipetteNumber
            If Pipette(i).IsDirty Then
                Debug.Print ("Pipette " & i & " wash command issued.")
                TipMask = TipMask + 2 ^ (i - 1)
            End If
        Next i
    
    
    'quick rinse with fastwash
        CleanerVolume = 5
        WasteVolume = 5
        FastWash = 1
        
        tString = WashCommandString( _
                    TipMask:=TipMask, _
                    WasteGrid:=WasteGrid, _
                    WasteSite:=WasteSite, _
                    CleanerGrid:=CleanerGrid, _
                    CleanerSite:=CleanerSite, _
                    WasteVolume:=WasteVolume, _
                    WasteTimeDelay:=WasteTimeDelay, _
                    CleanerVolume:=CleanerVolume, _
                    CleanerTimeDelay:=CleanerTimeDelay, _
                    AirgapVolume:=AirgapVolume, _
                    AirgapSpeed:=AirgapSpeed, _
                    RetractSpeed:=RetractSpeed, _
                    FastWash:=FastWash, _
                    PinchValve:=PinchValve, _
                    ActiveTipFrequency:=ActiveTipFrequency)
                    
        Commands.Append tString
    
    'proper wash
        CleanerVolume = 2.5
        WasteVolume = 2.5
        FastWash = 0
        tString = WashCommandString( _
                    TipMask:=TipMask, _
                    WasteGrid:=WasteGrid, _
                    WasteSite:=WasteSite, _
                    CleanerGrid:=CleanerGrid, _
                    CleanerSite:=CleanerSite, _
                    WasteVolume:=WasteVolume, _
                    WasteTimeDelay:=WasteTimeDelay, _
                    CleanerVolume:=CleanerVolume, _
                    CleanerTimeDelay:=CleanerTimeDelay, _
                    AirgapVolume:=AirgapVolume, _
                    AirgapSpeed:=AirgapSpeed, _
                    RetractSpeed:=RetractSpeed, _
                    FastWash:=FastWash, _
                    PinchValve:=PinchValve, _
                    ActiveTipFrequency:=ActiveTipFrequency)
                    
        Commands.Append tString
    
    
    'export resulting string
        WashCommand = Commands.Output
    
    'purge / clean the virtual pipettes
        Call Purge
    
    'purge commands from buffer
        Commands.Purge
    
End Function


'****************************************************************************************************
Public Function DispenseCommand(ByVal CommandType As String) As String
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
                                                TargetContainer:=FinalPlates(CurrentContainer), _
                                                CurrentColumn:=CurrentColumn, _
                                                StartingRow:=CurrentRow, _
                                                CommandType:=CommandType)
            
            If CurrentRow <> 0 Then
                
                If CurrentRow <> FinalRow Then
                    Spacing = FindBestSpacing(MainPipette:=MainPipetteIndex, _
                                        InputContainer:=FinalPlates(CurrentContainer), _
                                        CurrentColumn:=CurrentColumn, _
                                        MainRow:=CurrentRow, _
                                        CommandType:=CommandType)
                Else
                    Spacing = conDefaultSpacing
                End If
                
                Debug.Print (PrintMatrixYX(FinalPlates(CurrentContainer).PrimersArray))
                
                'construct the actual command and update all the object states
                DispenseCommandHere _
                    SeqPlate:=FinalPlates(CurrentContainer), _
                    ReferencePipetteNo:=MainPipetteIndex, _
                    ReferenceRow:=CurrentRow, _
                    TargetColumn:=CurrentColumn, _
                    Spacing:=Spacing, _
                    CommandType:=CommandType
                
            End If
                    
        Loop Until CurrentRow = 0 Or Not Pipette(MainPipetteIndex).Used
            
        If Free Then 'if the LIHA is emptied, don't check for samples further
            TryNextFlow = False
        Else
            TryNextFlow = True
        End If
        
        'try next pipette in same column
        If TryNextFlow Then
            
            'reset row
            CurrentRow = StartRow
            
            If MainPipetteIndex < PipetteNumber Then
            
                MainPipetteIndex = FirstUsedPipette(AfterPipetteNo:=MainPipetteIndex)
                                
                If MainPipetteIndex <> 0 Then
                    TryNextFlow = False
                Else
                    TryNextFlow = True
                End If
                
            Else
                
ActuallyFalse:  'try next search layer
                TryNextFlow = True
                
            End If
            
        End If
                
        'try next column in same container
        If TryNextFlow Then
        
            'reset row and pipette
            CurrentRow = StartRow
            MainPipetteIndex = FirstUsedPipette
            
            If CurrentColumn < FinalColumn Then
                        
                CurrentColumn = CurrentColumn + 1
                TryNextFlow = False
                
            Else
                
                TryNextFlow = True 'try next search layer
                
            End If
            
        End If
        
        'try next container
        If TryNextFlow Then
            
            'reset column, row, and pipette
            CurrentColumn = StartColumn
            CurrentRow = StartRow
            MainPipetteIndex = FirstUsedPipette
        
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



