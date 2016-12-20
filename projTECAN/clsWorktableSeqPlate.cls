VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorktableSeqPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pGrid As Integer
Private pSite As Integer
Private pType As String

Private pXDim As Integer
Private pYDim As Integer
Private pTotalWells As Integer

'Data collections
Private pPipettingStates() As enPlateStates
Private pSeqReactionList() As clsSeqReaction
Private pSequencingSamples As clsSeqSamples
Private pPrimers As clsPrimers

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    
    Dim i As Integer
        
        
    If pTotalWells > 0 Then 'this is jury-rigged. I need to properly check if the array is initialized!
    For i = LBound(pSeqReactionList) To UBound(pSeqReactionList)
        Set pSeqReactionList(i) = Nothing
    Next i
    End If
    
        
    Set pPrimers = Nothing
    Set pSequencingSamples = Nothing
    
End Sub

Public Property Get State(i As Integer) As enPlateStates
    State = pPipettingStates(i)
End Property

Public Property Let State(i As Integer, InputState As enPlateStates)
    pPipettingStates(i) = InputState
End Property

Public Property Get maxRow() As Integer
    maxRow = pYDim
End Property

Public Property Get maxColumn() As Integer
    maxColumn = pXDim
End Property

Public Property Get Grid() As Integer
    Grid = pGrid
End Property

Public Property Get Site() As Integer
    Site = pSite
End Property

Public Property Get PlateType() As String
    PlateType = pType
End Property

Public Property Get TotalNumber() As Integer
    TotalNumber = pTotalWells
End Property

Public Property Get PrimersArray() As Variant

    Dim i As Integer
    Dim j As Integer
    
    Dim tempArray() As String
    
    ReDim tempArray(1 To maxRow, 1 To maxColumn)
    
    For i = 1 To maxRow
        For j = 1 To maxColumn
            If Not WellXY(j, i) Is Nothing Then
                tempArray(i, j) = WellXY(j, i).Primer.Name
            End If
        Next j
    Next i
    
    PrimersArray = tempArray
    
End Property

Public Property Get TemplatesArray() As Variant

    Dim i As Integer
    Dim j As Integer
    
    Dim tempArray() As String
    
    ReDim tempArray(1 To maxRow, 1 To maxColumn)
    
    For i = 1 To maxRow
        For j = 1 To maxColumn
            If Not WellXY(j, i) Is Nothing Then
                tempArray(i, j) = WellXY(j, i).Template.Name
            End If
        Next j
    Next i
    
    TemplatesArray = tempArray
    
End Property

Public Property Get WellXY(ColumnNo As Integer, RowNo As Integer) As clsSeqReaction
    
    Dim tempIndex As Integer
    
    XYToWellIndex ColumnNo, RowNo, tempIndex
        
    Set WellXY = Well(tempIndex)
    
End Property

Public Sub TogglePrimerPipetted(InRow As Integer, InColumn As Integer)

    Dim tempIndex As Integer
    
    tempIndex = GetIndexFromXY(InColumn, InRow)

    Select Case State(tempIndex)
        Case eTaken
            State(tempIndex) = ePrimerPipetted
        Case eTemplatePipetted
            State(tempIndex) = eAllPipetted
        Case Else
            ErrorReportGlobal 71, "clsWorktableSeqPlate: Tried to pipette Primer to a well supposed to be empty!"
    End Select
    
End Sub

Public Sub ToggleTemplatePipetted(InRow As Integer, InColumn As Integer)

    Dim tempIndex As Integer
    
    tempIndex = GetIndexFromXY(InColumn, InRow)

    Select Case State(tempIndex)
        Case eTaken
            State(tempIndex) = eTemplatePipetted
        Case ePrimerPipetted
            State(tempIndex) = eAllPipetted
        Case Else
            ErrorReportGlobal 71, "clsWorktableSeqPlate: Tried to pipette Template DNA to a well supposed to be empty!"
    End Select
    
End Sub

Public Property Get Well(ByVal i As Integer) As clsSeqReaction
    Set Well = pSeqReactionList(i)
End Property

Public Property Set Well(ByVal i As Integer, q As clsSeqReaction)
    Set pSeqReactionList(i) = q
End Property

Public Sub DefineSite(ByVal Grid As Integer, ByVal Site As Integer)

    pGrid = Grid
    pSite = Site

End Sub

Public Sub Define(Optional Grid As Integer = 0, Optional Site As Integer = 0)
    
    Dim i As Integer
    
    pXDim = 12
    pYDim = 8
    pTotalWells = pXDim * pYDim
    ReDim pSeqReactionList(1 To pTotalWells)
    ReDim pPipettingStates(1 To pTotalWells)
    
    For i = 1 To pTotalWells
        pPipettingStates(i) = eFree
    Next i
    
    DefineSite Grid, Site
    
End Sub

Public Sub ImportRange(InputRange As Range)

    ImportArray InputRange.Value

End Sub


Private Sub ErrorReport(ErrorNumber As Integer, ErrorString As String)

    Err.Raise vbError + ErrorNumber, , ErrorString

End Sub

Private Sub WellIndexToXY(coordinate As Integer, x As Integer, y As Integer)

    x = 1 + (coordinate - 1) \ pYDim
    y = 1 + (coordinate - 1) Mod pYDim

End Sub

Private Sub XYToWellIndex(x As Integer, y As Integer, WellIndexOutput As Integer)

    WellIndexOutput = (x - 1) * pYDim + y

End Sub

Public Function GetIndexFromXY(x As Integer, y As Integer) As Integer
    XYToWellIndex x, y, GetIndexFromXY
End Function


