VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPipettingPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PlateXDim As Integer = 12
Private Const PlateYDim As Integer = 8

'Plate definition
Private pXDim As Integer
Private pYDim As Integer
Private pTotalNumber As Integer

'Data collections
Private pWells() As String


Private Sub Class_Initialize()

    'If pSeqReactionList() Is Nothing Then
    '    Set pPrimers = New Collection
    '    Set pPrimersKeys = New Collection
    'End If

End Sub


Private Sub Class_Terminate()

End Sub

Public Property Let Wells(Index As Integer, W As String)

    '###TODO check for out-of-bounds index
    pWells(Index) = W

End Property

Public Property Get Wells(Index As Integer) As Integer

    '###TODO check for out-of-bounds index
    Wells = pWells(Index)

End Property

Public Sub Define(x As Integer, y As Integer)
    
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


Public Function FindMaxInRow(RowNo As Integer)

Dim pos As Integer
Dim i As Integer
Dim j As Integer
Dim tempIndex As Integer

Dim RowCounts()

ReDim RowCounts(1 To pXDim)

Dim MaxPrimer As String

pos = RowNo

For i = 1 To pXDim
    RowData(i) = Wells(pos)
    pos = pos + pYDim
Next i

For i = 0 To pXDim - 1
    For j = 1 To pXDim
        tempIndex = i + j * pYDim
        RowCounts(i) = RowCounts(i) - (Wells(tempIndex) <> "") And (Wells(tempIndex) = Wells(tempIndex))
    Next j
Next i

FindMaxPrimerInRow = RowData(ArrayMaxElement(RowCounts))

End Function

'****************************************************************************************************
'====================================================================================================
'Catching all the errors.
'Juraj Ahel, 2015-11-23
'Last update 2015-11-23
'====================================================================================================

Private Sub ErrorReport(ErrorCode As Integer)
    Select Case ErrorCode
        Case 1
            R "No free tubes found in plate."
        Case 2
            R "Tried to move outside of plate."
        Case 3
            R "Tried to do a well movement in nonexisting direction."
    End Select
    
End Sub



