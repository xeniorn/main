VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPipettingCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const conQ As String = """"

'[Var Declaration]
Private pCommandType As String
Private ptipMask As String
Private pliquidClass As String
Private pVolumesString As String
Private pGrid As Integer
Private pSite As Integer
Private pSpacing As Integer
Private pWellSelectionString As String

'[Object Initialization]
'Set pWorktable = New clsWorktableSetup

'[Object Dereferencing]
'Set pWorktable = Nothing

'[Property Gets and Lets and Sets]
Public Property Let CommandType(InpVarCommandType As String)
    pCommandType = InpVarCommandType
End Property

Public Property Let TipMask(InpVartipMask As String)
    ptipMask = ConvertFlagToValue(InpVartipMask)
End Property

Public Property Let liquidClass(InpVarliquidClass As String)
    pliquidClass = InpVarliquidClass
End Property

Public Property Let VolumesString(InpVarVolumesString As String)
    pVolumesString = InpVarVolumesString
End Property

Public Property Let Grid(InpVarGrid As Integer)
    pGrid = InpVarGrid
End Property

Public Property Let Site(InpVarSite As Integer)
    pSite = InpVarSite
End Property

Public Property Let Spacing(InpVarSpacing As Integer)
    pSpacing = InpVarSpacing
End Property

Public Property Let WellSelectionString(InpVarWellSelectionString As String)
    pWellSelectionString = InpVarWellSelectionString
End Property

Public Property Get CommandType() As String
    CommandType = pCommandType
End Property

Public Property Get TipMask() As String
    TipMask = ptipMask
End Property

Public Property Get liquidClass() As String
    liquidClass = pliquidClass
End Property

Public Property Get VolumesString() As String
    VolumesString = pVolumesString
End Property

Public Property Get Grid() As Integer
    Grid = pGrid
End Property

Public Property Get Site() As Integer
    Site = pSite
End Property

Public Property Get Spacing() As Integer
    Spacing = pSpacing
End Property

Public Property Get WellSelectionString() As String
    WellSelectionString = pWellSelectionString
End Property

Public Sub Define(iCommandType As String, _
                    itipMask As String, _
                    iliquidClass As String, _
                    iVolumesString As String, _
                    iGrid As Integer, _
                    iSite As Integer, _
                    iSpacing As Integer, _
                    iWellSelectionString As String)
    

CommandType = iCommandType
TipMask = itipMask
liquidClass = iliquidClass
VolumesString = iVolumesString
Grid = iGrid
Site = iSite
Spacing = iSpacing
WellSelectionString = iWellSelectionString

End Sub

Private Function ConvertFlagToValue(BinaryString As String) As String

    Dim i As Integer
    Dim tempValue As Integer
    Dim flagLength As Integer
    
    flagLength = Len(BinaryString)
    
    For i = 1 To flagLength
        tempValue = tempValue + Mid(BinaryString, i, 1) * 2 ^ (i - 1)
    Next i
        
    ConvertFlagToValue = CStr(tempValue)

End Function

Public Property Get FullCommandText() As String
    
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
            ErrorReportGlobal 72, "Invalid pipetting command type parameter!", "clsPipettingCommand"
    End Select
    
    Parameters(1) = TipMask
    Parameters(2) = conQ & liquidClass & conQ
    Parameters(3) = VolumesString
    Parameters(4) = CStr(Grid)
    Parameters(5) = CStr(Site)
    Parameters(6) = CStr(Spacing)
    Parameters(7) = conQ & WellSelectionString & conQ
    Parameters(8) = CStr(connoOfLoopOptions)
    
    FullCommandText = Parameter0 & "(" & Join(Parameters, conSeparator) & ");"

End Property


