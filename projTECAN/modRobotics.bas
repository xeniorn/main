Attribute VB_Name = "modRobotics"
Dim welly() As Byte
'ReDim welly(1 To 12, 1 To 8)

Public Function GetWellString(ByRef Wells() As Byte) As String
    
    Dim BitCounter As Byte
    Dim BitShift As Byte
    Dim BitMask As Byte
    Dim XWells As Integer
    Dim YWells As Integer
    Dim Char As String
    Dim x As Integer
    Dim y As Integer
    
    XWells = UBound(Wells, 1)
    YWells = UBound(Wells, 2)
    
    BitMask = 0
    BitCounter = 0
    
    Char = Hex(XWells)
    If Len(Char) = 1 Then Char = "0" & Char
    GetWellString = GetWellString & Char
    
    Char = Hex(YWells)
    If Len(Char) = 1 Then Char = "0" & Char
    GetWellString = GetWellString & Char
    
    For x = 1 To XWells
        For y = 1 To YWells
        
            If Wells(x, y) And 1 Then
                BitShift = 2 ^ BitCounter
                BitMask = BitMask Or BitShift
            End If
            
            BitCounter = BitCounter + 1
            
            If BitCounter > 6 Then
                Char = Chr(Asc("0") + BitMask)
                GetWellString = GetWellString & Char
                BitCounter = 0
                BitMask = 0
            End If
        
        Next y
    Next x
    
    If BitCounter > 0 Then
        Char = Chr(Asc("0") + BitMask)
        GetWellString = GetWellString & Char
    End If

End Function

Sub resetWelly()

ReDim welly(1 To 12, 1 To 8)

For i = 1 To 12
    For j = 1 To 8
        welly(1, j) = 0
    Next j
Next i

End Sub

Sub DispenseCommand()

'Juraj Ahel, 2015-12-17
'Last Update 2015-12-17

'THIS ACTUALLY FRIGGIN' WORKS, PIPETTES WHAT I MEANT FOR IT TO PIPETTE!!!

Dim Inp As Range
Dim InpArray()
ReDim InpArray(1 To 8, 1 To 12)

'Set inp = Selection
Set Inp = Range("solo")
InpArray = Inp.Value

ReDim welly(1 To 12, 1 To 8)
Dim CommandString(1 To 17) As String

welly(5, 1) = 1
welly(5, 2) = 1
welly(5, 3) = 1
welly(5, 4) = 0
welly(5, 5) = 0
welly(5, 6) = 1
welly(5, 7) = 1
welly(5, 8) = 1

For i = 1 To 12
    resetWelly
    For j = 1 To 8
        If InpArray(j, i) > 0 Then welly(i, j) = 1
    Next j
    wellLabel = GetWellString(welly)
    CommandString(1) = 0
    CommandString(2) = """Water""" 'liquid type
    For k = 1 To 8 'fill up volume parameters
        If InpArray(k, i) > 0 Then
            CommandString(2 + k) = """" & InpArray(k, i) & """"
        Else
            CommandString(2 + k) = "0"
        End If
        CommandString(1) = CommandString(1) + welly(i, k) * 2 ^ (k - 1)
    Next k
    For k = 1 To 4
        CommandString(10 + k) = "0" 'extra zeroes because I have 8 and not 12 tips
    Next k
    CommandString(15) = "18,1,1" 'position, position within rack, spacing
    CommandString(16) = """" & wellLabel & """"
    CommandString(17) = "0" 'looping options
    
    CommandStart = "Dispense(" 'command
    CommandEnd = ");"
    CommandTotal = CommandStart & Join(CommandString, ",") & CommandEnd
    Range("A" & i) = CommandTotal
    
    
Next i


testt = GetWellString(welly)

If testt = "0C08¯1000000000000" Then MsgBox ("SUCCESS BEYBE!!!!")



End Sub
