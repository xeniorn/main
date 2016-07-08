Attribute VB_Name = "modGeneral"
Dim b(1 To 12, 1 To 8) As Byte

Sub test22()

Dim aa As clsWorktableContainer
Dim b(1 To 12, 1 To 8) As Byte

Set aa = New clsWorktableContainer

aa.Define "96", 5, 1

aa.ImportRange Range("Container1")





End Sub

Sub test1()

a = StringToBitCode("0")
 b = StringToBitCode("1")
c = StringToBitCode("111")
d = StringToBitCode("1110")

k = BitFlag("1235678", 8)

End Sub



Function IsMemberOf(TestElement As Variant, SetList As Collection) As Boolean

    Dim i As Integer
    
    If SetList Is Nothing Then GoTo 999
    
    For i = 1 To SetList.Count
        If TestElement = SetList.Item(i) Then IsMemberOf = True
    Next i
    

999 End Function

Function StringToBitCode(InputString As String) As Integer
'This is actually just binary to decimal, lol
'Might write a more general function

Dim i As Integer
Dim StringLength As Integer
Dim tempResult As Integer
Dim RightFormat As Boolean

'make a check whether it's the right format...
RightFormat = True
If RightFormat Then
    
    StringLength = Len(InputString)
    
    For i = 0 To StringLength - 1
        tempResult = tempResult + 2 ^ i * (Val(Mid(InputString, StringLength - i, 1)))
    Next i

End If

StringToBitCode = tempResult

End Function

Sub ResetArray(InpArray As Variant, Optional targetValue = 0)

Dim i As Integer

For i = LBound(InpArray) To UBound(InpArray)
    InpArray(i) = targetValue
Next i

End Sub


Function BitFlag(InputString As String, flagLength As Integer) As String
'This is actually just binary to decimal, lol
'Might write a more general function

Dim i As Integer
Dim StringLength As Integer
Dim tempResult() As String
Dim RightFormat As Boolean
Dim FlagFormatString As String

ReDim tempResult(1 To flagLength)

ResetArray tempResult, "0"

'make a check whether it's the right format...
RightFormat = True
If RightFormat Then
    
    StringLength = Len(InputString)
    
    For i = 1 To StringLength
        tempResult(CInt(Mid(InputString, i, 1))) = "1"
    Next i

End If

BitFlag = Join(tempResult, "")

'For i = 1 To FlagLength
'    FlagFormatString = FlagFormatString & "0"
'Next i

'BitFlag = Format(TempResult, FlagFormatString)

End Function
