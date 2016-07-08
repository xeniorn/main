Attribute VB_Name = "modTesting"
Sub testD1()

Dim d, e

Dim a As IContainerSample
Dim b As IContainerSample

Dim a1 As clsTest
Dim b1 As clsTest1

Dim c As clsTestComparer

Set a = New clsTest
Set b1 = New clsTest1

Set c = New clsTestComparer




Set a1 = a
Set b = b1

f = b.Name

e = a1.IAmTest

d = c.Compare(b1, a1)



End Sub


Sub testABC1()

Dim a, b, c, d, e, f, g

a = " " & ConvertFlagToValue("10000000")
b = " " & ConvertFlagToValue("00000011")
c = " " & ConvertFlagToValue("00011000")
d = " " & ConvertFlagToValue("00001000")
e = " " & ConvertFlagToValue("10010001")
f = " " & ConvertFlagToValue("11111111")

MsgBox a & b & c & d & e & f



End Sub

Sub testABCD()

Dim a

a = PrintMatrixXY(Range("testo1").Value)

MsgBox a

a = PrintMatrixYX(Range("testo1").Value)

MsgBox a
End Sub

Sub testabc()

Dim ba As Boolean, bb As Boolean, bc As Boolean
Dim ia As Integer, ib As Integer
Dim bt1 As Boolean, bt2 As Boolean, bt3 As Boolean

ba = True
bb = True
bc = True

bt1 = ba * bb * bc
bt2 = ba And bb And bc


ia = ba * bb * bc
ib = bt2
ic = CInt(ba) * CInt(bb) * CInt(bc)



End Sub

Sub test11()

Dim workt As clsWorktableSetup
Dim plat As clsWorktableSeqPlate

Set plat = New clsWorktableSeqPlate
Set workt = New clsWorktableSetup

workt.FindFreeLocationFor plat


End Sub

Sub test4()

Dim arr()
Dim i As Integer
Dim j As Integer
Dim a As String
Dim b As String

arr = Array(1, 2, 5, 5, 4, 3, 2)

i = ArrayMaxElement(arr)

End Sub

Sub test3()

Dim t As New clsDNA
Dim ssam As New clsSeqSamples
Dim ssam1 As New clsSeqSample
Dim ssr As New clsSeqReaction
Dim dna1 As New clsDNA
Dim primer1 As New clsPrimer
Dim prs As New clsPrimers




t.Define Name:="plas1", Sequence:="ATG", Circular:=True


End Sub

Sub test1()
Attribute test1.VB_ProcData.VB_Invoke_Func = "Q\n14"
 
R 1
 
Dim plate As clsVirtualPlate
Set plate = New clsVirtualPlate
 
R 2
 
plate.Define 8, 12
 
Dim temppri As clsPrimer
Dim collec As clsPrimers
Set collec = New clsPrimers

R 3

For i = 1 To 13
    Set temppri = New clsPrimer
    temppri.Name = i & RandomName
    collec.AddPrimer temppri
Next i
    
   R 4
    plate.FillWithPrimerList collec
    Set collec = Nothing
    Set plate = Nothing

End Sub

Function RandomName() As String
    'Randomize
    Dim maxi As Integer, i As Integer
    maxi = 3 + Rnd() * 10
    
    
    For i = 1 To maxi
        'Randomize
        chari = 65 + Rnd() * 25
        Name = Name & Chr(chari)
    Next i
    
    RandomName = Name
    
End Function
Sub TEST2()

Dim plate As clsVirtualSequencingPlate
Set plate = New clsVirtualSequencingPlate

plate.ImportSequencingList


Dim sara(1 To 8)
Dim i As Integer

For i = 1 To 8
    sara(i) = plate.FindMaxPrimerInRow(i)
Next i

End Sub
