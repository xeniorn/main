Attribute VB_Name = "modExtra"
Sub DoTheThingAdvanced()
'Sets up the automatic dispensing from a virtual plate
'And generates the actual commands
'This one should handle things in a time-optimized way
'Juraj Ahel, 2016-01-11
'Last Update 2016-01-12

Dim InpRan As Range
Dim OutRan As Range
Dim tempRan As Range
Dim InpPlate0(), InpPlate()
Dim tempMax() As String
Dim tempCount() As Integer
Dim rcount As Integer, ccount As Integer
Dim i As Integer, j As Integer

Set InpRan = Selection
Set OutRan = InpRan.Offset(12, 0)
InpPlate0 = Selection.Value
InpPlate = InpPlate0

rcount = InpRan.Rows.Count
ccount = InpRan.Columns.Count

ReDim tempMax(1 To rcount)
ReDim tempCount(1 To rcount)

For i = 1 To rcount
    tempMax(i) = ""
    tempCount(i) = 0
Next i

Set tempRan = Nothing

For i = 1 To rcount
    Set tempRan = InpRan.Offset(i - 1, 0).Resize(1, ccount)
    tempMax(i) = MaxMatrixInRow(tempRan)
    tempCount(i) = MatrixElementCount(tempMax(i), InpPlate, i, 0)
    MatrixElementReplace InpPlate, tempMax(i), "", i, 0
Next i
    
For i = 1 To rcount
    tempCount(i) = tempCount(i) + MatrixElementCount(tempMax(i), InpPlate)
    MatrixElementReplace InpPlate, tempMax(i), ""
Next i

'1 count the max element in each row
For i = 1 To rcount
    tempMax(i) = MatrixMaxElement(InpPlate, i, 0)
    tempCount(i) = MatrixElementCount(tempMax(i), InpPlate, i, 0)
Next i

Dim MaxRedundancy As Integer

'2 check if max in two rows are the same
MaxRedundancy = MatrixMaxCount(tempMax)

If MaxRedundacy <> 0 Then
'2a if yes, and there are not many (maybe require at least 6, or only if it's a full row
'   then make the most populous row "eat" the smaller row (even if 2 pipette tips are
'   faster than 1, washing the tips is the highest-cost step - maybe once I can write
'   an optimization algorithm that calculates the optimal procedure, but not now. This
'   also minimizes chances for cross-contamination) and recalculate the smaller row and
'   all the subsequent rows as well, as the smaller row might now eat yet another row
'2b if yes, and there are really many in both, split the loot between the 2 biggest master rows
'   this shouldn't happen very often, in principle, as it would mean there are many many samples
'   with the same primer, which is something rather easy to handle in general, but for the sake
'   of completeness, I would also include this, at least as a TODO


End If

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

OutRan.Value = InpPlate

'###OUTRAN is output and forwarded on to the next instance of the iterator for recursive solution


End Sub

