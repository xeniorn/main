Attribute VB_Name = "XmodMatrix"


'****************************************************************************************************
Function MatrixDimesionNumber(XArray As Variant)
'====================================================================================================
'Finds the number of dimensions of an array
'https://support.microsoft.com/en-us/kb/152288
'Microsoft, taken 2016-01-08
'Juraj Ahel
'Last update 2016-03-16
'====================================================================================================
      
    Dim DimNum As Long
    
    'Sets up the error handler.
    On Error GoTo FinalDimension
    
    'Visual Basic for Applications arrays can have up to 60000
    'dimensions; this allows for that.
    For DimNum = 1 To 60000
    
       'It is necessary to do something with the LBound to force it
       'to generate an error.
       ErrorCheck = LBound(XArray, DimNum)
    
    Next DimNum
    
    DimNum = DimNum + 1
    
    ' The error routine.
FinalDimension:

    On Error GoTo 0
    
    MatrixDimesionNumber = DimNum - 1

   End Function
   
'****************************************************************************************************
Function PrintMatrixXY(InputMatrix As Variant, Optional MaxDimension = 0) As String
'====================================================================================================
'Prints out a 2D matrix (as strings) so that the first dimension is printed in rows, and second in columns!
'Juraj Ahel, 2016-03-15
'Last update 2016-03-16
'====================================================================================================


    Const conSep As String = ", "
    Const conSep2 As String = vbCrLf
    
    Dim S1, S2, e1, e2, i, j
    Dim temprowstring()
    Dim tempcolumnstring()
    
    If MaxDimension = 0 Then MaxDimension = 100
    
    S1 = LBound(InputMatrix, 1)
    S2 = LBound(InputMatrix, 2)
    
    e1 = UBound(InputMatrix, 1)
    e2 = UBound(InputMatrix, 2)
    
    ReDim temprowstring(S1 To e1)
    ReDim tempcolumnstring(S2 To e2)
    
    For j = S2 To e2
        For i = S1 To e1
            temprowstring(i) = CStr(InputMatrix(i, j))
        Next i
        tempcolumnstring(j) = Join(temprowstring, conSep)
    Next j
    
    PrintMatrixXY = Join(tempcolumnstring, conSep2)

End Function
'****************************************************************************************************
Function PrintMatrixYX(InputMatrix As Variant, Optional MaxDimension = 0) As String
'====================================================================================================
'Prints out a 2D matrix (as strings) so that the first dimension is printed in columns, and second in rows!
'Juraj Ahel, 2016-03-15
'Last update 2016-03-16
'====================================================================================================


    Const conSep As String = ", "
    Const conSep2 As String = vbCrLf
    
    Dim S1, S2, e1, e2, i, j
    Dim temprowstring()
    Dim tempcolumnstring()
    
    If MaxDimension = 0 Then MaxDimension = 100
    
    S1 = LBound(InputMatrix, 1)
    S2 = LBound(InputMatrix, 2)
    
    e1 = UBound(InputMatrix, 1)
    e2 = UBound(InputMatrix, 2)
    
    ReDim temprowstring(S2 To e2)
    ReDim tempcolumnstring(S1 To e1)
    
    For i = S1 To e1
        For j = S2 To e2
            temprowstring(j) = CStr(InputMatrix(i, j))
        Next j
        tempcolumnstring(i) = Join(temprowstring, conSep)
    Next i
    
    PrintMatrixYX = Join(tempcolumnstring, conSep2)

End Function


Function MaxMatrixInRow(InpRange As Range)

Dim Arrayos(1 To 12) As Long
Dim InpMatrix()
InpMatrix = InpRange.Value
Dim i As Long, j As Long, k As Long

For k = 1 To InpRange.Rows.Count
    For i = 1 To InpRange.Columns.Count
        If InpMatrix(1, i) <> "" Then
            Arrayos(i) = MatrixElementCount(InpMatrix(1, i), InpMatrix)
        End If
    Next i
Next k

j = 1

For i = 2 To InpRange.Columns.Count
    If Arrayos(i) > Arrayos(j) Then j = i
Next i

MaxMatrixInRow = InpMatrix(1, j)
If MaxMatrixInRow = "" Then MaxMatrixInRow = "EMPTY ROW"

End Function


Function MatrixMaxElement(Matrix As Variant, _
                            Optional OnlyRowN As Long = 0, _
                            Optional OnlyColumnN As Long = 0, _
                            Optional IgnoreList As Collection = Nothing, _
                            Optional IncludeOnlyList As Collection = Nothing _
                            ) As Variant
                            
                            
    Dim S1 As Long, S2 As Long, e1 As Long, e2 As Long
    Dim MaxEl As Variant
    Dim tempcounter As Long, maxcounter As Long
    
    S1 = LBound(Matrix, 1)
    S2 = LBound(Matrix, 2)
    e1 = UBound(Matrix, 1)
    e2 = UBound(Matrix, 2)
    
    If OnlyRowN > 0 Then
        S1 = OnlyRowN
        e1 = OnlyRowN
    End If
    
    If OnlyColumnN > 0 Then
        S2 = OnlyColumnN
        e2 = OnlyColumnN
    End If
    
    maxcounter = 0
    tempcounter = 0
    MaxEl = Matrix(S1, S2)
    
    For i = S1 To e1
        For j = S2 To e2
            tempcounter = MatrixElementCount(Matrix(i, j), Matrix, OnlyRowN, OnlyColumnN, IgnoreList)
            If tempcounter > maxcounter Then
                MaxEl = Matrix(i, j)
                maxcounter = tempcounter
            End If
        Next j
    Next i
    
    If maxcounter > 0 Then
        MatrixMaxElement = MaxEl
    Else
        MatrixMaxElement = Empty
    End If

End Function

Function MatrixMaxCount(Matrix As Variant, _
                            Optional OnlyRowN As Long = 0, _
                            Optional OnlyColumnN As Long = 0, _
                            Optional IgnoreList As Collection = Nothing, _
                            Optional IncludeOnlyList As Collection = Nothing _
                            ) As Variant

    Dim S1 As Long, S2 As Long, e1 As Long, e2 As Long
    Dim MaxEl As Variant
    
    S1 = LBound(Matrix, 1)
    S2 = LBound(Matrix, 2)
    e1 = UBound(Matrix, 1)
    e2 = UBound(Matrix, 2)
    
    If OnlyRowN > 0 Then
        S1 = OnlyRowN
        e1 = OnlyRowN
    End If
    
    If OnlyColumnN > 0 Then
        S2 = OnlyColumnN
        e2 = OnlyColumnN
    End If
    
    MaxEl = MatrixMaxElement(Matrix, OnlyRowN, OnlyColumnN, IgnoreList, IncludeOnlyList)
    MatrixMaxCount = MatrixElementCount(MaxEl, Matrix, OnlyRowN, OnlyColumnN, IgnoreList)
    
End Function

'****************************************************************************************************
Function MatrixElementCount(ByVal Element As Variant, _
                            ByVal Matrix As Variant, _
                            Optional ByVal OnlyRowN As Long = 0, _
                            Optional ByVal OnlyColumnN As Long = 0, _
                            Optional ByVal IgnoreList As Collection = Nothing, _
                            Optional ByVal CountAllExceptIgnored As Boolean = False, _
                            Optional ByVal Recursion As Long = -1 _
                            ) As Variant
                            
'====================================================================================================
'
'Juraj Ahel, 2016-01-11
'Last update 2016-03-16
'====================================================================================================
'default ignore and include lists include all!
'this variant allows infinitely nested arrays and collections as input for element by default
'Recursive -1 = infinite recursions; 0 = no recursions; 1 = single-level recursion
                
    Const Debugging As Boolean = False

    Dim S1 As Long, S2 As Long, e1 As Long, e2 As Long
    Dim counter As Long
    Dim ShouldICount As Boolean
    
    Dim tempType As Long
    Dim st As Long, et As Long
    
    
    '[Parsing input parameters]
    
    S1 = LBound(Matrix, 1)
    S2 = LBound(Matrix, 2)
    e1 = UBound(Matrix, 1)
    e2 = UBound(Matrix, 2)
    
    If OnlyRowN > 0 Then
        S1 = OnlyRowN
        e1 = OnlyRowN
    End If
    
    If OnlyColumnN > 0 Then
        S2 = OnlyColumnN
        e2 = OnlyColumnN
    End If
    
    counter = 0
      
    Select Case Recursion
        Case -1, 0
            'leave it, either allowed infinitely or not allowed
        Case Is > 1
            'a limited number is allowed, reduced in every pass
            Recursion = Recursion - 1
        Case Else
            'uknown parameter
            ErrorReportGlobal 5075, "Unrecognized recursion type parameter (must be Long -1/0/+x)!", "modMatrix:MatrixElementCount"
    End Select
                    
    '[Parsing main inputs]
    'check if element is a simple object or a collection thereof, and recursively solve if recursion is allowed
    Select Case VarType(Element)
        
        'Simple collections are allowed #TODO add support for my own clsTypeCollection (type-regulated collection)
        Case vbObject
            
            If TypeOf Element Is Collection Then
                
                If Recursion = 0 Then
                    ErrorReportGlobal 5076, "Element is a collection, but recursion is deeper than allowed!", "modMatrix:MatrixElementCount"
                End If
                                
                For i = 1 To Element.Count
                    counter = counter + MatrixElementCount(Element.Item(i), Matrix, OnlyRowN, OnlyColumnN, IgnoreList, CountAllExceptIgnored, Recursion)
                Next i
                
            Else
            
                ErrorReportGlobal 5077, "Element to be counted is not a simple object or a collection/array!", "modMatrix:MatrixElementCount"
                
            End If
            
        'Arrays are allowed - array type is defined by vbArray + vb[ArrayVarType], so VarType for an array is >vbArray
        Case Is > vbArray
        
            If Recursion = 0 Then
                ErrorReportGlobal 5078, "Element is a collection, but recursion is deeper than allowed!", "modMatrix:MatrixElementCount"
            End If
            
            If MatrixDimesionNumber(Element) <> 1 Then
                ErrorReportGlobal 5079, "Element is an Array, but not 1D!", "modMatrix:MatrixElementCount"
            End If
            
            For i = LBound(Element) To UBound(Element)
                counter = counter + MatrixElementCount(Element(i), Matrix, OnlyRowN, OnlyColumnN, IgnoreList, CountAllExceptIgnored, Recursion)
            Next i
            
            
        Case Else  'simple type: vbBoolean, vbByte, vbLong, vbLong, vbSingle, vbDouble,vbString, vbDate, vbCurrency, vbDecimal
            'proceed to counting
    End Select
        
    '[Counting] 'True is -1 as an Long, therefore multiplication of 3 true values is -1,
                'so it must be "substracted" from counter to increase it
    For i = S1 To e1
        For j = S2 To e2
    
            ShouldICount = ((Element = Matrix(i, j)) Or CountAllExceptIgnored) And Not (IsElementOf(CStr(Element), IgnoreList))
            counter = counter + ShouldICount

        Next j
    Next i
    
    MatrixElementCount = counter

End Function

Sub MatrixElementReplace(Element As Variant, _
                            Matrix As Variant, _
                            Optional Replacement As Variant = Empty, _
                            Optional OnlyRowN As Long = 0, _
                            Optional OnlyColumnN As Long = 0, _
                            Optional IgnoreList As Collection = Nothing, _
                            Optional IncludeOnlyList As Collection = Nothing _
                            )

    Dim S1 As Long, S2 As Long, e1 As Long, e2 As Long
    
    S1 = LBound(Matrix, 1)
    S2 = LBound(Matrix, 2)
    e1 = UBound(Matrix, 1)
    e2 = UBound(Matrix, 2)
    
    If OnlyRowN > 0 Then
        S1 = OnlyRowN
        e1 = OnlyRowN
    End If
    
    If OnlyColumnN > 0 Then
        S2 = OnlyColumnN
        e2 = OnlyColumnN
    End If
    
    
    For i = S1 To e1
        For j = S2 To e2
            If Matrix(i, j) = Element Then Matrix(i, j) = Replacement
        Next j
    Next i
    

End Sub

Function MatrixSum(Matrix As Variant, Optional OnlyRowN As Long = 0, Optional OnlyColumnN As Long = 0) As Variant

    Dim S1 As Long, S2 As Long, e1 As Long, e2 As Long
    Dim tempsum
    
    S1 = LBound(Matrix, 1)
    S2 = LBound(Matrix, 2)
    e1 = UBound(Matrix, 1)
    e2 = UBound(Matrix, 2)
    
    If OnlyRowN > 0 Then
        S1 = OnlyRowN
        e1 = OnlyRowN
    End If
    
    If OnlyColumnN > 0 Then
        S2 = OnlyColumnN
        e2 = OnlyColumnN
    End If
    
    tempsum = 0
    
    For i = S1 To e1
        For j = S2 To e2
            tempsum = tempsum + Matrix(i, j)
        Next j
    Next i
    
    MatrixSum = tempsum

End Function
