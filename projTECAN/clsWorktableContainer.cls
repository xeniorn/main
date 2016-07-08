VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorktableContainer"
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

Private pWells() As clsSample

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    
    Dim i As Integer
    
    For i = LBound(pWells) To UBound(pWells)
        Set pWells(i) = Nothing
    Next i

End Sub

Public Property Get RowsCount() As Integer
    RowsCount = pYDim
End Property

Public Property Get ColumnsCount() As Integer
    ColumnsCount = pXDim
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

Public Property Get Well(i As Integer) As clsSample
    Set Well = pWells(i)
End Property

Public Property Set Well(i As Integer, q As clsSample)
    Set pWells(i) = q
End Property

Public Sub Define(PlateType As String, Grid As Integer, Site As Integer)
    
    Dim i As Integer
    
    pGrid = Grid
    pSite = Site
    
    Select Case UCase(PlateType)
        Case "96", "96WP", "DEF", "DEFAULT", "1", ""
            pXDim = 12
            pYDim = 8
            pTotalWells = pXDim * pYDim
        Case Else
            'Throw Error
    End Select
    
    ReDim pWells(1 To TotalNumber)
    
    For i = 1 To TotalNumber
        Set Well(i) = New clsSample
    Next i
    
End Sub

Public Sub ImportRange(InputRange As Range)

    ImportArray InputRange.Value

End Sub

Public Sub ImportArray(SampleArray As Variant)

'imports the sample names from, say, a table in excel

    Dim i As Integer
    Dim j As Integer
    
    Dim tempInput As String
    Dim tempIndex As Integer
    
    ' Do checks...
    
    For i = LBound(SampleArray, 1) To UBound(SampleArray, 1)
        For j = LBound(SampleArray, 2) To UBound(SampleArray, 2)
            
            tempInput = CStr(SampleArray(i, j))
            If tempInput <> "" Then
                XYToWellIndex i, j, tempIndex
                Well(tempIndex).Name = tempInput
                
            End If
            
        Next j
    Next i
                    
            

End Sub

Public Function LocateSample(SoughtSample As clsSample) As clsWorktableSampleLocation

    Dim i As Integer
    Dim tempLocation As Integer
    Dim tempFound As Boolean
    Dim tempcLocationObj As clsWorktableSampleLocation
    
    Dim ErrorVal As Integer
    Dim ErrorString As String
    
    ErrorVal = 0
    ErrorString = ""
    
    tempFound = False
    tempLocation = 0
    
    For i = 1 To TotalNumber
    
        If SoughtSample.Name = Well(i).Name Then
        
            If Not tempFound Then
                tempFound = True
                Set tempcLocationObj = New clsWorktableSampleLocation
            Else
                ErrorVal = 1
                GoTo ErrorHandling
            End If
            
            tempLocation = i
        
        End If
        
    Next i
    
ErrorHandling:
    
    If ErrorVal <> 0 Then
        Select Case ErrorVal
            Case 1
                ErrorString = "Sample is not Unique"
            Case Else
                ErrorString = "Generic Error"
        End Select
        
        ErrorReport ErrorVal, ErrorString
        
    Else
    
        If Not (tempcLocationObj Is Nothing) Then
    
            Set tempcLocationObj.Container = Me
            tempcLocationObj.WellCoordinate = tempLocation
            
        End If
        
        Set LocateSample = tempcLocationObj
        
    End If
        
Set tempcLocationObj = Nothing

End Function

Private Sub ErrorReport(ErrorNumber As Integer, ErrorString As String)

    Err.Raise vbError + ErrorNumber, , ErrorString

End Sub

Public Function test1(a)

test1 = GetWellString(a)

End Function

Private Sub WellIndexToXY(Coordinate As Integer, x As Integer, y As Integer)

    x = 1 + (Coordinate - 1) \ pYDim
    y = 1 + (Coordinate - 1) Mod pYDim

End Sub

Private Sub XYToWellIndex(x As Integer, y As Integer, WellIndexOutput As Integer)

    WellIndexOutput = (y - 1) * pYDim + x

End Sub

Public Function DefineWellSelectionString(InputParameter As Variant) As String

    Dim i As Integer
    Dim j As Integer
    
    Dim x As Integer
    Dim y As Integer
    
    Dim tempcSample As clsSample
    Dim tempInput As Object
    
    Dim InputType As VbVarType
    Dim ParsedInput As Collection
    
    
    Dim Plate2D() As Byte
    
    Dim ErrorVar As Integer
    Dim ErrorMsg As String
    
    InputType = VarType(InputParameter)
    
    'parse the inputs, to allow arrays of samples, single samples, or collections thereof
    Select Case InputType
    
        Case vbObject 'object type, like a class
        
            If TypeOf InputParameter Is clsSample Then 'single sample
            
                Set ParsedInput = New Collection
                ParsedInput.Add InputParameter
                
            ElseIf TypeOf InputParameter Is Collection Then 'collection of samples
            
                For i = 1 To InputParameter.Count
                    If Not (TypeOf InputParameter.Item(i) Is clsSample) Then
                        ErrorVar = 1 'Throw error, it is not a valid input
                        GoTo ErrorHandling
                    End If
                Next i
                
                Set ParsedInput = InputParameter 'in case there was no error
                
            Else 'the parameter is neither a sample nor a collection of them
            
                ErrorVar = 1 'Throw error, it is not a valid input
                GoTo ErrorHandling
                
            End If
            
        Case vbArray + vbObject 'Array of objects
            
            Set ParsedInput = New Collection
            
            For i = LBound(InputParameter) To UBound(InputParameter)
                
                Set tempInput = InputParameter(i)
                
                If Not (TypeOf tempInput Is clsSample) Then
                    ErrorVar = 1 'Throw error, it is not a valid input
                    GoTo ErrorHandling
                Else
                    ParsedInput.Add tempInput
                End If
            
            Next i
        
        Case Else
        
            ErrorVar = 1 'Throw error, it is not a valid input
            GoTo ErrorHandling
        
    End Select
                
    'Define the byte plate needed for generation of the string by the "official function"
    
    ReDim Plate2D(1 To pXDim, 1 To pYDim)
    
    For i = 1 To pXDim
        For j = 1 To pYDim
            Plate2D(i, j) = 0
        Next j
    Next i
    
    'Flip the bits in the array where there is sample to be pipetted
    For i = 1 To ParsedInput.Count
        
        Set tempcSample = ParsedInput.Item(i)
            
        WellIndexToXY tempcSample.Location.WellCoordinate, x, y
        Plate2D(x, y) = 1
        
    Next i
            
ErrorHandling:
    
    If ErrorVar <> 0 Then
    
        Select Case ErrorVar
            Case 1
                ErrorMsg = "Invalid Input"
        End Select
    
    End If
                    
    Set tempInput = Nothing
    Set tempcSample = Nothing
            
    DefineWellSelectionString = GetWellString(Plate2D)

End Function

Private Function GetWellString(ByRef Wells As Variant) As String
    
    'took it from the Gemini Worklist Command Set example code directly
    'would be good to rewrite it to be more robust (this one has implicit type narrowing)
    'and more me-like
    
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
        
            If CByte(Wells(x, y)) And 1 Then
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

