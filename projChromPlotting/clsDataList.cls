VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'this class holds a single array, but of a defined yet variable type!
'type is stored in a separate variable (pVariableType) and the actual array is stored in the
'pCollectionObject
'
'Juraj Ahel, 2016-06-08, for structure unpacking
'Last update 2016-06-09
'====================================================================================================
Option Explicit

Private Const conClassName As String = "clsDataList"

Private pVariableType As VbVarType
Private pArray As Variant

Private Sub Class_Initialize()

    pVariableType = vbEmpty

End Sub

Private Sub Class_Terminate()
        
End Sub

Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub

Public Sub SetType(TargetType As VbVarType)
    
    'TODO make a proper check, whether the array already exists, etc
    If pVariableType = vbEmpty Then
        pVariableType = TargetType
    End If

End Sub

Public Sub SetSize(TargetSize As Long)

    ReDim pArray(1 To TargetSize)

End Sub

Public Property Let Data(DataIndex As Long, DataElement As Variant)

    Select Case VariableType
        Case vbSingle: pArray(DataIndex) = CSng(DataElement)
        Case vbDouble: pArray(DataIndex) = CDbl(DataElement)
        Case vbInteger: pArray(DataIndex) = CInt(DataElement)
        Case vbLong: pArray(DataIndex) = CLng(DataElement)
        Case vbByte: pArray(DataIndex) = CByte(DataElement)
        Case vbBoolean: pArray(DataIndex) = CBool(DataElement)
        Case vbString: pArray(DataIndex) = CStr(DataElement)
        Case Else: ErrorReport
    End Select

End Property

Public Property Get VariableType() As VbVarType
    VariableType = pVariableType
End Property

Public Property Get DataArray() As Variant

    DataArray = pArray

End Property

Public Property Let DataArray(InputArray As Variant)
    
    Dim ArrayDouble() As Double
    Dim ArrayLong() As Long
    Dim ArrayByte() As Byte
    Dim ArrayString() As String
    Dim ArrayInteger() As Integer
    Dim ArraySingle() As Single
            
    pVariableType = VarType(InputArray) - vbArray
            
    'If (VarType(InputArray) - vbArray) <> VariableType Then
    '    Call ErrorReport(, "Input array has a wrong vartype")
    'End If
    
    Select Case (VarType(InputArray) - vbArray)
        Case vbDouble
            Call ArrayCopy(InputArray, ArrayDouble)
            pArray = ArrayDouble
        Case vbLong
            Call ArrayCopy(InputArray, ArrayLong)
            pArray = ArrayLong
        Case vbByte
            Call ArrayCopy(InputArray, ArrayByte)
            pArray = ArrayByte
        Case vbString
            Call ArrayCopy(InputArray, ArrayString)
            pArray = ArrayString
        Case vbInteger
            Call ArrayCopy(InputArray, ArrayInteger)
            pArray = ArrayInteger
        Case vbSingle
            Call ArrayCopy(InputArray, ArraySingle)
            pArray = ArraySingle
        Case Else
            ErrorReport , "vartype not supported yet"
    End Select
        
    
    
End Property

