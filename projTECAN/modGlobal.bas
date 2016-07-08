Attribute VB_Name = "modGlobal"
Option Explicit
Option Compare Text

'****************************************************************************************************
Function TempTimeStampName() As String

'====================================================================================================
'A simple function that generates a timestamp string, containing full date and time without delimiters
'(YYYYMMDDhhmmss format)
'Juraj Ahel, 2015-02-11, for creating (almost certainly) unique files for GibsonTest
'Last update 2015-02-11
'====================================================================================================

Dim t As String

t = Now
t = Replace(t, " ", "")
t = Replace(t, ":", "")
t = Replace(t, "-", "")

TempTimeStampName = t

End Function


'****************************************************************************************************
Sub ErrorReportGlobal(ErrorNumber As Integer, ErrorString As String, Optional Source)

'====================================================================================================
'Throws an error, using the internal mechanisms of VBA IDE
'Juraj Ahel, 2015-12-xx
'Last update 2016-03-16
'====================================================================================================
    
    Dim SourceString As String
    
    Select Case VarType(Source)
        Case vbString
            SourceString = Source
        Case vbObject
            'todo later
        Case Else
            'todo later
    End Select
            
    If SourceString <> "" Then
        Err.Raise vbError + ErrorNumber, SourceString, ErrorString
    Else
        Err.Raise vbError + ErrorNumber, , ErrorString
    End If

End Sub


'****************************************************************************************************
Sub R(code)
'====================================================================================================
'Export a debug code to an external file (needed to debug in-class errors, as the VBA error handler
'catches and displays the error only once it's back to the module, apparently
'Juraj Ahel, 2015-11-01
'Juraj Ahel, 2015-11-01
'====================================================================================================
    Dim TT As String
    
    TT = TempTimeStampName
    
    ExportDataToTextFile CStr(code), "C:\ExcelExports\Debug\" & TT & "_DebugInfo.log"

End Sub

'****************************************************************************************************
Sub ExportDataToTextFile(DataToOutput As String, OutputFilename As String)
'====================================================================================================
'Wrapper for cleanly writing string data as-is to a file
'Juraj Ahel, 2014-xx-xx
'Juraj Ahel, 2014-xx-xx
'====================================================================================================

Open OutputFilename For Output As #1

Print #1, DataToOutput
Close #1

End Sub

'****************************************************************************************************
'====================================================================================================
'Catching all the errors.
'Juraj Ahel, 2015-11-23
'Last update 2016-01-07
'====================================================================================================

Sub ErrorReport(ErrorCode As Integer)

    Dim VDesc As String
    Dim VTitle As String
    

    Select Case ErrorCode
        Case 1000
            VDesc = "Tried to add wrong type to a clsTypeCollection"
            VTitle = VDesc
            R VDesc
            
        Case 1001
            VDesc = "Tried to change type of nonempty clsTypeCollection"
            VTitle = VDesc
            R VDesc
        Case 1002
            VDesc = "Too many primers of same name"
            VTitle = VDesc
            R VDesc
        Case 1003
            VDesc = "Tried to add two primers with the same name to a clsPrimers without AllowRedundancy"
            VTitle = VDesc
            R VDesc
    End Select
    
    Err.Raise vbObjectError + ErrorCode, , VTitle 'add description, etc

End Sub

' TODO: change this to the name of your module
'Private Const sMODULE As String = "MVbaUtils"

Public Function IsElementOf(sKey As String, oCollection As Collection) As Boolean
    'Const scSOURCE As String = "IsElementOf"

    Dim lErrNumber As Long
    Dim sErrDescription As String

    lErrNumber = 0
    sErrDescription = "unknown error occurred"
    Err.Clear
    On Error Resume Next
        ' note: just access the item - no need to assign it to a dummy value
        ' and this would not be so easy, because we would need different
        ' code depending on the type of object
        ' e.g.
        '   Dim vItem as Variant
        '   If VarType(oCollection.Item(sKey)) = vbObject Then
        '       Set vItem = oCollection.Item(sKey)
        '   Else
        '       vItem = oCollection.Item(sKey)
        '   End If
        oCollection.Item sKey
        lErrNumber = CLng(Err.Number)
        sErrDescription = Err.Description
    On Error GoTo 0

    If lErrNumber = 5 Then ' 5 = not in collection
        IsElementOf = False
    ElseIf (lErrNumber = 0) Then
        IsElementOf = True
    'Else
        ' Re-raise error
    '    err.Raise lErrNumber, mscMODULE & ":" & scSOURCE, sErrDescription
    End If
End Function



'****************************************************************************************************
Function ArrayMaxElement(TestArray As Variant, _
    Optional DimensionIndex As Long = 1, _
    Optional col1 As Long, _
    Optional col2 As Long, _
    Optional col3 As Long, _
    Optional col4 As Long, _
    Optional col5 As Long _
    ) As Long
'====================================================================================================
'Finds the index of the maximum value in an array, using the dimension DimensionIndex
'supports arrays up to dim5 in size
'https://support.microsoft.com/en-us/kb/152288
'Microsoft, taken 2016-01-08
'Last update 2016-01-08
'====================================================================================================

Dim FirstIndex As Long
Dim LastIndex As Long
Dim MaxValue
Dim i As Long
Dim MaxIndex As Long

FirstIndex = LBound(TestArray, DimensionIndex)
LastIndex = UBound(TestArray, DimensionIndex)

MaxIndex = FirstIndex
MaxValue = TestArray(FirstIndex)

For i = FirstIndex + 1 To LastIndex
    If TestArray(i) > MaxValue Then
        MaxValue = TestArray(i)
        MaxIndex = i
    End If
Next i

ArrayMaxElement = MaxIndex

End Function
