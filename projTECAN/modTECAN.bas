Attribute VB_Name = "modTECAN"
Option Explicit
Option Compare Text


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
    
    WriteTextFile CStr(code), "C:\ExcelExports\Debug\" & TT & "_DebugInfo.log"

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
