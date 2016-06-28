Attribute VB_Name = "XmodMetaCode"
Option Explicit
'****************************************************************************************************
Sub ConvertCodeToVBAString()
'====================================================================================================
'shows an inputbox - pasted code will be converted to a string declaration to be used in VBA
'Juraj Ahel, 2016-05-09
'Last update 2016-05-09
'====================================================================================================
'TODO: make the input form get created programatically!
'TODO: make the code keep the commented part unchanged (tabs, ...)!
'TODO: make the program recognize the number of different indentation levels, instead of assuming Tab = 4 spaces!
'      e.g. if current line has 0 spaces, and next one has 3, that's 1 tab, and if it remains 3 it's still 1 tab
'      and if it increases (e.g. to 5), it adds another tab (regardless of number of spaces), and if it decreases, then it checks
'      whether it has decreased to the same level (e.g. back to 3) it brings back 1 tab level. Otherwise, it needs to introduce
'      an intermediate tab level, if the level is e.g. 4!
'      this is not easy to do
'TODO: alternatively, simply copy the spaces instead of tabs...

    'the first input I used it for!
    
    Const DefaultInput As String = _
        "Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)" & vbCrLf & _
        vbCrLf & _
        vbTab & "Const conDefaultErrorN As Long = 1" & vbCrLf & _
        vbTab & vbCrLf & _
        vbTab & "If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN" & vbCrLf & _
        vbTab & vbCrLf & _
        vbTab & "'If Len(ErrorString) = 0 Then ErrorString = conDefaultError" & vbCrLf & _
        vbTab & vbCrLf & _
        vbTab & "If Len(ErrorString) = 0 Then" & vbCrLf & _
        vbTab & vbTab & "Err.Raise vbError + ErrorNumber, conClassName, ErrorString" & vbCrLf & _
        vbTab & "Else" & vbCrLf & _
        vbTab & vbTab & "Err.Raise vbError + ErrorNumber, conClassName" & vbCrLf & _
        vbTab & "End If" & vbCrLf & _
        vbCrLf & _
        "End Sub" & vbCrLf


    Dim InputString As String
    Dim OutputString As String
    
    'Invoke Input Form:
    InputForm.TextBox.Text = DefaultInput
    InputForm.Show
    
    'Take values from the input form:
    InputString = InputForm.TextBox.Text
        
    'Unload InputForm '- keeps giving focus to Input Form if ran, don't know how to stop that
        
    OutputString = InputString
    
    'Replace quotes with double quotes (Chr 34 = " )
    OutputString = Replace(OutputString, Chr(34), Chr(34) & Chr(34))
    
    'Replace every 4 spaces with a declared tab characters
    'This can be done in a smarter way, recursively defining levels!
    'Usually when you copy the code from VBA into here, it's pasted as spaces, not tabs
    OutputString = Replace(OutputString, VBA.Strings.Space$(4), """ & vbTab & """)
    
    'Do the same for actual tabs
    OutputString = Replace(OutputString, vbTab, """ & vbTab & """)
    
    'Replace endlines as declared Carriage Return + Line Feed ( vbCrLf )
    OutputString = Replace(OutputString, vbCrLf, """ & vbCrLf & """)
    
    'Clear excess string joins (empty strings) - multiple times, just in case
    OutputString = Replace(OutputString, "& " & Chr(34) & Chr(34) & " &", "&")
    OutputString = Replace(OutputString, "& " & Chr(34) & Chr(34) & " &", "&")
    OutputString = Replace(OutputString, "& " & Chr(34) & Chr(34) & " &", "&")
    OutputString = Replace(OutputString, "& " & Chr(34) & Chr(34) & " &", "&")
    OutputString = Replace(OutputString, "& " & Chr(34) & Chr(34) & " &", "&")
    
    ''Replace every 4 spaces with a tab
    'OutputString = Replace(OutputString, "vbTab & vbCrLf", "vbCrLf")
    
    'add first and last quote character
    OutputString = Chr(34) & OutputString & " & vbCrLf & " & Chr(34)
    
    Debug.Print OutputString
    Range("A1") = OutputString

End Sub

'****************************************************************************************************
Sub CreateClassFromPropertyList()
'====================================================================================================
'
'Juraj Ahel, 2016-02-xx
'Last update 2016-05-09
'====================================================================================================

'Prints out to cell A1 all the basic declarations of the class
'Input format is a selection of cells, 6 cols wide, any number of rows:
'["Private"/"Public"]["Object"/.*][VarType (Long, ...)][Property Name][Get Private/Public/.*][Set/Let Private/Public/.*]
'
'it declares:
'1) the internal variable as Private p[PropertyName] or Public [Propertyname], if so requested (blank creates no variable, just the properties!
'2) object constructors and deconstructors if necessary (for Class_init and terminate)
'3) Property Get and property Let/Set for each private internal variable, with scope as defined in columns 5 and 6

'TODO:
'add support to automatically set up the Cloning, based on variables marked as "y" in another column, which need to be copied over

'ErrorReport function added to each class

Const conErrorReport As String = _
        "Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)" & vbCrLf & _
        vbCrLf & _
        vbTab & "Const conDefaultErrorN As Long = 1" & vbCrLf & _
        vbTab & vbCrLf & _
        vbTab & "If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN" & vbCrLf & _
        vbTab & vbCrLf & _
        vbTab & "'If Len(ErrorString) = 0 Then ErrorString = conDefaultError" & vbCrLf & _
        vbTab & vbCrLf & _
        vbTab & "If Len(ErrorString) = 0 Then" & vbCrLf & _
        vbTab & vbTab & "Err.Raise vbError + ErrorNumber, conClassName, ErrorString" & vbCrLf & _
        vbTab & "Else" & vbCrLf & _
        vbTab & vbTab & "Err.Raise vbError + ErrorNumber, conClassName" & vbCrLf & _
        vbTab & "End If" & vbCrLf & _
        vbCrLf & _
        "End Sub" & vbCrLf
       
       
Const conInputVarPrefix = "inp_"


Dim OutputRange         As Range
Dim Inp                 As Range
Dim Inputs()            As Variant
Dim tempTextArray(1 To 13) As String
Dim tempString          As String
Dim tempStringDecl      As String
Dim tempStringInit      As String
Dim tempStringDestr     As String
Dim tempVarPrefix       As String

Dim TodayDate           As String

Dim CodeHeader          As String
Dim OutputStringArray() As String


Dim NumberOfRows As Long

Dim i As Long

Dim PrivateOrPublic     As String
Dim ObjectOrNot         As String
Dim VariableType        As String
Dim PropertyName        As String
Dim ScopeGet            As String
Dim ScopeSetLet         As String

Dim AmIDone             As Boolean
Dim HasVariable         As Boolean


'Custom header with today's date (dynamic)!

TodayDate = Format(VBA.Date, "YYYY-MM-DD")

CodeHeader = _
        "'****************************************************************************************************" & vbCrLf & _
        "'====================================================================================================" & vbCrLf & _
        "'This class does briefly this and that" & vbCrLf & _
        "'Juraj Ahel, " & TodayDate & ", for this and that purpose" & vbCrLf & _
        "'Last update " & TodayDate & vbCrLf & _
        "'====================================================================================================" & vbCrLf & _
        "Option Explicit" & vbCrLf & vbCrLf & _
        "Private Const conClassName As String = ""clsClassTemplate""" & vbCrLf


'Inputs
Set Inp = Selection
Inputs = Inp.Value

Set OutputRange = Range("A1")

NumberOfRows = Inp.Rows.Count

For i = 1 To NumberOfRows
    
    'reset array ...
    tempTextArray(1) = ""
    tempTextArray(2) = ""
    tempTextArray(3) = ""
    tempTextArray(4) = ""
    tempTextArray(5) = ""
    tempTextArray(6) = ""
    tempTextArray(7) = ""
    tempTextArray(8) = ""
    tempTextArray(9) = ""
    tempTextArray(10) = ""
    tempTextArray(11) = ""
    tempTextArray(12) = ""
    tempTextArray(13) = ""

    AmIDone = False
    
    PrivateOrPublic = Inputs(i, 1)
    ObjectOrNot = Inputs(i, 2)
    VariableType = Inputs(i, 3)
    PropertyName = Inputs(i, 4)
    ScopeGet = Inputs(i, 5)
    ScopeSetLet = Inputs(i, 6)
    
    
    'private variables get "p" as prefix
    Select Case UCase(PrivateOrPublic)
        Case "PRIVATE"
            tempVarPrefix = "p"
            AmIDone = False
            HasVariable = True
        Case "PUBLIC"
            tempVarPrefix = ""
            AmIDone = True 'variable directly accessible, no property definition
            HasVariable = True
        Case "PROTECTED"
            tempVarPrefix = "p"
            AmIDone = False
            HasVariable = True
        Case Else
            'there will be no variable assignment, the property is calculated from elsewhere
            tempVarPrefix = ""
            AmIDone = False
            HasVariable = False
            
    End Select
    
    'Declarations
    If HasVariable Then
        tempStringDecl = tempStringDecl & PrivateOrPublic & " " & tempVarPrefix & PropertyName & " As " & VariableType & vbCrLf
    
    
        'Construction / destruction of objects
        If UCase(ObjectOrNot) = "OBJECT" Then
            tempStringInit = tempStringInit & vbTab & "Set " & tempVarPrefix & PropertyName & " = New " & VariableType & vbCrLf
            tempStringDestr = tempStringDestr & vbTab & "Set " & tempVarPrefix & PropertyName & " = Nothing" & vbCrLf
        End If
    End If
    
    'In case variable is public, we don't need property gets and sets
    If Not AmIDone Then
        
        Select Case UCase(ScopeGet)
            Case "PRIVATE", "PUBLIC", "PROTECTED"
                'Property Get
                tempTextArray(1) = ScopeGet 'Private
                tempTextArray(2) = " Property "
                tempTextArray(3) = "Get "
                tempTextArray(4) = PropertyName
                tempTextArray(5) = "() "
                tempTextArray(6) = ""
                tempTextArray(7) = "as "
                tempTextArray(8) = VariableType
                tempTextArray(9) = vbCrLf & vbTab
                If HasVariable Then
                    If ObjectOrNot = "Object" Then
                        tempTextArray(10) = "Set "
                    Else
                        tempTextArray(10) = ""
                    End If
                    
                    tempTextArray(11) = PropertyName
                    tempTextArray(12) = " = " & tempVarPrefix & PropertyName
                End If
                tempTextArray(13) = vbCrLf & "End Property" & vbCrLf & vbCrLf
            
                tempString = tempString & Join(tempTextArray, "")
            Case Else
                'do nothing
        End Select
           
        Select Case UCase(ScopeSetLet)
            Case "PRIVATE", "PUBLIC", "PROTECTED"
                'Property Let/set
                tempTextArray(1) = ScopeSetLet 'Private
                tempTextArray(2) = " Property "
                If ObjectOrNot = "Object" Then
                    tempTextArray(3) = "Set "
                Else
                    tempTextArray(3) = "Let "
                End If
                tempTextArray(4) = PropertyName
                tempTextArray(5) = "("
                tempTextArray(6) = conInputVarPrefix & PropertyName
                tempTextArray(7) = " as "
                tempTextArray(8) = VariableType
                tempTextArray(9) = ")" & vbCrLf & vbTab
                If HasVariable Then
                    If ObjectOrNot = "Object" Then
                        tempTextArray(10) = "Set "
                    Else
                        tempTextArray(10) = ""
                    End If
                    tempTextArray(11) = tempVarPrefix & PropertyName
                    tempTextArray(12) = " = " & conInputVarPrefix & PropertyName
                End If
                tempTextArray(13) = vbCrLf & "End Property" & vbCrLf & vbCrLf
                
                tempString = tempString & Join(tempTextArray, "")
            Case Else
                'do nothing
        End Select
        
    End If

NextVariable:
Next i

tempStringDecl = "'[Var Declaration]" & vbCrLf & _
                                                    tempStringDecl
                                                            
tempStringInit = "'[Object Initialization]" & vbCrLf & _
                                                        "Private Sub Class_Initialize()" & vbCrLf & vbCrLf & _
                                                        tempStringInit & vbCrLf & _
                                                        "End Sub" & vbCrLf
                                                            
tempStringDestr = "'[Object Dereferencing]" & vbCrLf & _
                                                        "Private Sub Class_Terminate()" & vbCrLf & vbCrLf & _
                                                        tempStringDestr & vbCrLf & _
                                                        "End Sub" & vbCrLf
                                                                                                                        
                                                            
tempString = "'[Property Gets and Lets and Sets]" & vbCrLf & _
                                                            tempString


ReDim OutputStringArray(9 To 14)

'OutputStringArray(1) = tempStringDecl
'OutputStringArray(2) = tempStringDecl
'OutputStringArray(3) = tempStringDecl
'OutputStringArray(4) = tempStringDecl
'OutputStringArray(5) = tempStringDecl
'OutputStringArray(6) = tempStringDecl
'OutputStringArray(7) = tempStringDecl
'OutputStringArray(8) = tempStringDecl

OutputStringArray(9) = CodeHeader
OutputStringArray(10) = tempStringDecl
OutputStringArray(11) = tempStringInit
OutputStringArray(12) = tempStringDestr
OutputStringArray(13) = conErrorReport
OutputStringArray(14) = tempString
'OutputStringArray(15) = tempString


'tempString = tempStringDecl & vbCrLf & _
'                tempStringInit & vbCrLf & _
'                tempStringDestr & vbCrLf & _
'                conErrorReport & vbCrLf & _
'                tempString

' "'" character in beginning is neccessary just because we are outputing to excel cell
tempString = Join(OutputStringArray, vbCrLf)
OutputRange.Value = "'" & tempString
Debug.Print (tempString)


End Sub

