Attribute VB_Name = "XMain"
'Juraj Ahel, 2017-01-03
'module containing stuff that's implemented everywhere

'###depends on:
'
'

Option Explicit

Public Const JA_InteractiveTesting As Boolean = False
Public Const Debugging As Boolean = True

Public Const jaErr As Long = 10000

Public Type ErrorStructure
    Description As String
    HelpContext As Long
    HelpFile As String
    LastDllError As Long
    Number As Long
    Source As String
End Type

'************************************************************************************************
Sub ErrReraise()
'===============================================================================
'Re-raises existing error
'Juraj Ahel, 2017-01-03
'===============================================================================

    Call Err.Raise(Err.Number)

End Sub

'************************************************************************************************
Function ApplyNewError( _
    Optional Number As Long, _
    Optional Source As String, _
    Optional Description As String, _
    Optional HelpFile As String, _
    Optional HelpContext As Long _
    ) As ErrorStructure
    
'===============================================================================
'
'Juraj Ahel, 2017-01-03
'===============================================================================
        
    Dim DefErr As ErrorStructure
    
    With DefErr
        .Description = Description
        .Source = Source
        .Number = Number
        .HelpFile = HelpFile
        .HelpContext = HelpContext
    End With
    
    Call ApplyError(DefErr)

End Function
    
'===============================================================================
'Copies error details into a structure
'Juraj Ahel, 2017-01-03
'===============================================================================



'************************************************************************************************
Sub ApplyError(ES As ErrorStructure)
    
'===============================================================================
'Copies error details into a structure
'Juraj Ahel, 2017-01-03
'===============================================================================
    
    With Err
        .Description = ES.Description
        .HelpContext = ES.HelpContext
        .HelpFile = ES.HelpFile
        '.LastDllESor = ES.LastDllESor
        .Number = ES.Number
        .Source = ES.Source
    End With
    
End Sub

'************************************************************************************************
Function DefineError( _
    Optional Number As Long, _
    Optional Source As String, _
    Optional Description As String, _
    Optional HelpFile As String, _
    Optional HelpContext As Long _
    ) As ErrorStructure
    
'===============================================================================
'Initializes an ErrorStructure
'Juraj Ahel, 2017-01-03
'===============================================================================
    
    With DefineError
        .Description = Description
        .Source = Source
        .Number = Number
        .HelpFile = HelpFile
        .HelpContext = HelpContext
    End With

End Function
    
    

'************************************************************************************************
Function CopyError() As ErrorStructure
        
'===============================================================================
'Copies error details into a structure
'Juraj Ahel, 2017-01-03
'===============================================================================
    
    With CopyError
        .Description = Err.Description
        .HelpContext = Err.HelpContext
        .HelpFile = Err.HelpFile
        '.LastDllError = Err.LastDllError
        .Number = Err.Number
        .Source = Err.Source
    End With
        
End Function


Sub Test()

    Dim a As ErrorStructure
    
    a = DefineError(jaErr, "test", "aaa")
    
    'On Error Resume Next
    
    ApplyError a
    
    Err.Raise jaErr + 20

End Sub
