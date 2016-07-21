Attribute VB_Name = "modResearchDatabase"
Option Explicit

Const conClassName As String = "modResearchDatabase"

Private Const DatabasePath As String = "E:\PhD\ExperimentsDatabase"
Private Const DatabaseName As String = "JA_DATABASE.jadb"

Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    'If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    If Len(ErrorString) = 0 Then
        Err.Raise vbError + ErrorNumber, conClassName, ErrorString
    Else
        Err.Raise vbError + ErrorNumber, conClassName
    End If

End Sub

Sub DatabaseWrapper()

    Dim DB As clsResearchDatabase
    Dim DatabaseFullFilename As String
    Dim Sep As String
    Dim tempAnswer As Long
    
    Sep = Application.PathSeparator
    
    DatabaseFullFilename = DatabasePath & Sep & DatabaseName
    
    If Not FileSystem_FileExists(DatabasePath) Then
        
        tempAnswer = VBA.MsgBox("Database folder not found. Create a new folder/database?", vbOKCancel + vbExclamation)
    
        Select Case tempAnswer
            Case vbOK
                Call FileSystem_CreatePath(DatabasePath)
                Call CreateEmptyFile(DatabaseFullFilename)
            Case vbCancel
                Call MsgBox("Program aborted", vbOKOnly + vbInformation)
            Case Else
                Call ErrorReport(, "Unsupported answer to query")
        End Select
        
    Else
    
        If Not FileSystem_FileExists(DatabaseFullFilename) Then
        
            tempAnswer = VBA.MsgBox("Database file not found. Create a new database?", vbOKCancel + vbExclamation)
        
            Select Case tempAnswer
                Case vbOK
                    Call FileSystem_CreatePath(DatabasePath)
                    Call CreateEmptyFile(DatabaseFullFilename)
                Case vbCancel
                    Call MsgBox("Program aborted - change path if neccessary", vbOKOnly + vbInformation)
                Case Else
                    Call ErrorReport(, "Unsupported answer to query")
            End Select
            
        End If
        
    End If
        
    
    Set DB = New clsResearchDatabase
    
    Call DatabaseRun_A(DB)

End Sub

Sub DatabaseRun_A(ByRef DB As clsResearchDatabase)

    Call DB.AddNewElement("TestElement1", "1")

End Sub
