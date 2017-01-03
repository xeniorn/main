Attribute VB_Name = "MASTER"
'Juraj Ahel, 2016-06-15
'This is a special module, which can export / import all modules, forms, and classes
'to my GIT repository folder or import back here. It's a bit crappy in that it has a hardcoded name
'(for ignoring itself when importing / exporting) and hardcoded GIT path
'also, if it crashes upon import, the local version is lost

'###depends on:
'Microsoft VB Extensibility library
'

Option Explicit

Private Const GITPath As String = "C:\Users\juraj.ahel\Documents\GitHub\main\"

Public Sub ExportAllComponents()

'reference to extensibility library

    Dim objMyProj As VBProject
    Dim objVBComp As VBComponent
    
    Dim FullPath As String
    Dim tempExtension As String
             
    Set objMyProj = Application.VBE.ActiveVBProject
    '
    FullPath = GITPath & objMyProj.Name & Application.PathSeparator
    
    FullPath = Replace(FullPath, "\", Application.PathSeparator)
    FullPath = Replace(FullPath, "/", Application.PathSeparator)
    
    Call FileSystem_CreatePath(FullPath)
     
    For Each objVBComp In objMyProj.VBComponents
    
        With objVBComp
            
            If .Name <> "MASTER" Then
                
                tempExtension = ""
                
                Select Case .Type
                    Case vbext_ct_StdModule
                        tempExtension = ".bas"
                    Case vbext_ct_ClassModule
                        tempExtension = ".cls"
                    Case vbext_ct_MSForm
                        tempExtension = ".frm"
                End Select
                
                If tempExtension <> "" Then
                    .Export FullPath & .Name & tempExtension
                End If
                
            End If
            
        End With
    
    Next objVBComp
    
    Set objMyProj = Nothing
    Set objVBComp = Nothing
     
End Sub

Public Sub ImportAllComponents()

' reference to extensibility library

    Dim objMyProj As VBProject
    Dim objVBComp As VBComponent
    
    Dim FullPath As String
    Dim tempFile As String
    Dim ext As String
    Dim AllFiles As VBA.Collection
    Dim i As Long
    
    Dim answer As Long
    
    answer = MsgBox( _
             "The macro will import all the modules, classes, and forms from the source to this project." & _
             vbCrLf & _
             "Files that exist in this module that have the same name will be lost! Continue?" & _
             vbCrLf & _
             "Source = " & GITPath, _
             vbYesNo + vbCritical, _
             "Confirm import")
        
    If answer = vbYes Then
        answer = MsgBox("Are you SURE?", vbYesNo + vbExclamation, "REALLY?")
    Else
        Exit Sub
    End If
    
    If answer = vbYes Then
        
                 
        Set objMyProj = Application.VBE.ActiveVBProject
        '
        FullPath = GITPath & objMyProj.Name & "\"
        
        FullPath = Replace(FullPath, "\", Application.PathSeparator)
         
        Set AllFiles = FileSystem_GetDirContents(FullPath, IncludePath:=True)
         
        Dim ttt As String
         
        For i = 1 To AllFiles.Count
        
            tempFile = AllFiles.Item(i)
        
            ttt = Mid(tempFile, InStrRev(tempFile, "\") + 1, Len(tempFile))
            ext = Mid(tempFile, InStrRev(tempFile, ".") + 1, Len(tempFile))
            ttt = Left(ttt, InStrRev(ttt, ".") - 1)
            
            Select Case UCase(ext)
                Case "FRM", "BAS", "CLS"
                    
                    If ObjectExists(ttt) Then
                        Set objVBComp = objMyProj.VBComponents.Item(ttt)
                        Call objVBComp.Export("tempy")
                        Call objMyProj.VBComponents.Remove(objVBComp)
                        Set objVBComp = Nothing
                    End If
                            
                    Call objMyProj.VBComponents.Import(tempFile)
                    
            End Select
                            
        Next i
        
        Set objMyProj = Nothing
        Set objVBComp = Nothing
        
    End If
     
End Sub


Private Function ObjectExists(ByVal Name As String) As Boolean

    Dim i As Long
    
    Dim objMyProj As VBProject
    Dim objVBComp As VBComponent
    
    Set objMyProj = Application.VBE.ActiveVBProject
    
    ObjectExists = False
    
    For Each objVBComp In objMyProj.VBComponents
        If objVBComp.Name = Name Then
            ObjectExists = True
            Exit For
        End If
    Next objVBComp
        
End Function

