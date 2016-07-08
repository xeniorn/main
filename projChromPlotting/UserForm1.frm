VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4236
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11472
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DatabaseFolder As String

Private Sub ctrlFileSelection_Click()
'picking the file to import using Windows native file selection dialog

    Dim conFileDialog As FileDialog
    Dim tempFile As String

    Set conFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With conFileDialog
    
        .AllowMultiSelect = False
        
        .InitialFileName = ctrlSelFileTxtBox.Text
                
        .Show
        
        If .SelectedItems.Count > 0 Then
        
            tempFile = .SelectedItems.Item(1)
            
            'MsgBox ("Selected item: " & .SelectedItems.Item(1))
            
            ctrlSelFileTxtBox.Text = tempFile
        
        Else
        
            'ctrlSelFileTxtBox.Text = vbEmpty
            'FileName = vbEmpty
            
        End If
                
        
    End With

End Sub

Private Sub ctrlSelFileTxtBox_Change()
    DatabaseFolder = ctrlSelFileTxtBox.Text
End Sub
