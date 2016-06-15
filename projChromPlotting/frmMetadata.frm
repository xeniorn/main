VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMetadata 
   Caption         =   "Experiment details"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6624
   OleObjectBlob   =   "frmMetadata.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2016-05-20,
'Last update 2016-05-27
'====================================================================================================


Option Explicit

Const conNumberOfEntries As Long = 6

Private LineCollection As Collection

Private MetaDataNames(1 To conNumberOfEntries) As String
Private DefaultValue As Collection

Public SEC As clsGeneralizedChromatography

Private Sub CheckBox1_Click(): Call SetState: End Sub
Private Sub CheckBox2_Click(): Call SetState: End Sub
Private Sub CheckBox3_Click(): Call SetState: End Sub
Private Sub CheckBox4_Click(): Call SetState: End Sub
Private Sub CheckBox5_Click(): Call SetState: End Sub
Private Sub CheckBox6_Click(): Call SetState: End Sub



Private Sub SetState()

    For i = 1 To conNumberOfEntries
        With LineCollection.Item(i)
            If .Item("CHECKBOX").Value = True Then
                With .Item("TEXTBOX")
                    .Enabled = False
                    .BackColor = &H8000000F
                    .Text = DefaultValue.Item(MetaDataNames(i))
                End With
            Else
                With .Item("TEXTBOX")
                    .Enabled = True
                    .BackColor = &H80000005
                End With
            End If
        End With
    Next i
        
                

End Sub

Public Sub ManualInitialize()
    
    Set DefaultValue = New Collection
    
    With DefaultValue
        'column volume
            If SEC.Metadata.ColumnVolume <> 0 Then
                .Add SEC.Metadata.ColumnVolume, MetaDataNames(1)
            Else
                .Add "use full chromatogram volume", MetaDataNames(1)
            End If
        'loop volume
            If SEC.Metadata.SampleVolume <> 0 Then
                .Add SEC.Metadata.SampleVolume, MetaDataNames(2)
            Else
                .Add "unknown (as if zero)", MetaDataNames(2)
            End If
        'system
            If SEC.Metadata.SystemUsed <> "" Then
                .Add SEC.Metadata.SystemUsed, MetaDataNames(3)
            Else
                .Add "unknown", MetaDataNames(3)
            End If
        'Scientist
            If SEC.Metadata.ExperimentScientist <> "" Then
                .Add SEC.Metadata.ExperimentScientist, MetaDataNames(4)
            Else
                .Add "Juraj Ahel", MetaDataNames(4)
            End If
        'Experiment name
            If SEC.Metadata.ExperimentName <> "" Then
                .Add SEC.Metadata.ExperimentName, MetaDataNames(5)
            Else
                .Add "unknown", MetaDataNames(5)
            End If
        'Date
            If SEC.Metadata.ExperimentDate <> "" Then
                .Add SEC.Metadata.ExperimentDate, MetaDataNames(6)
            Else
                .Add "unknown", MetaDataNames(6)
            End If
    End With
    
    Call SetState

End Sub



Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not (IsNumeric(TextBox1.Value)) Then
        Cancel = 1
        Call MsgBox("The input value must be numeric!", vbExclamation + vbOKOnly)
    End If
End Sub

Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not (IsNumeric(TextBox2.Value)) Then
        Cancel = 1
        Call MsgBox("The input value must be numeric!", vbExclamation + vbOKOnly)
    End If
End Sub

Private Sub UserForm_Initialize()

    Dim i As Long

    Dim tempCollection As Collection
        
    Set LineCollection = New Collection
        
    'create the collection of controls, with each line as separate
    'subcollection
    For i = 1 To conNumberOfEntries
        
        MetaDataNames(i) = Left(Controls("Label" & i).Caption, Len(Controls("Label" & i).Caption))
        
        Set tempCollection = New Collection
            tempCollection.Add Controls("Label" & i), "LABEL"
            tempCollection.Add Controls("TextBox" & i), "TEXTBOX"
            tempCollection.Add Controls("CheckBox" & i), "CHECKBOX"
            
        LineCollection.Add tempCollection, MetaDataNames(i)
        
    Next i
    
    Set tempCollection = Nothing
    

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Dim answer As Long
    
    If CloseMode = 0 Then
    
        answer = MsgBox("Upon exiting this form, all metadata will be " & _
                    "replaced with any inputs you provided. Continue?", _
                    vbQuestion + vbOKCancel, "Confirm data input")
        
        Select Case answer
        
            Case vbCancel
                Cancel = 1
                
            Case vbOK
                With LineCollection.Item(1)
                    If .Item("CHECKBOX").Value = False And IsNumeric(.Item("TEXTBOX").Value) Then
                        SEC.Metadata.ColumnVolume = CDbl(.Item("TEXTBOX").Value)
                    End If
                End With
                With LineCollection.Item(2)
                    If .Item("CHECKBOX").Value = False And IsNumeric(.Item("TEXTBOX").Value) Then
                        SEC.Metadata.SampleVolume = CDbl(.Item("TEXTBOX").Value)
                    End If
                End With
                With LineCollection.Item(3)
                    If .Item("CHECKBOX").Value = False Then
                        SEC.Metadata.SystemUsed = .Item("TEXTBOX").Value
                    End If
                End With
                With LineCollection.Item(4)
                    If .Item("CHECKBOX").Value = False Then
                        SEC.Metadata.ExperimentScientist = .Item("TEXTBOX").Value
                    End If
                End With
                With LineCollection.Item(5)
                    If .Item("CHECKBOX").Value = False Then
                        SEC.Metadata.ExperimentName = .Item("TEXTBOX").Value
                    End If
                End With
                With LineCollection.Item(6)
                    If .Item("CHECKBOX").Value = False Then
                        SEC.Metadata.ExperimentDate = .Item("TEXTBOX").Value
                    End If
                End With
                
        End Select
    End If
    
End Sub

Private Sub UserForm_Terminate()

    Set LineCollection = Nothing
    Set DefaultValue = Nothing
    Set SEC = Nothing

End Sub
