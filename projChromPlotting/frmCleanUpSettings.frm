VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCleanUpSettings 
   Caption         =   "CleanUp Settings"
   ClientHeight    =   9252.001
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5952
   OleObjectBlob   =   "frmCleanUpSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCleanUpSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2016-06-13,
'Last update 2016-06-13
'2016-06-14 add a more robust setting import capability
'====================================================================================================
'2016-11-16 add option for UV autozero
'2017-01-13 "Trunc_end_opt2_Volume.Value" in Import Settings was "Normvol_opt2_Volume.Value", making wrong stuff get imported! -> corrected"

Option Explicit

Private Const conAutozeroUV As String = "AUTOZEROUV"
Private Const conNormUV As String = "NORMUV"
Private Const conNormVol As String = "NORMVOL"
Private Const conTrunc As String = "TRUNC"
Private Const conThin As String = "THIN"
Private Const conAlign As String = "ALIGN"


Public SEC As clsGeneralizedChromatography
Public defaultCleanUpOptions As VBA.Collection
Public tempCleanUpOptions As VBA.Collection

Private CollUVOff As VBA.Collection
Private CollNormUV As VBA.Collection
Private CollNormVol As VBA.Collection
Private CollTrunc As VBA.Collection
Private CollThin As VBA.Collection
Private CollAlign As VBA.Collection

Private IAmInitialized As Boolean
Private InjectionsExist As Boolean
Private ColumnVolumeExists As Boolean

Private HackNumericInputCollection As VBA.Collection


Public Sub ManualInitialize()

    Call UserForm_Initialize
    
    Set HackNumericInputCollection = New VBA.Collection
    
    With HackNumericInputCollection
        .Add UV_Offset_Val
        .Add NormVol_opt1_Volume
        .Add NormVol_opt2_Volume
        .Add Trunc_end_opt1_Volume
        .Add Trunc_end_opt2_Volume
        .Add Thin_opt1_Detail
        .Add Thin_opt2_Detail
        .Add NormUV_version_opt1_Detail
        .Add NormUV_version_opt2_Detail
        .Add NormUV_region_opt2_StartValue
        .Add NormUV_region_opt2_EndValue
    End With
    
    IAmInitialized = False
    
    If Not ((SEC Is Nothing) Or (defaultCleanUpOptions Is Nothing)) Then
                
        'allow picking of injections if there are any, set default (last injection)
        InjectionsExist = False
        If Not SEC.Injections Is Nothing Then
            If SEC.Injections.Count > 0 Then
                InjectionsExist = True
            End If
        End If
        
        If InjectionsExist Then
            With NormVol_opt1_Injection
                .List = SEC.Injections.AnnotationArray
            End With
        Else
            Call DisableInjections
        End If
        
        'allow picking of Column Volume option if CV is provided in the data
        If SEC.Metadata.ColumnVolume > 0 Then
            Trunc_end_opt1_Volume = SEC.Metadata.ColumnVolume
            Trunc_end_opt2_Volume = SEC.Chromatograms.Item(1).Xmax
            ColumnVolumeExists = True
        Else
            ColumnVolumeExists = False
            Trunc_end_opt1.Visible = False
            Trunc_end_opt1_Volume.Visible = False
            Trunc_end_opt3.Value = True
        End If
                           
        'Call DefineDefaults
        'Call SettingsCopy(defaultCleanUpOptions, tempCleanUpOptions)
                           
        Call RevertToDefaultSettings
        
        'NormUV_version_opt1_Detail.Value = 1
        'NormUV_version_opt2_Detail.Value = 1
        
        IAmInitialized = True
        
    End If

End Sub

Private Sub SettingsCopy( _
    ByVal InputSettings As VBA.Collection, _
    ByRef TargetSettings As VBA.Collection)

    Dim i As Long, j As Long
    Dim tempInput As VBA.Collection
    Dim tempCollection As VBA.Collection
    Dim tempIDs As VBA.Collection
    Dim tempID As String
    
    Set tempIDs = New VBA.Collection
    
    With tempIDs
        .Add conAutozeroUV
        .Add conNormVol
        .Add conAlign
        .Add conTrunc
        .Add conThin
        .Add conNormUV
    End With
    
    For j = 1 To tempIDs.Count
    
        tempID = tempIDs.Item(j)
        
        'grab the subsetting collection, prepare a fresh collection
        Set tempInput = InputSettings.Item(tempID)
        Set tempCollection = New VBA.Collection
        
        'copy all the elements from the subsetting input to output
        For i = 1 To tempInput.Count
            tempCollection.Add tempInput(i)
        Next i
        
        'if the target already has the subsetting, remove it
        If IsElementOf(tempID, TargetSettings) Then
            Call TargetSettings.Remove(tempID)
        End If
        
        'add the subsetting to target setting collection
        Call TargetSettings.Add(tempCollection, tempID)
        
    
    Next j
        
    Set tempIDs = Nothing
    Set tempInput = Nothing
    Set tempCollection = Nothing

End Sub

Private Sub DisableInjections()

    InjectionsExist = False
    
    NormVol_opt1.Visible = False
    NormVol_opt1_Injection.Visible = False
    NormVol_opt1_Volume.Visible = False
    
    NormVol_opt2.Value = True

End Sub

Public Sub DefineDefaults()
'defines defaults for all settings

    Dim tempSetting As VBA.Collection
    
    '===AUTOZEROUV
    '1: [do I offset] 2: [how much]
        Set tempSetting = New VBA.Collection
        defaultCleanUpOptions.Add tempSetting, conAutozeroUV
        With tempSetting
            .Add True
            If Not SEC.Chromatograms.Item(1) Is Nothing Then
                .Add CDbl(-1 * SEC.Chromatograms.Item(1).Ymin)
            Else
                .Add 0
            End If
        End With
    
    
    '===NORMALIZE VOLUME
    '1: [do I normalize] 2: [which injection] 3: [which volume]
        Set tempSetting = New VBA.Collection
        defaultCleanUpOptions.Add tempSetting, conNormVol
        With tempSetting
            .Add True
            If Not SEC.Injections Is Nothing Then
                If SEC.Injections.Count > 0 Then
                    InjectionsExist = True
                End If
            End If
            If InjectionsExist Then
                .Add SEC.Injections.Annotation(SEC.Injections.Count)
            Else
                .Add ""
            End If
            .Add CDbl(0)
        End With
    
    '===ALIGN
    '1: [do I align]
        Set tempSetting = New VBA.Collection
        defaultCleanUpOptions.Add tempSetting, conAlign
        With tempSetting
            .Add True
        End With
    
    '===TRUNCATE DATA
    '1: [do I trunc] 2: [start volume] 3: [end volume]
        Set tempSetting = New VBA.Collection
        defaultCleanUpOptions.Add tempSetting, conTrunc
        With tempSetting
            .Add True
            .Add CDbl(0)
            If ColumnVolumeExists Then
                .Add SEC.Metadata.ColumnVolume
            Else
                .Add SEC.Chromatograms.Item(1).Xmax
            End If
        End With
    
    '===THIN DATA
    '1: [do I thin] 2: [distance between points]
        Set tempSetting = New VBA.Collection
        defaultCleanUpOptions.Add tempSetting, conThin
        With tempSetting
            .Add True
            .Add CDbl(0.5)
            .Add SEC.Chromatograms.Item(1).NumberOfPoints
        End With
        
    '===NORMALIZE UV
    '1: [do I norm] 2: [INTEGRAL/MAXVALUE] 3: [startvolume] 4: [endvolume]
        Set tempSetting = New VBA.Collection
        defaultCleanUpOptions.Add tempSetting, conNormUV
        With tempSetting
            .Add False
            .Add "MAXVALUE"
            .Add SEC.Chromatograms.Item(1).Xmin
            .Add SEC.Chromatograms.Item(1).Xmax
        End With
    
    Set tempSetting = Nothing
    
    Call ImportSettings(defaultCleanUpOptions)

End Sub

Private Sub ImportSettings(ChosenSettings As VBA.Collection)
'sets the form controls to their real settings

    Dim tempSetting As VBA.Collection
    
    With ChosenSettings
        
        '===AUTOZERO UV
        '1: [do I AZ] 2: [manual offset]
            Set tempSetting = .Item(conAutozeroUV)
            With tempSetting
                Autozero_tick = .Item(1)
                UV_Offset_Val = .Item(2)
            End With
            
        
        '===NORMALIZE VOLUME
        '1: [do I normalize] 2: [which injection] 3: [which volume]
            Set tempSetting = .Item(conNormVol)
            With tempSetting
                NormVol_tick = .Item(1)
                NormVol_opt1_Injection.Value = .Item(2)
                NormVol_opt2_Volume.Value = .Item(3)
            End With
            
        '===ALIGN
        '1: [do I align]
            Set tempSetting = .Item(conAlign)
            With tempSetting
                Align_tick = .Item(1)
            End With
                 
        '===TRUNCATE DATA
        '1: [do I trunc] 2: [start volume] 3: [end volume]
            Set tempSetting = .Item(conTrunc)
            With tempSetting
                Trunc_tick = .Item(1)
                If .Item(2) = 0 Then
                    Trunc_start_opt1.Value = True
                Else
                    Trunc_start_opt2.Value = True
                End If
                Trunc_end_opt2_Volume.Value = .Item(3)
            End With
                 
        '===THIN DATA
        '1: [do I thin] 2: [distance between points]
            Set tempSetting = .Item(conThin)
            With tempSetting
                Thin_tick = .Item(1)
                Thin_opt1_Detail.Value = .Item(2)
                Thin_opt2_Detail.Value = CLng( _
                    (SEC.Chromatograms.Item(1).Xmax - SEC.Chromatograms.Item(1).Xmin) / .Item(2) + 1)
            End With
             
        '===NORMALIZE UV
        '1: [do I norm] 2: [INTEGRAL/MAXVALUE] 3: [startvolume] 4: [endvolume]
            Set tempSetting = .Item(conNormUV)
            With tempSetting
                NormUV_tick = .Item(1)
                Select Case .Item(2)
                    Case "MAXVALUE"
                        NormUV_version_opt1.Value = True
                    Case "INTEGER"
                        NormUV_version_opt2.Value = True
                    Case Else
                        Err.Raise 1001, , "Unknown setting detected (frmCleanUpSettings)"
                End Select
                NormUV_region_opt2_StartValue.Value = .Item(3)
                NormUV_region_opt2_EndValue.Value = .Item(4)
            End With
            
    End With
    
    Set tempSetting = Nothing

End Sub

Private Sub RevertToDefaultSettings()
    
    If defaultCleanUpOptions Is Nothing Then
        Set defaultCleanUpOptions = New VBA.Collection
    End If
    
    If defaultCleanUpOptions.Count = 0 Then
        Call DefineDefaults
    End If

    Call ImportSettings(defaultCleanUpOptions)

End Sub


Private Sub Autozero_tick_Click()
    If Autozero_tick.Value = True Then
        UV_Offset_Val.Enabled = True
    Else
        UV_Offset_Val.Enabled = False
        UV_Offset_Val.Value = -1 * SEC.Chromatograms.Item(1).Ymin
    End If
        
End Sub


Private Sub UV_Offset_Val_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not ((IsNumeric(UV_Offset_Val.Value)) Or Len(UV_Offset_Val.Value) = 0) Then
        Cancel = 1
        Call MsgBox("The input value must be numeric!", vbExclamation + vbOKOnly)
        UV_Offset_Val.SetFocus
    End If

End Sub

Private Sub ctrlReset_Click()

    Call RevertToDefaultSettings

End Sub

Private Sub ctrlSave_Click()

    Unload Me
    
End Sub

Private Sub NormVol_opt1_Injection_Change()
    NormVol_opt1_Volume = SEC.Injections.XData(CLng(NormVol_opt1_Injection.Value))
End Sub

Private Sub NormVol_tick_AfterUpdate()

End Sub

'===============================NORM VOLUME===============================

Private Sub NormVol_tick_Change()

    If NormVol_tick.Value = True Then
    
        NormVol_opt1.Enabled = True
        NormVol_opt1_Injection.Enabled = True
        NormVol_opt1_Volume.Enabled = True
        
        NormVol_opt2.Enabled = True
        NormVol_opt2_Volume.Enabled = True
        
    Else
    
        NormVol_opt1.Enabled = False
        NormVol_opt1_Injection.Enabled = False
        NormVol_opt1_Volume.Enabled = False
        
        NormVol_opt2.Enabled = False
        NormVol_opt2_Volume.Enabled = False
        
    End If
    
End Sub

Private Sub NormVol_opt2_Volume_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not ((IsNumeric(NormVol_opt2_Volume.Value)) Or Len(NormVol_opt2_Volume.Value) = 0) Then
        Cancel = 1
        Call MsgBox("The input value must be numeric!", vbExclamation + vbOKOnly)
        NormVol_opt2_Volume.SetFocus
    End If

End Sub


'===============================TRUNC===============================

Private Sub Trunc_tick_Change()

    If Trunc_tick.Value = True Then
        
        Trunc_start_label.Enabled = True
        
        Trunc_start_opt1.Enabled = True
        Trunc_start_opt2.Enabled = True
        
        
        Trunc_end_label.Enabled = True
        
        Trunc_end_opt1.Enabled = True
        Trunc_end_opt1_Volume.Enabled = True
        Trunc_end_opt2.Enabled = True
        Trunc_end_opt2_Volume.Enabled = True
        Trunc_end_opt3.Enabled = True
        
    Else
    
        Trunc_start_label.Enabled = False
        
        Trunc_start_opt1.Enabled = False
        Trunc_start_opt2.Enabled = False
        
        
        Trunc_end_label.Enabled = False
        
        Trunc_end_opt1.Enabled = False
        Trunc_end_opt1_Volume.Enabled = True
        Trunc_end_opt2.Enabled = False
        Trunc_end_opt2_Volume.Enabled = False
        Trunc_end_opt3.Enabled = False

    End If

End Sub

Private Function CheckNumericOrEmpty(ByVal ControlName As String, ByRef Cancel As MSForms.ReturnBoolean) As Boolean

    If Not ((IsNumeric(Controls.Item(ControlName).Value) Or Len(Controls.Item(ControlName).Value)) = 0) Then
        Cancel = 1
        Call MsgBox("The input value must be numeric!", vbExclamation + vbOKOnly)
        Controls.Item(ControlName).SetFocus
    End If

End Function

Private Sub Trunc_end_opt2_Volume_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not ((IsNumeric(Trunc_end_opt2_Volume.Value)) Or Len(Trunc_end_opt2_Volume.Value) = 0) Then
        Cancel = 1
        Call MsgBox("The input value must be numeric!", vbExclamation + vbOKOnly)
        Trunc_end_opt2_Volume.SetFocus
    End If

End Sub

'===============================THIN===============================
Private Sub Thin_tick_Change()

    If Thin_tick.Value = True Then
        
        Thin_opt1.Enabled = True
        Thin_opt1_Detail.Enabled = True
        
        Thin_opt2.Enabled = True
        Thin_opt2_Detail.Enabled = True
        
    Else
    
        Thin_opt1.Enabled = False
        Thin_opt1_Detail.Enabled = False
        
        Thin_opt2.Enabled = False
        Thin_opt2_Detail.Enabled = False

    End If

End Sub

Private Sub Thin_opt1_Detail_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not ((IsNumeric(Thin_opt1_Detail.Value)) Or Len(Thin_opt1_Detail.Value) = 0) Then
        Cancel = 1
        Call MsgBox("The input value must be numeric!", vbExclamation + vbOKOnly)
        Thin_opt1_Detail.SetFocus
    Else
        If Thin_opt2_Detail.Value <= 0 And Len(Thin_opt2_Detail.Value) > 0 Then
            Cancel = 1
            Call MsgBox("The input value must be greater than zero!", vbExclamation + vbOKOnly)
            Thin_opt1_Detail.SetFocus
        End If
    End If
    

End Sub

Private Sub Thin_opt2_Detail_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not ((IsNumeric(Thin_opt2_Detail.Value)) Or Len(Thin_opt2_Detail.Value) = 0) Then
        Cancel = 1
        Call MsgBox("The input value must be numeric!", vbExclamation + vbOKOnly)
        Thin_opt2_Detail.SetFocus
    Else
        If Thin_opt2_Detail.Value <= 30 Then
            Cancel = 1
            Call MsgBox("The input value must be greater than 30!", vbExclamation + vbOKOnly)
            Thin_opt2_Detail.SetFocus
        End If
    End If

End Sub


'===============================NORM UV===============================
Private Sub NormUV_tick_Change()

    If NormUV_tick.Value = True Then
    
        NormUV_version_opt1.Enabled = True
        NormUV_version_opt1_Detail.Enabled = True
        
        NormUV_version_opt2.Enabled = True
        NormUV_version_opt2_Detail.Enabled = True
        
        NormUV_region_label.Enabled = True
        
        NormUV_region_opt1.Enabled = True
        
        NormUV_region_opt2.Enabled = True
        
        NormUV_region_opt2_StartLabel.Enabled = True
        NormUV_region_opt2_StartValue.Enabled = True
        
        NormUV_region_opt2_EndLabel.Enabled = True
        NormUV_region_opt2_EndValue.Enabled = True
        
        NormUV_region_notice.Enabled = True
    
    Else
    
        NormUV_version_opt1.Enabled = False
        NormUV_version_opt1_Detail.Enabled = False
        
        NormUV_version_opt2.Enabled = False
        NormUV_version_opt2_Detail.Enabled = False
        
        NormUV_region_label.Enabled = False
        
        NormUV_region_opt1.Enabled = False
        
        NormUV_region_opt2.Enabled = False
        
        NormUV_region_opt2_StartLabel.Enabled = False
        NormUV_region_opt2_StartValue.Enabled = False
        
        NormUV_region_opt2_EndLabel.Enabled = False
        NormUV_region_opt2_EndValue.Enabled = False
        
        NormUV_region_notice.Enabled = False
        
    End If

End Sub

Private Sub NormUV_version_opt1_Detail_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not ((IsNumeric(NormUV_version_opt1_Detail.Value)) Or Len(NormUV_version_opt1_Detail.Value) = 0) Then
        Cancel = 1
        Call MsgBox("The input value must be numeric!", vbExclamation + vbOKOnly)
        NormUV_version_opt1_Detail.SetFocus
    End If

End Sub

Private Sub NormUV_version_opt2_Detail_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not ((IsNumeric(NormUV_version_opt2_Detail.Value)) Or Len(NormUV_version_opt2_Detail.Value) = 0) Then
        Cancel = 1
        Call MsgBox("The input value must be numeric!", vbExclamation + vbOKOnly)
        NormUV_version_opt2_Detail.SetFocus
    End If

End Sub

Private Sub NormUV_region_opt2_StartValue_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not ((IsNumeric(NormUV_region_opt2_StartValue.Value)) Or Len(NormUV_region_opt2_StartValue.Value) = 0) Then
        Cancel = 1
        Call MsgBox("The input value must be numeric!", vbExclamation + vbOKOnly)
        NormUV_region_opt2_StartValue.SetFocus
    End If

End Sub

Private Sub NormUV_region_opt2_EndValue_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not ((IsNumeric(NormUV_region_opt2_EndValue.Value)) Or Len(NormUV_region_opt2_EndValue.Value) = 0) Then
        Cancel = 1
        Call MsgBox("The input value must be numeric!", vbExclamation + vbOKOnly)
        NormUV_region_opt2_EndValue.SetFocus
    End If

End Sub


'===============================GENERAL===============================

Private Sub UserForm_Initialize()
            
    'listRegions.RemoveItem (0)
    
End Sub


Private Sub HackyHackOfHacks()

    '#20161125
    'THIS IS A HACK AND UGLY, YUCK
        'go through all inputboxes supposed to be numeric and replace faulty formatted ones
        'with a standard one (decimal comma, no thousand separators). This is only needed because
        'apparently sometimes, but in a non-100 % reproducible fashion, on some machines, Forms
        'convert inputs to a wrong format. This can _change_ without changing any code!
        
    Dim HackInputBox As Control
    Dim i As Long
    Dim temptext As String
    Dim dotcount As Long
    
        For i = 1 To HackNumericInputCollection.Count
            Set HackInputBox = HackNumericInputCollection.Item(i)
            With HackInputBox
                dotcount = StringCharCount(.Text, ".")
                If dotcount > 0 Then
                    If StringCharCount(.Text, ",") > 0 Then
                        Call Err.Raise(1, "Some crappy error with decimal stuff, excel bug")
                    Else
                        Debug.Print ("Bad formatting: " & .Text & ", replace with standard format.")
                        'remove all dots but the last one
                        .Text = Replace(.Text, ".", "", 1, dotcount - 1)
                        'convert last dot to comma
                        .Text = Replace(.Text, ".", ",", 1, 1)
                    End If
                End If
            End With
        Next i
    '\THIS IS A HACK AND UGLY, YUCK

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Dim InputValuesMakeSense As Boolean
    Dim tempMsg As String
    Dim tempValue As Double
    Dim ControlToFocus As Control
       
    Call HackyHackOfHacks
    
    'AUTOZERO UV (UV OFFSET)
    '1: [do I?] 2: [offset value]
    Set CollUVOff = New VBA.Collection
        With CollUVOff
            .Add Autozero_tick.Value
            .Add UV_Offset_Val.Value
        End With
    
    'VOLUME NORMALIZATION
    '1: [do I normalize] 2: [which injection] 3: [which volume]
    Set CollNormVol = New VBA.Collection
        With CollNormVol
            .Add NormVol_tick.Value
            If NormVol_tick.Value = True Then
                If NormVol_opt1.Value = True Then
                    .Add CLng(NormVol_opt1_Injection.Value)
                    .Add CDbl(NormVol_opt1_Volume.Value)
                Else
                    .Add ""
                    .Add NormVol_opt2_Volume.Value
                End If
            Else
                .Add defaultCleanUpOptions.Item(conNormVol).Item(2)
                .Add defaultCleanUpOptions.Item(conNormVol).Item(3)
            End If
        End With
    
    'DATA ALIGNMENT
    '1: [do I align]
    Set CollAlign = New VBA.Collection
        CollAlign.Add Align_tick.Value
        
    'DATA TRUNCATION
    '1: [do I trunc] 2: [start volume] 3: [end volume]
    Set CollTrunc = New VBA.Collection
        With CollTrunc
            .Add Trunc_tick.Value
            If Trunc_tick.Value = True Then
                If Trunc_start_opt1.Value = True Then
                    .Add CDbl(0)
                Else
                    .Add SEC.Chromatograms.Item(1).Xmin
                End If
                If Trunc_end_opt1.Value = True Then
                    .Add CDbl(SEC.Metadata.ColumnVolume)
                ElseIf Trunc_end_opt2.Value = True Then
                    .Add CDbl(Trunc_end_opt2_Volume.Value)
                Else
                    .Add CDbl(SEC.Chromatograms.Item(1).Xmax)
                End If
            Else
                .Add defaultCleanUpOptions.Item(conTrunc).Item(2)
                .Add defaultCleanUpOptions.Item(conTrunc).Item(3)
            End If
        End With
        
    'DATA THINNING
    '1: [do I thin] 2: [distance between points]
    Set CollThin = New VBA.Collection
        With CollThin
            .Add Thin_tick.Value
            If Thin_tick.Value = True Then
                If Thin_opt1.Value = True Then
                    .Add CDbl(Thin_opt1_Detail.Value)
                Else
                    'If Thin_opt2_Detail.Value <> SEC.Chromatograms.Item(1).NumberOfPoints Then
                    tempValue = SEC.Chromatograms.Item(1).Xmax - SEC.Chromatograms.Item(1).Xmin
                    'End If
                    .Add CDbl((tempValue) / (CLng(Thin_opt2_Detail.Value) - 1))
                End If
            Else
                .Add defaultCleanUpOptions.Item(conThin).Item(2)
            End If
        End With
        
    
    'UV NORMALIZATION
    '1: [do I norm] 2: [INTEGRAL/MAXVALUE] 3: [startvolume] 4: [endvolume]
    Set CollNormUV = New VBA.Collection
        With CollNormUV
            .Add NormUV_tick.Value
            If NormUV_tick.Value = True Then
                If NormUV_tick.Value = True Then
                    If NormUV_version_opt2.Value = True Then
                        .Add "INTEGRAL"
                    Else
                        .Add "MAXVALUE"
                    End If
                    If NormUV_region_opt1.Value = True Then
                        .Add SEC.Chromatograms.Item(1).Xmin
                        .Add SEC.Chromatograms.Item(1).Xmax
                    Else
                        .Add CDbl(NormUV_region_opt2_StartValue.Value)
                        .Add CDbl(NormUV_region_opt2_EndValue.Value)
                    End If
                End If
            Else
                .Add defaultCleanUpOptions.Item(conNormUV).Item(2)
                .Add defaultCleanUpOptions.Item(conNormUV).Item(3)
                .Add defaultCleanUpOptions.Item(conNormUV).Item(4)
            End If
        End With
        
    'store the settings into the object
    Set tempCleanUpOptions = New VBA.Collection
    Call DefineCleanUpOptions(tempCleanUpOptions)
    
    '=====test whether inputs make sense
    
    InputValuesMakeSense = True
    tempMsg = ""
    
    'TODO: TEMPORARILY REMOVED CHECKS - CONFLICT WHEN
    'normalization volume check
    If tempCleanUpOptions.Item(conNormVol).Item(1) = True Then
        If tempCleanUpOptions.Item(conNormVol).Item(3) < SEC.Chromatograms.Item(1).Xmin Then
            tempMsg = "Normalization volume cannot be below the lowest data point (" & _
                SEC.Chromatograms.Item(1).Xmin & " " & SEC.Chromatograms.Item(1).XAxisInfo.Unit & ")"
            Call MsgBox(tempMsg, vbCritical + vbOKOnly, "Error in input")
            'InputValuesMakeSense = False
            'Cancel = 1
            Set ControlToFocus = NormVol_opt2_Volume
        End If
        
        If tempCleanUpOptions.Item(conNormVol).Item(3) > SEC.Chromatograms.Item(1).Xmax Then
            tempMsg = "Normalization volume cannot be above the largest data point (" & _
                SEC.Chromatograms.Item(1).Xmax & " " & SEC.Chromatograms.Item(1).XAxisInfo.Unit & ")"
            Call MsgBox(tempMsg, vbCritical + vbOKOnly, "Error in input")
            'InputValuesMakeSense = False
            'Cancel = 1
            Set ControlToFocus = NormVol_opt2_Volume
        End If
    End If
        
    'normalization UV check
    If tempCleanUpOptions.Item(conNormUV).Item(1) = True Then
        If tempCleanUpOptions.Item(conNormUV).Item(3) < SEC.Chromatograms.Item(1).Xmin Then
            tempMsg = "UV normalization volume cannot be below the lowest data point (" & _
                SEC.Chromatograms.Item(1).Xmin & " " & SEC.Chromatograms.Item(1).XAxisInfo.Unit & ")"
            Call MsgBox(tempMsg, vbCritical + vbOKOnly, "Error in input")
            'InputValuesMakeSense = False
            'Cancel = 1
            Set ControlToFocus = NormUV_region_opt2_StartValue
        End If
        
        If tempCleanUpOptions.Item(conNormUV).Item(4) > SEC.Chromatograms.Item(1).Xmax Then
            tempMsg = "UV normalization volume cannot be above the largest data point (" & _
                SEC.Chromatograms.Item(1).Xmax & " " & SEC.Chromatograms.Item(1).XAxisInfo.Unit & ")"
            Call MsgBox(tempMsg, vbCritical + vbOKOnly, "Error in input")
            'InputValuesMakeSense = False
            'Cancel = 1
            Set ControlToFocus = NormUV_region_opt2_EndValue
        End If
        
        If tempCleanUpOptions.Item(conNormUV).Item(4) <= tempCleanUpOptions.Item(conNormUV).Item(4) Then
            tempMsg = "Start volume for UV normalization must be lower than end volume!"
            Call MsgBox(tempMsg, vbCritical + vbOKOnly, "Error in input")
            'InputValuesMakeSense = False
            'Cancel = 1
            Set ControlToFocus = NormUV_region_opt2_StartValue
        End If
    End If
        
        
    'finalize
    If Not InputValuesMakeSense Then
        Call EraseSettings(tempCleanUpOptions)
        ControlToFocus.SetFocus
    Else
        Call SettingsCopy(tempCleanUpOptions, defaultCleanUpOptions)
    End If
    
    Set ControlToFocus = Nothing
    
End Sub

Private Sub EraseSettings(ChosenSettings As VBA.Collection)

    Dim i As Long

    For i = 1 To ChosenSettings.Count
        ChosenSettings.Remove (1)
    Next i

End Sub

Private Sub DefineCleanUpOptions(ChosenSettings As VBA.Collection)

        'Set CleanUpOptions = New VBA.Collection
        
        Call EraseSettings(ChosenSettings)
        
        With ChosenSettings
        
            If Not CollUVOff Is Nothing Then
                .Add CollUVOff, conAutozeroUV
            End If
            
            If Not CollNormUV Is Nothing Then
                .Add CollNormUV, conNormUV
            End If
            
            If Not CollNormVol Is Nothing Then
                .Add CollNormVol, conNormVol
            End If
            
            If Not CollTrunc Is Nothing Then
                .Add CollTrunc, conTrunc
            End If
            
            If Not CollThin Is Nothing Then
                .Add CollThin, conThin
            End If
            
            If Not CollAlign Is Nothing Then
                .Add CollAlign, conAlign
            End If
            
        End With

End Sub


Private Sub UserForm_Terminate()
           
    Dim i As Long
    
    Call EraseSettings(tempCleanUpOptions)
    
    Set SEC = Nothing
    Set tempCleanUpOptions = Nothing
    Set defaultCleanUpOptions = Nothing
    
    Set CollNormUV = Nothing
    Set CollNormVol = Nothing
    Set CollTrunc = Nothing
    Set CollThin = Nothing
    Set CollAlign = Nothing
    
End Sub



Private Sub UV_Offset_Val_Change()

End Sub
