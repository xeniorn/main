VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateChromatograms 
   Caption         =   "Chromatogram import"
   ClientHeight    =   6024
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   11508
   OleObjectBlob   =   "frmCreateChromatograms.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateChromatograms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2016-05-12,
'Last update 2016-06-10
'2016-06-12 simple autodetect filetype
'2016-06-13 code for implementing cleanup settings
'====================================================================================================
'TODO: add settings for cleanup!
'2016-11-16 v0.86 added support for option for autozeroUV / offset UV

Option Explicit

'constants
Const conUni3_title As String = "Unicorn 3-5 .res file"
Const conUni6_title As String = "Unicorn 6+ .zip file"

Const conUni3_ID As String = "UNICORN3"
Const conUni6_ID As String = "UNICORN6"

Private Const conAutozeroUV As String = "AUTOZEROUV"
Private Const conNormUV As String = "NORMUV"
Private Const conNormVol As String = "NORMVOL"
Private Const conTrunc As String = "TRUNC"
Private Const conThin As String = "THIN"
Private Const conAlign As String = "ALIGN"

Const conAllowedCurveTypesNumber = 5

Private AbsoluteMinAllowedVolume As Double
Private AbsoluteMaxAllowedVolume As Double

Private AbsoluteMinAllowedUV As Double
Private AbsoluteMaxAllowedUV As Double

'control variables
Private DataIsImported As Boolean

'public interactors
Public SelectedFiles As FileDialogSelectedItems
Public DefaultStartFolder As String
Public DefaultFileName As String


'internal variables
Private FileName As String
Private FileType As String

Private EEC As Byte

'options for data cleanup
'this is a collection of collections (with keys corresponding to cleanup type)
'each of which contains the neccessary parameters
Private CleanUpOptions As VBA.Collection


'remembers which data goes on which axis
Private AxisAssignment As Collection

'remembers the color setting
Private ColorChoiceSetting(0 To conAllowedCurveTypesNumber - 1) As String

'remembers the actual color
Private ColorChoiceRGB(0 To conAllowedCurveTypesNumber - 1) As Long

Private SelectedInjection As Long


'temporary way to set up and down limits of axes
Private XAxisMin As Double
Private XAxisMax As Double
Private YAxisMin As Double
Private YAxisMax As Double

Private XAutoMin As Double
Private XAutoMax As Double
Private YAutoMin As Double
Private YAutoMax As Double


'Dim SEC As clsSizeExclusionChromatography
Public WithEvents SEC As clsGeneralizedChromatography
Attribute SEC.VB_VarHelpID = -1

'internal
Private SupportedFileTypes() As String

'Meta
Private CurveTickCollection As Collection
Private NamesArray() As String

Private InputFileValid As Boolean

'Events

Public Event EImportCmdDone(IDNumber As Long)

Private Sub UserForm_EImportCmdDone()

    MsgBox ("ID")

End Sub

Private Sub ctrlCleanUpSettings_Click()

    Dim tempFrm As frmCleanUpSettings
    
    If DataIsImported Then
        
        Set tempFrm = New frmCleanUpSettings
        
        With tempFrm
            'import SEC object
                Set .SEC = SEC
                Set .defaultCleanUpOptions = CleanUpOptions
                .ManualInitialize
                .Show (vbModal)
                'Set CleanUpOptions = tempFrm.CleanUpOptions
        End With
        
        
        
        'unload it from memory after it's done
        Unload tempFrm
        
    Else
    
        Call Err.Raise(1001, , "No data imported")
        
    End If

End Sub

Private Sub ctrlButtonExit_Click()

    Dim answer As Long

    answer = MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Exit?")
    
    If answer = vbYes Then
        Debug.Print ("Form closed")
        'MsgBox ("Bye!")
        Unload Me
    Else
        Call MsgBox("Exit form cancelled!", vbInformation + vbOKOnly)
    End If

End Sub


Private Sub ctrlFileType_Change()

    FileType = FileTypeString(ctrlFileType.Value)

End Sub


Private Sub ctrlCleanUpButton_Click()

    Dim tempMsg As String
    Dim answer As Long
    
    Dim i As Long

    If SEC Is Nothing Then
        
        MsgBox ("No file currently loaded, import file first")
        
    Else
        
        If Not CleanUpOptions Is Nothing Then
            
            tempMsg = "Do you want to carry out the following actions?" & vbCrLf
            With CleanUpOptions
                tempMsg = tempMsg & "Set the minimum absorbance to zero (prevent negative absorbance)"
                If .Item(conNormVol).Item(1) = True Then
                    tempMsg = tempMsg & "Volume normalization" & vbCrLf
                End If
                If .Item(conTrunc).Item(1) = True Then
                    tempMsg = tempMsg & "Data truncation" & vbCrLf
                End If
                If .Item(conThin).Item(1) = True Then
                    tempMsg = tempMsg & "Data thinning" & vbCrLf
                End If
                If .Item(conAlign).Item(1) = True Then
                    tempMsg = tempMsg & "Data alignment to zero" & vbCrLf
                End If
                If .Item(conNormUV).Item(1) = True Then
                    tempMsg = tempMsg & "Normalization/scaling of UV signal" & vbCrLf
                End If
            End With
                    
            answer = MsgBox(tempMsg, vbYesNo + vbQuestion, "Confirm actions")
            
            If answer = vbYes Then
                    
                'reset the object to default state, so that clicking cleanup multiple times
                'works as intended
                Call SEC.RevertAllToOriginal
                'Call SEC.TempCleanUp(SelectedInjection)
                
                With CleanUpOptions.Item(conAutozeroUV)
                    If .Item(1) = True Then
                        Call SEC.PostHocUVOffset(.Item(2))
                    End If
                End With
                
                'VOLUME NORMALIZATION
                '1: [do I normalize] 2: [which injection] 3: [which volume]
                With CleanUpOptions.Item(conNormVol)
                    If .Item(1) = True Then
                        If .Item(2) <> "" Then
                            Call SEC.NormalizeVolumeToInjectionNumber(.Item(2))
                        Else
                            Call SEC.ShiftDataX(-1 * .Item(3))
                        End If
                    End If
                End With
                
                'DATA TRUNCATION
                '1: [do I trunc] 2: [start volume] 3: [end volume]
                With CleanUpOptions.Item(conTrunc)
                    If .Item(1) = True Then
                        Call SEC.TruncateToVolumeRange(.Item(2), .Item(3))
                    End If
                End With
                
                'DATA THINNING
                '1: [do I thin] 2: [distance between points]
                With CleanUpOptions.Item(conThin)
                    If .Item(1) = True Then
                        Call SEC.ThinData(.Item(2))
                    End If
                End With
                
                'DATA ALIGNMENT
                '1: [do I align]
                With CleanUpOptions.Item(conAlign)
                    If .Item(1) = True Then
                        Call SEC.AlignDataStartToZero
                    End If
                End With
                
                'UV NORMALIZATION
                '1: [do I norm] 2: [INTEGRAL/MAXVALUE] 3: [startvolume] 4: [endvolume]
                With CleanUpOptions.Item(conNormUV)
                    If .Item(1) = True Then
                        Select Case .Item(2)
                            Case "MAXVALUE"
                                Call SEC.NormalizeToMaxValueInRange(.Item(3), .Item(4), 1)
                            Case "INTEGRAL"
                                Call SEC.NormalizeToIntegralInRange(.Item(3), .Item(4), 1)
                        End Select
                    End If
                End With
                
                'update the form's controls
                Call UpdateMinMax
                
                Debug.Print ("##Data successfully cleaned up.")
                
            Else
            
                Call MsgBox("Clean up cancelled!", vbOKOnly + vbInformation, "Info")
                
            End If
            
        Else
        
            MsgBox ("No settings detected")
            
        End If
        
    End If

End Sub

Private Function GetTickedCollection() As Collection

    Dim tempColl As Collection
    
    Set tempColl = New Collection

    If tickCurveUV1.Value = True Then Call tempColl.Add(NamesArray(0), NamesArray(0))
    If tickCurveUV2.Value = True Then Call tempColl.Add(NamesArray(1), NamesArray(1))
    If tickCurveUV3.Value = True Then Call tempColl.Add(NamesArray(2), NamesArray(2))
    
    If tickCurvePercentB.Value = True Then Call tempColl.Add(NamesArray(3), NamesArray(3))
    
    If tickCurveConductivity.Value = True Then Call tempColl.Add(NamesArray(4), NamesArray(4))
    
    Set GetTickedCollection = tempColl
    
    Set tempColl = Nothing
    
End Function

Private Sub FormatAxisTitle(ByRef TargetAxis As Excel.Axis)

    If TargetAxis.AxisTitle.Text Like "A[0-9][0-9][0-9]*" Then
        TargetAxis.AxisTitle.Characters(1, 1).Font.Italic = True
        TargetAxis.AxisTitle.Characters(2, 3).Font.Subscript = True
    End If
    
   'add other cases here

End Sub

Private Sub ctrlDrawButton_Click()
    
    Dim GraphData As clsGraphData
    Dim SheetName As String
    Dim GraphSheet As Excel.Worksheet
    Dim RowsN As Long
    
    Dim TickedCollection As Collection
    Dim NumberOfCurves As Long
    
    Dim TargetChart As Excel.ChartObject
    Dim CurrentSeries As Excel.Series
    Dim tempAxis As Excel.Axis
            
    Dim IncludeSecondaryAxis As Boolean
    
    Dim XRange As Excel.Range
    Dim YRange As Excel.Range
    
    Dim tempOutArray() As Double
    
    Dim CurveID As String
    
    Dim i As Long, j As Long
            
    'check if file has already been imported
    If Not DataIsImported Then
    
        MsgBox ("No file currently loaded, import file first")
        
    Else
    
        'check which curves we are supposed to draw
        Set TickedCollection = GetTickedCollection
        NumberOfCurves = TickedCollection.Count
        
        SheetName = CreateSheetFromName("Unicorn")
        Set GraphSheet = ActiveWorkbook.Worksheets.Item(SheetName)
        
        'if there is anything assigned to be drawn on the secondary axis
        For i = 0 To conAllowedCurveTypesNumber - 1
            If (CurveTickCollection.Item(NamesArray(i)).Value = True) And _
                (AxisAssignment.Item(NamesArray(i)) = xlSecondary) Then
                    IncludeSecondaryAxis = True
            End If
        Next i
            
        'If (tickCurveUV1.Value = True) And (AxisAssignment.Item("UV1") = xlSecondary) Or _
        '    (tickCurveUV2.Value = True) And (AxisAssignment.Item("UV2") = xlSecondary) Or _
        '    (tickCurveUV3.Value = True) And (AxisAssignment.Item("UV3") = xlSecondary) Or _
        '    (tickCurvePercentB.Value = True) And (AxisAssignment.Item("CONC") = xlSecondary) Or _
        '    (tickCurveConductivity.Value = True) And (AxisAssignment.Item("COND") = xlSecondary) _
        'Then IncludeSecondaryAxis = True
        
        Set TargetChart = Chromatography_AddChart(IncludeSecondaryAxis)
        
        For i = 1 To NumberOfCurves
                
            CurveID = TickedCollection.Item(i)
                
            Set GraphData = SEC.TempGetGraph(CurveID)
            RowsN = UBound(GraphData.XDataArray) - LBound(GraphData.XDataArray, 1) + 1
            
            'temporary output to sheet (not neccessary...)
                GraphSheet.Range("A1").Offset(0, 2 * (i - 1)).Value = CurveID
                
                ReDim tempOutArray(1 To RowsN, 1 To 2)
                
                For j = 1 To RowsN
                    tempOutArray(j, 1) = GraphData.XData(j)
                    tempOutArray(j, 2) = GraphData.YData(j)
                Next j
                
                GraphSheet.Range("A2").Resize(RowsN, 2).Offset(0, 2 * (i - 1)).Value = tempOutArray
                            
                'MsgBox ("Graph data exported to " & SheetName & "!")
                        
            '\temp
                
            'add the current series
            Set CurrentSeries = ChartAddSeriesToChartDirect(GraphData.XDataArray, _
                                                            GraphData.YDataArray, _
                                                            TargetChart, _
                                                            CurveID, _
                                                            WhichAxisAmI(CurveID), _
                                                            GetDefaultSettings(CurveID))
            
            'label the axes
            With CurrentSeries
            
                .AxisGroup = WhichAxisAmI(CurveID)
                
                If .AxisGroup = xlPrimary Then
                    'Volume [mL]
                    If GraphData.XAxisInfo.Label <> "" Then
                        Set tempAxis = TargetChart.Chart.Axes(xlCategory, .AxisGroup)
                        tempAxis.AxisTitle.Text = GraphData.XAxisLabel & " [" & GraphData.XAxisUnit & "]"
                        Call FormatAxisTitle(tempAxis)
                    End If
                End If
                
                'UV280 [mAU]
                If GraphData.YAxisInfo.Label <> "" Then
                    Set tempAxis = TargetChart.Chart.Axes(xlValue, .AxisGroup)
                    tempAxis.HasTitle = True 'hacky
                    tempAxis.AxisTitle.Text = GraphData.YAxisLabel & " [" & GraphData.YAxisUnit & "]"
                    Call FormatAxisTitle(tempAxis)
                End If
                
            End With
            
            'Debugging
            
            
        Next i
        
        'TEMP JURY RIGGED
        With TargetChart.Chart
            With .Axes(xlCategory, xlPrimary)
                'HACKYYY, should import separate settings' profiles for each axis
                    .MajorTickMark = xlTickMarkCross
                .MinimumScale = XAxisMin
                .MaximumScale = XAxisMax
                Select Case XAxisMax - XAxisMin
                    Case Is > 1000: .MajorUnit = 200
                    Case Is > 500: .MajorUnit = 100
                    Case Is > 200: .MajorUnit = 50
                    Case Is > 100: .MajorUnit = 25
                    Case Else: .MajorUnit = 10
                End Select
                .MinorUnit = .MajorUnit / 5
            End With
            
            With .Axes(xlValue, xlPrimary)
                .MinimumScale = YAxisMin
                .MaximumScale = YAxisMax
                Select Case YAxisMax - YAxisMin
                    Case Is > 10000: .MajorUnit = 5000
                    Case Is > 5000: .MajorUnit = 2500
                    Case Is > 1000: .MajorUnit = 500
                    Case Is > 500: .MajorUnit = 250
                    Case Is > 100: .MajorUnit = 50
                    Case Is > 50: .MajorUnit = 25
                    Case Is > 10: .MajorUnit = 5
                    Case Is > 5: .MajorUnit = 2
                    Case Else: .MajorUnit = 1
                End Select
                .MinorUnit = .MajorUnit / 2
            End With
            
            If .HasAxis(xlValue, xlSecondary) Then
                With .Axes(xlValue, xlSecondary)
                    .MinimumScale = 0
                    .MaximumScale = 100
                    .MinorUnit = .MajorUnit / 2
                End With
            End If
        
        .HasTitle = True
        .ChartTitle.Text = SEC.ExperimentName & " (" & SEC.ExperimentDate & ")"
        
        End With
            
            
    End If
    
    
    'Draw annotated regions!!!!!
    Dim TempX(1 To 4) As Double
    Dim TempY(1 To 4) As Double
    
    Dim TempName As String
    
    If Not (SEC.LabeledRegions Is Nothing) Then
        If SEC.LabeledRegions.Count > 0 Then
            For i = 1 To SEC.LabeledRegions.Count
            
                TempName = SEC.LabeledRegions.Annotation(i)
            
                TempX(1) = SEC.LabeledRegions.Xstart(i)
                TempX(2) = TempX(1)
                TempX(3) = SEC.LabeledRegions.Xend(i)
                TempX(4) = TempX(3)
                
                With TargetChart.Chart.Axes(xlValue, xlPrimary)
                    TempY(1) = .MaximumScale
                    TempY(2) = .MinimumScale
                    TempY(3) = .MinimumScale
                    TempY(4) = .MaximumScale
                End With
                
                Set CurrentSeries = ChartAddSeriesToChartDirect(TempX, _
                                                                    TempY, _
                                                                    TargetChart, _
                                                                    TempName, _
                                                                    xlPrimary)
                
                With CurrentSeries
                    .Smooth = False
                    .Format.Line.Weight = 0.1
                End With
                
            Next i
        End If
    End If
                

    
    Dim tempDelColl As VBA.Collection
    Set tempDelColl = New VBA.Collection
    
    'TODO: deleting only the "helper series". It is frustratingly hard. So for now I kill the legend
    With TargetChart.Chart
        If .HasLegend Then
            .Legend.Delete
        
       '
        '    For i = 1 To .FullSeriesCollection.Count
         '       Set CurrentSeries = .FullSeriesCollection.Item(i)
          '      If Not IsElementOf(CurrentSeries.Name, TickedCollection) Then
           '         If CurrentSeries.Format.Line.Visible = msoTrue Then
            '            CurrentSeries.Format.Line.Visible = msoFalse
             '           CurrentSeries.Name = vbNullString
              '      End If
                
   '                 tempDelColl.Add i
              '  End If
            'Next i
       
  '          For i = 0 To tempDelColl.Count - 1
  '              .Legend.LegendEntries(tempDelColl(tempDelColl.Count - i)).Delete
   '         Next i
            
        End If
    End With
        
    
    'MsgBox ("Chart placed in " & SheetName & "!")
        
    'cleanup
    Set TickedCollection = Nothing
    Set GraphData = Nothing
    Set tempAxis = Nothing
    Set GraphSheet = Nothing
    Set TargetChart = Nothing
    Set CurrentSeries = Nothing
    
End Sub

Private Function GetCurveID(ByVal TabIndex As Long) As String

    GetCurveID = NamesArray(TabIndex)

End Function


Private Function WhichAxisAmI(ByVal CurveString As String) As XlAxisGroup

    WhichAxisAmI = AxisAssignment.Item(CurveString)

End Function

Private Sub UpdateMinMax()
    
    Dim i As Long
    Dim CurveID As String
    
    If Not (SEC Is Nothing) Then
    
        i = 0
        CurveID = NamesArray(i)
        Do While ((CurveTickCollection.Item(CurveID) = False) Or _
                    (AxisAssignment.Item(CurveID) = xlSecondary)) And (i < conAllowedCurveTypesNumber - 1)
            i = i + 1
            CurveID = NamesArray(i)
        Loop
        
        If AxisAssignment(CurveID) = xlPrimary Then
            
            XAutoMin = SEC.Chromatograms.Item(CurveID).Xmin
            XAutoMax = SEC.Chromatograms.Item(CurveID).Xmax
            
            YAutoMin = 0
            YAutoMax = Int(SEC.Chromatograms.Item(CurveID).Ymax * 1.1 / 10 + 1) * 10
            '(=increase by 10 % for comfortable overhead, then round to nearest 10)
            
        End If
        
    End If
    
    XAxisMin = XAutoMin
    XAxisMax = XAutoMax
    
    YAxisMin = YAutoMin
    YAxisMax = YAutoMax
    
    ctrlXAxisMinScroll.Value = XAutoMin
    ctrlXAxisMaxScroll.Value = XAutoMax
    
    ctrlYAxisMinScroll.Value = YAutoMin
    ctrlYAxisMaxScroll.Value = YAutoMax
        

End Sub

Private Sub ctrlImportButton_Click()
    
    Dim tempFrm As frmCleanUpSettings
    
    Debug.Print ("Creating Chromatography object...")
    Set SEC = New clsGeneralizedChromatography
    
    InputFileValid = True
    
    Debug.Print ("Calling import function")
    Call SEC.ImportFile(FileName, FileType)
    
    If InputFileValid Then
    
        Debug.Print ("Updating form fields...")
        
        'update the sliders for plotting
            Call UpdateMinMax
        
        'set the initial selected injection
            If Not SEC.Injections Is Nothing Then
                SelectedInjection = SEC.Injections.Count
            Else
                SelectedInjection = 0
            End If
        
        RaiseEvent EImportCmdDone(0)
        
        'enable control buttons
            Debug.Print ("Activating form controls...")
            Call EnableDataProcessing
            
        'set default cleanup settings
            Debug.Print ("Defining default settings...")
            Set tempFrm = New frmCleanUpSettings
            Set CleanUpOptions = New VBA.Collection
            With tempFrm
                Set .SEC = SEC
                Set .defaultCleanUpOptions = CleanUpOptions
                .ManualInitialize
            End With
            Unload tempFrm
            
    End If
    
    Set tempFrm = Nothing
    
End Sub

Private Sub EnableDataProcessing()

    DataIsImported = True

    With ctrlCleanUpButton
        .Enabled = True
        .ControlTipText = "Must import data first."
    End With
    
    With ctrlDrawButton
        .Enabled = True
        .ControlTipText = "Draws the chromatogram."
    End With
    
    With ctrlSelectInjection
        .Enabled = True
        .ControlTipText = "Select the reference injection for volume normalization."
    End With
    
    With ctrlSelectFractions
        .Enabled = True
        .ControlTipText = "Select the fractions to be highlighted."
    End With
    
    With ctrlCleanUpSettings
        .Enabled = True
        .ControlTipText = "Adjust settings for data truncation, normalization, etc."
    End With
    
    With ctrlCleanUpButton
        .Enabled = True
        .ControlTipText = "Normalize and adjust data to prepare it for graph output."
    End With
    
    With ctrlReviseMetadata
        .Enabled = True
        .ControlTipText = "Revise the metadata (Column Volume, Experiment Name, ...)."
    End With
    
    frameCurveSettings.Visible = True
    frameXAxisSettings.Visible = True
    frameYAxisSettings.Visible = True
    frameCurveSelect.Visible = True

End Sub

Private Function FileTypeString(ByVal InputString As String) As String
'when updating this, also check initialization where the combobox is populated!

Select Case InputString
    Case conUni3_title
        FileTypeString = conUni3_ID
    Case conUni6_title
        FileTypeString = conUni6_ID
    Case Else
        Call Err.Raise(1001, , "unsupported file type")
        
End Select

End Function

Private Sub ctrlFileSelection_Click()
'picking the file to import using Windows native file selection dialog

    Dim conFileDialog As FileDialog
    Dim tempFile As String

    Set conFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With conFileDialog
    
        .AllowMultiSelect = False
        
        .InitialFileName = DefaultFileName
                
        .Show
        
        If .SelectedItems.Count > 0 Then
        
            tempFile = .SelectedItems.Item(1)
            
            'MsgBox ("Selected item: " & .SelectedItems.Item(1))
            
            ctrlSelFileTxtBox.Text = tempFile
            FileName = tempFile
        
        Else
        
            'ctrlSelFileTxtBox.Text = vbEmpty
            'FileName = vbEmpty
            
        End If
                
        
    End With

End Sub


Public Sub ManualInitialize()

    Call UserForm_Initialize

End Sub

Private Sub ctrlReviseMetadata_Click()
    
    Dim tempForm As frmMetadata
    
    If DataIsImported Then
        
        Set tempForm = New frmMetadata
        
        With tempForm
            Set .SEC = SEC
            Call .ManualInitialize
            .Show (vbModal)
            Debug.Print "Metadata imported"
            'Call MsgBox("Metadata imported", vbOKOnly + vbInformation)
        End With
        
        Unload tempForm
        
    Else
    
        Call Err.Raise(1001, , "No data imported")
        
    End If
        
End Sub

Private Sub ctrlSelectFractions_Click()

    Dim tempInj As frmPickFractions
    
    If DataIsImported Then
        
        Set tempInj = New frmPickFractions
        
        With tempInj
            'import SEC object
                Set .SEC = SEC
                .ManualInitialize
                .Show (vbModal)
                'Set SEC.LabeledRegions = .AnnotationObject
        End With
                
        'Call MsgBox("Form closing.", _
        '        vbInformation + vbOKOnly)
                
        'unload it from memory
        Unload tempInj
        
    Else
    
        Call Err.Raise(1001, , "No data imported")
        
    End If

End Sub

Private Sub ctrlSelectInjection_Click()

    Dim tempInj As frmPickInjection
    Dim tempMsg As String
    
    If DataIsImported Then
        
        Set tempInj = New frmPickInjection
        
        With tempInj
            'kind of HACKY, to return to this form
                Set .ParentObject = Me
            'import SEC object
                Set .SEC = SEC
            'import default starting value
                .SelectedInjection = SelectedInjection
            'initialize subform
                Call .ManualInitialize
                Call .RefreshLayout
                '.Show (vbModeless)
                .Show (vbModal)
            'get the choice
                SelectedInjection = .SelectedInjection
        End With
                
        tempMsg = "Default injection set to Injection #" & SelectedInjection & _
                " (" & SEC.Injections.XData(SelectedInjection) & " mL)"
                
        'Call MsgBox(tempMsg, vbInformation + vbOKOnly)
        Debug.Print tempMsg
                
        'unload it from memory
        Unload tempInj
        
    Else
    
        Call Err.Raise(1001, , "No data imported")
        
    End If
    
End Sub

Private Sub ctrlSelFileTxtBox_Change()
    
    FileName = ctrlSelFileTxtBox.Text
    
    Select Case UCase(FileSystem_GetExtension(FileName))
        Case "RES"
            ctrlFileType.Value = conUni3_title
        Case "ZIP"
            ctrlFileType.Value = conUni6_title
    End Select
    
End Sub

Private Sub ctrlXAxisMax_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim tempValue As Double
    Dim OldValue As Variant
        
    OldValue = ctrlXAxisMax.Value
        
    If IsNumeric(OldValue) Then
    
        tempValue = CDbl(OldValue)
        If tempValue >= AbsoluteMinAllowedVolume And tempValue <= AbsoluteMaxAllowedVolume Then
            XAxisMax = tempValue
            ctrlXAxisMaxScroll.Value = tempValue
        Else
            Cancel = True
        End If
    
    Else
        Cancel = True
    End If
    
    If Cancel = True Then MsgBox ("Only numeric entries between " & AbsoluteMinAllowedVolume & " and " & AbsoluteMaxAllowedVolume & " are allowed.")
    
    If XAxisMax < XAxisMin Then
        ctrlXAxisMin.Value = XAxisMax
        ctrlXAxisMinScroll.Value = XAxisMax
    End If
    
End Sub

Private Sub ctrlXAxisMaxScroll_Change()
    ctrlXAxisMax.Value = ctrlXAxisMaxScroll.Value
    XAxisMax = ctrlXAxisMaxScroll.Value
    If XAxisMax < XAxisMin Then
        ctrlXAxisMin.Value = XAxisMax
        ctrlXAxisMinScroll.Value = XAxisMax
    End If
End Sub

Private Sub ctrlXAxisMin_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim tempValue As Double
    Dim OldValue As Variant
        
    OldValue = ctrlXAxisMin.Value
        
    If IsNumeric(OldValue) Then
    
        tempValue = CDbl(OldValue)
        If tempValue >= AbsoluteMinAllowedVolume And tempValue <= AbsoluteMaxAllowedVolume Then
            XAxisMin = tempValue
            ctrlXAxisMinScroll.Value = tempValue
        Else
            Cancel = True
        End If
    
    Else
        Cancel = True
    End If
    
    If Cancel = True Then MsgBox ("Only numeric entries between " & AbsoluteMinAllowedVolume & " and " & AbsoluteMaxAllowedVolume & " are allowed.")
    
    If XAxisMin > XAxisMax Then
        ctrlXAxisMax.Value = XAxisMin
        ctrlXAxisMaxScroll.Value = XAxisMin
    End If
    
End Sub

Private Sub ctrlXAxisMinScroll_Change()
    ctrlXAxisMin.Value = ctrlXAxisMinScroll.Value
    XAxisMin = ctrlXAxisMinScroll.Value
    If XAxisMin > XAxisMax Then
        ctrlXAxisMax.Value = XAxisMin
        ctrlXAxisMaxScroll.Value = XAxisMin
    End If
End Sub

Private Sub ctrlYAxisMax_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim tempValue As Double
    Dim OldValue As Variant
        
    OldValue = ctrlYAxisMax.Value
        
    If IsNumeric(OldValue) Then
    
        tempValue = CDbl(OldValue)
        If tempValue >= AbsoluteMinAllowedUV And tempValue <= AbsoluteMaxAllowedUV Then
            YAxisMax = tempValue
            ctrlYAxisMaxScroll.Value = tempValue
        Else
            Cancel = True
        End If
    
    Else
        Cancel = True
    End If
    
    If Cancel = True Then MsgBox ("Only numeric entries between " & AbsoluteMinAllowedUV & " and " & AbsoluteMaxAllowedUV & " are allowed.")
    If YAxisMax < YAxisMin Then
        ctrlYAxisMin.Value = YAxisMax
        ctrlYAxisMinScroll.Value = YAxisMax
    End If
    
End Sub

Private Sub ctrlYAxisMaxScroll_Change()
    ctrlYAxisMax.Value = ctrlYAxisMaxScroll.Value
    YAxisMax = ctrlYAxisMaxScroll.Value
    If YAxisMax < YAxisMin Then
        ctrlYAxisMin.Value = YAxisMax
        ctrlYAxisMinScroll.Value = YAxisMax
    End If
End Sub

Private Sub ctrlYAxisMin_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim tempValue As Double
    Dim OldValue As Variant
        
    OldValue = ctrlYAxisMin.Value
        
    If IsNumeric(OldValue) Then
    
        tempValue = CDbl(OldValue)
        If tempValue >= AbsoluteMinAllowedUV And tempValue <= AbsoluteMaxAllowedUV Then
            YAxisMin = tempValue
            ctrlYAxisMinScroll.Value = tempValue
        Else
            Cancel = True
        End If
    
    Else
        Cancel = True
    End If
    
    If Cancel = True Then MsgBox ("Only numeric entries between " & AbsoluteMinAllowedUV & " and " & AbsoluteMaxAllowedUV & " are allowed.")
    
    If YAxisMin > YAxisMax Then
        ctrlYAxisMax.Value = YAxisMin
        ctrlYAxisMaxScroll.Value = YAxisMin
    End If
    
End Sub

Private Sub ctrlYAxisMinScroll_Change()
    ctrlYAxisMin.Value = ctrlYAxisMinScroll.Value
    YAxisMin = ctrlYAxisMinScroll.Value
    If YAxisMin > YAxisMax Then
        ctrlYAxisMax.Value = YAxisMin
        ctrlYAxisMaxScroll.Value = YAxisMin
    End If
End Sub



Private Sub frameActions_Click()

End Sub

Private Sub imgIMP_Click()
    EEC = EEC + 1
    If EEC = 5 Then
        EEC = 0
        frmHmm.Show
    End If
End Sub

Private Sub lblAbout1_Click()

End Sub

Private Sub optbutAxisPrimary_Click()
    
    Dim CurrentCurve As Long
    
    CurrentCurve = tabCurveSettings.SelectedItem.Index
    AxisAssignment.Remove NamesArray(CurrentCurve)
    AxisAssignment.Add xlPrimary, NamesArray(CurrentCurve)
    
End Sub

Private Sub optbutAxisSecondary_Click()
    
    Dim CurrentCurve As Long
    
    CurrentCurve = tabCurveSettings.SelectedItem.Index
    AxisAssignment.Remove NamesArray(CurrentCurve)
    AxisAssignment.Add xlSecondary, NamesArray(CurrentCurve)
    
End Sub

Private Sub optbutColorAuto_Click()
    ColorChoiceSetting(tabCurveSettings.SelectedItem.Index) = "AUTO"
End Sub

Private Sub optbutColorCustom_Click()
    ColorChoiceSetting(tabCurveSettings.SelectedItem.Index) = "CUSTOM"
End Sub

Private Sub SEC_InvalidInput()
    InputFileValid = False
    MsgBox ("Invalid input")
End Sub

Private Sub tabCurveSettings_Change()
    Call TabCurveSettingsSwitch
End Sub

Private Sub tickCurveConductivity_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    If tickCurveConductivity.Value = False And NumberOfSelectedCurves = 0 Then
        Cancel = True
        tickCurveConductivity.Value = True
        Call MsgBox("Cannot unselect all curves", vbExclamation + vbOKOnly, "Unable to perform action")
    Else
        Call RefreshTabCurveSettings
    End If
    
End Sub

Private Sub tickCurvePercentB_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    If tickCurvePercentB.Value = False And NumberOfSelectedCurves = 0 Then
        Cancel = True
        tickCurvePercentB.Value = True
        Call MsgBox("Cannot unselect all curves", vbExclamation + vbOKOnly, "Unable to perform action")
    Else
        Call RefreshTabCurveSettings
    End If
    
End Sub

Private Sub tickCurveUV1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    If tickCurveUV1.Value = False And NumberOfSelectedCurves = 0 Then
        Cancel = True
        tickCurveUV1.Value = True
        Call MsgBox("Cannot unselect all curves", vbExclamation + vbOKOnly, "Unable to perform action")
    Else
        Call RefreshTabCurveSettings
    End If
    
End Sub

Private Sub tickCurveUV2_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    If tickCurveUV2.Value = False And NumberOfSelectedCurves = 0 Then
        Cancel = True
        tickCurveUV2.Value = True
        Call MsgBox("Cannot unselect all curves", vbExclamation + vbOKOnly, "Unable to perform action")
    Else
        Call RefreshTabCurveSettings
    End If
    
End Sub

Private Sub tickCurveUV3_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    If tickCurveUV#.Value = False And NumberOfSelectedCurves = 0 Then
        Cancel = True
        tickCurveUV3.Value = True
        Call MsgBox("Cannot unselect all curves", vbExclamation + vbOKOnly, "Unable to perform action")
    Else
        Call RefreshTabCurveSettings
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim i As Long
    
    Debug.Print ("=====Chromatograms Form=====")
    
    ReDim NamesArray(0 To 4)
        NamesArray(0) = "UV1"
        NamesArray(1) = "UV2"
        NamesArray(2) = "UV3"
        NamesArray(3) = "CONC"
        NamesArray(4) = "COND"
    
    
    ReDim SupportedFileTypes(0 To 1)
    
    
    'Populate supported file types list
    'when updating this, also check the function which translates these titles!
    '(function name at least used to be FileTypeString)
    'sorry for poor WET coding, this was my first attempt at this form stuff
    'and I'm lazy to rewrite it (2016-06-10)
        'DataTypes:
        'UNICORN3:
            SupportedFileTypes(0) = conUni3_title
        'UNICORN6:
            SupportedFileTypes(1) = conUni6_title
        ctrlFileType.List = SupportedFileTypes
        
        'Initialize default FileType
        ctrlFileType.Value = SupportedFileTypes(0)

        
    'gather the curve ticks into a collection
        Set CurveTickCollection = New Collection
        
        With CurveTickCollection
            .Add tickCurveUV1, NamesArray(0)
            .Add tickCurveUV2, NamesArray(1)
            .Add tickCurveUV3, NamesArray(2)
            .Add tickCurvePercentB, NamesArray(3)
            .Add tickCurveConductivity, NamesArray(4)
        End With
    
    
    'Set default filename
        FileName = DefaultFileName
        ctrlSelFileTxtBox.Text = FileName
    
        
    'Initialize page names / visibility (settings for different curves)
        tickCurveUV1.Value = True
        tickCurveUV2.Value = False
        tickCurveUV3.Value = False
        tickCurveConductivity.Value = False
        tickCurvePercentB.Value = False
    
    'Initialize Curve Options
        For i = 0 To conAllowedCurveTypesNumber - 1
            ColorChoiceSetting(i) = "AUTO"
        Next i
        
        Set AxisAssignment = New Collection
        'UV:
            AxisAssignment.Add xlPrimary, NamesArray(0)
            AxisAssignment.Add xlPrimary, NamesArray(1)
            AxisAssignment.Add xlPrimary, NamesArray(2)
        'Conc, Cond:
            AxisAssignment.Add xlSecondary, NamesArray(3)
            AxisAssignment.Add xlSecondary, NamesArray(4)
    
    'set min and max values for plotting parameters
        AbsoluteMinAllowedVolume = -100
        AbsoluteMaxAllowedVolume = 1000
        
        AbsoluteMinAllowedUV = -500
        AbsoluteMaxAllowedUV = 10000
        
        XAutoMin = 0
        XAutoMax = 0
        
        YAutoMin = 0
        YAutoMax = 0
        
        With ctrlXAxisMinScroll
            .Min = AbsoluteMinAllowedVolume
            .Max = AbsoluteMaxAllowedVolume
        End With
        
        With ctrlXAxisMaxScroll
            .Min = AbsoluteMinAllowedVolume
            .Max = AbsoluteMaxAllowedVolume
        End With
        
        With ctrlYAxisMinScroll
            .Min = AbsoluteMinAllowedUV
            .Max = AbsoluteMaxAllowedUV
        End With
        
        With ctrlYAxisMaxScroll
            .Min = AbsoluteMinAllowedUV
            .Max = AbsoluteMaxAllowedUV
        End With
        
        Call UpdateMinMax
                
    'Update visuals
        Call RefreshTabCurveSettings
        Call TabCurveSettingsSwitch
    
    'Set the initial state of action buttons
        With ctrlCleanUpButton
            .Enabled = False
            .ControlTipText = "Must import data first."
        End With
        
        With ctrlDrawButton
            .Enabled = False
            .ControlTipText = "Must import data first."
        End With
        
        With ctrlSelectInjection
            .Enabled = False
            .ControlTipText = "Must import data first."
        End With
        
        With ctrlSelectFractions
            .Enabled = False
            .ControlTipText = "Must import data first."
        End With
        
        With ctrlCleanUpSettings
            .Enabled = False
            .ControlTipText = "Must import data first."
        End With
        
        With ctrlReviseMetadata
            .Enabled = False
            .ControlTipText = "Must import data first."
        End With
        
        frameCurveSettings.Visible = False
        frameXAxisSettings.Visible = False
        frameYAxisSettings.Visible = False
        frameCurveSelect.Visible = False
        
        'Cleanup settings
        Set CleanUpOptions = New VBA.Collection

End Sub

Private Function NumberOfSelectedCurves() As Long

    Dim i As Long
    Dim tempcounter As Long
    
    tempcounter = 0
    
    For i = 1 To CurveTickCollection.Count
        If CurveTickCollection.Item(i).Value = True Then
            tempcounter = tempcounter + 1
        End If
    Next i
    
    NumberOfSelectedCurves = tempcounter

End Function

Private Sub TabCurveSettingsSwitch()
    
    Dim tempInd As Long
    
    If NumberOfSelectedCurves = 0 Then
        
        frameCurveSettings.Enabled = False
        
    Else
        If frameCurveSettings.Enabled = False Then
            frameCurveSettings.Enabled = True
            'TODO: this doesn't work, and will give an error
            tabCurveSettings.Value = 0
        End If
    
        If Not (tabCurveSettings.SelectedItem Is Nothing) Then
            
            tempInd = tabCurveSettings.SelectedItem.Index
        
            If AxisAssignment.Item(NamesArray(tempInd)) = xlPrimary Then
                optbutAxisPrimary.Value = True
            Else
                optbutAxisSecondary.Value = True
            End If
            
            If UCase(ColorChoiceSetting(tempInd)) = "AUTO" Then
                optbutColorAuto.Value = True
            Else
                optbutColorCustom.Value = True
            End If
        
        End If
        
    End If

End Sub

Private Sub RefreshTabCurveSettings()
    
    Call TabCurveSettingsSwitch
    
    With tabCurveSettings.Tabs.Item(0)
        .Caption = tickCurveUV1.Caption
        .Visible = tickCurveUV1.Value
    End With
    
    With tabCurveSettings.Tabs.Item(0)
        .Caption = tickCurveUV1.Caption
        .Visible = tickCurveUV1.Value
    End With
    
    With tabCurveSettings.Tabs.Item(1)
        .Caption = tickCurveUV2.Caption
        .Visible = tickCurveUV2.Value
    End With
    
    With tabCurveSettings.Tabs.Item(2)
        .Caption = tickCurveUV3.Caption
        .Visible = tickCurveUV3.Value
    End With
    
    With tabCurveSettings.Tabs.Item(3)
        .Caption = tickCurvePercentB.Caption
        .Visible = tickCurvePercentB.Value
    End With
        
    With tabCurveSettings.Tabs.Item(4)
        .Caption = tickCurveConductivity.Caption
        .Visible = tickCurveConductivity.Value
    End With
        

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Dim answer As Long
    
    'if pressing X
    If CloseMode = 0 Then
        
        answer = MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Exit?")
        
        If answer = vbYes Then
            'MsgBox ("Bye!")
            Debug.Print "Form terminated normally."
            'Unload Me
        Else
            Call MsgBox("Exit form cancelled!", vbInformation + vbOKOnly)
            Cancel = 1
        End If
        
    End If
    
End Sub

Private Sub UserForm_Terminate()
    Set SEC = Nothing
    Set AxisAssignment = Nothing
    Set CleanUpOptions = Nothing
    Set CurveTickCollection = Nothing
    Unload Me
End Sub
