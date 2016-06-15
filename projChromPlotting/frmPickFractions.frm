VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPickFractions 
   Caption         =   "Fractions"
   ClientHeight    =   2856
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   10548
   OleObjectBlob   =   "frmPickFractions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPickFractions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2016-05-21,
'Last update 2016-05-27
'2016-06-12 import annotated ranges when loaded (so you can open and close it multiple times without
'           losing the visibility of annotations
'====================================================================================================



Option Explicit


Public SEC As clsGeneralizedChromatography

Private IAmInitialized As Boolean
Private DefaultColor As Long
Private SelectedAnnotation As String

Private RangeCollection As Collection
Public AnnotationObject As clsGraphRegions


Public Sub ManualInitialize()

    Dim i As Long
    Dim StartV As Double, EndV As Double
    Dim RangeName As String

    Call UserForm_Initialize
    
    If Not (SEC Is Nothing) Then
        
        'fill out the boxes for choosing fractions and set defaults
        ctrlStartFraction.List = SEC.Fractions.AnnotationArray
        ctrlEndFraction.List = SEC.Fractions.AnnotationArray
        ctrlStartFraction.Value = ctrlStartFraction.List(0)
        ctrlEndFraction.Value = ctrlStartFraction.List(ctrlStartFraction.ListCount - 1)
        IAmInitialized = True
        
        'populate table with regions already annotated:
        If Not SEC.LabeledRegions Is Nothing Then
        
            For i = 1 To SEC.LabeledRegions.Count
                
                'extract data
                With SEC.LabeledRegions
                    StartV = .Xstart(i)
                    EndV = .Xend(i)
                    RangeName = .Annotation(i)
                End With
                
                'add the region
                Call AddRange(StartV, EndV, RangeName)
                
            Next i
            
        End If
        
    End If
    
    

End Sub

Private Sub DISABLED_ColorPicker_Click()
    
    Const AbusedColorIndex As Long = 32
    Dim SavedColor As Long
    Dim ColorDialog As Object
    Dim NewColor As Long
    
    
    'HACKYYYY
    SavedColor = ActiveWorkbook.Colors(AbusedColorIndex)
    
    If Application.Dialogs(xlDialogColorPalette).Show(AbusedColorIndex) = True Then
        
        NewColor = ActiveWorkbook.Colors(AbusedColorIndex)
        ActiveWorkbook.Colors(AbusedColorIndex) = SavedColor
        ColorPicker.BackColor = NewColor
        
    End If
    

End Sub

Private Function GetFractionIndex(ByVal FractionName As String) As Long

    Dim i As Long
    
    For i = 1 To SEC.Fractions.Count
        If SEC.Fractions.Annotation(i) = FractionName Then
            GetFractionIndex = i
            Exit For
        End If
    Next i

End Function

Private Function TruncateStringArray(ByRef SourceArray() As String, _
                                     ByVal StartIndex As Long, _
                                     ByVal EndIndex As Long _
                                     ) As String()
                                     
    Dim i As Long
    
    Dim tempArray() As String
    
    ReDim tempArray(1 To EndIndex - StartIndex + 1)
    
    For i = StartIndex To EndIndex
        tempArray(i - StartIndex + 1) = SourceArray(i)
    Next i
    
    TruncateStringArray = tempArray
                                     
End Function

Private Sub ctrlRemove_Click()

    Dim TempName As String
    Dim i As Long
    
    If Not listRegions.Value = vbNull Then
        
        TempName = listRegions.Value
        
        For i = 0 To listRegions.ListCount - 1
            If listRegions.List(i) = TempName Then
                listRegions.RemoveItem (i)
                Exit For
            End If
        Next i
        
        RangeCollection.Remove (TempName)
        
    End If

End Sub

Private Sub AddRange( _
    ByVal StartVolume As Double, _
    ByVal EndVolume As Double, _
    ByVal Name As String)

    Dim tempRecord As clsFractionAnnotationDescriptor
    Set tempRecord = New clsFractionAnnotationDescriptor
    
    With tempRecord
        .Color = DefaultColor
        '.StartFraction = StartName
        '.EndFraction = EndName
        .StartVolume = StartVolume
        .EndVolume = EndVolume
        .Name = Name
    End With
    
    If Not (IsElementOf(tempRecord.Name, RangeCollection)) Then
        'add the full descriptor to the collection
        RangeCollection.Add tempRecord, tempRecord.Name
        'add entry to the visual table
        listRegions.AddItem tempRecord.Name
    End If
    
    Set tempRecord = Nothing

End Sub

Private Sub ctrlAdd_Click()
    
    Dim StartFraction As String
    Dim EndFraction As String
    
    Dim StartVolume As Double
    Dim EndVolume As Double
    
    Dim Name As String
    
    StartFraction = ctrlStartFraction.Value
    EndFraction = ctrlEndFraction.Value
    
    StartVolume = SEC.Fractions.Xstart(GetFractionIndex(StartFraction))
    EndVolume = SEC.Fractions.Xend(GetFractionIndex(EndFraction))
    
    Name = StartFraction & " - " & EndFraction
    
    Call AddRange(StartVolume, EndVolume, Name)
    
End Sub

Private Sub ctrlEndFraction_Change()
    
    If IAmInitialized = True Then
        ctrlStartFraction.List = TruncateStringArray(SEC.Fractions.AnnotationArray, _
                                    1, _
                                    GetFractionIndex(ctrlEndFraction.Value))
    End If
    
End Sub

Private Sub ctrlStartFraction_Change()
    
    If IAmInitialized = True Then
        ctrlEndFraction.List = TruncateStringArray(SEC.Fractions.AnnotationArray, _
                                GetFractionIndex(ctrlStartFraction.Value), _
                                SEC.Fractions.Count)
    End If

End Sub



Private Sub listRegions_Change()
    
    Dim tempDescriptor As clsFractionAnnotationDescriptor
    
    Set tempDescriptor = RangeCollection.Item(listRegions.Value)
    
    With tempDescriptor
        TextBoxVolumeStart.Text = .StartVolume
        TextBoxVolumeEnd.Text = .EndVolume
        ColorPicker.BackColor = .Color
    End With
    
    Set tempDescriptor = Nothing
    
End Sub

Private Function GenerateAnnotationObject() As clsGraphRegions

    Dim StartEndArray() As Double
    Dim Annotation() As String
    Dim tempcGraphRegions As clsGraphRegions
    Dim tempRecord As clsFractionAnnotationDescriptor
    Dim NumberOfAnnotations As Long
    
    NumberOfAnnotations = RangeCollection.Count
    
    If NumberOfAnnotations > 0 Then
        
        ReDim StartEndArray(1 To NumberOfAnnotations, 1 To 2)
        ReDim Annotation(1 To NumberOfAnnotations)
        
        'extract values
        For i = 1 To NumberOfAnnotations
        
            Set tempRecord = RangeCollection.Item(i)
            With tempRecord
                StartEndArray(i, 1) = .StartVolume
                StartEndArray(i, 2) = .EndVolume
                Annotation(i) = .Name
            End With
            
        Next i
        
        'store values into the object
        Set tempcGraphRegions = New clsGraphRegions
            
        With tempcGraphRegions
            .AnnotationArray = Annotation
            .XStartEndArray = StartEndArray
        End With
        
        Set GenerateAnnotationObject = tempcGraphRegions
        
    End If
        
        Set tempRecord = Nothing
        Set tempcGraphRegions = Nothing

End Function

Private Sub UserForm_Initialize()
    
    DefaultColor = vbRed
    
    ColorPicker.BackColor = DefaultColor
    ColorPicker.ForeColor = vbWhite
    ColorPicker.Caption = String(6 - Len(Hex(DefaultColor)), "0") & Hex(DefaultColor)
    
    If RangeCollection Is Nothing Then
        Set RangeCollection = New Collection
    End If
    
    'listRegions.RemoveItem (0)
    
End Sub




Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

Dim answer As Long
    
    If CloseMode = 0 Then
    
        answer = MsgBox("Selected regions will be imported back to the experiment." & _
                    vbCrLf & "Continue?", _
                    vbQuestion + vbOKCancel, "Confirm data input")
        
        Select Case answer
        
            Case vbCancel
                Cancel = 1
                
            Case vbOK
                Set AnnotationObject = GenerateAnnotationObject()
                Set SEC.LabeledRegions = AnnotationObject
                
        End Select
    End If

End Sub



Private Sub UserForm_Terminate()
    
    Set SEC = Nothing
    Set RangeCollection = Nothing
    Set AnnotationObject = Nothing

End Sub

