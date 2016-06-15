VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPickInjection 
   Caption         =   "Injections"
   ClientHeight    =   5544
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   4548
   OleObjectBlob   =   "frmPickInjection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPickInjection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2016-05-15,
'Last update 2016-05-27
'====================================================================================================



Option Explicit

Private Const conFirstControlIndex As Long = 1
Private Const conLastControlIndex As Long = 30

Public ParentObject As frmCreateChromatograms

Public IHaveDoneMyJob As Boolean

Public SEC As clsGeneralizedChromatography

Private NumberOfInjections As Long

Private ControlCollection As Collection

Public SelectedInjection As Long



Private Sub DISABLED_UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Dim SavedTop As Double
    Dim SavedLeft As Double
        
    If CloseMode = 0 Then
        If Not (ParentObject Is Nothing) Then
            Cancel = True
            
            'ParentObject.ctrlSelectInjection.SetFocus
            
            SavedTop = ParentObject.Top
            SavedLeft = ParentObject.Left
            
            
            Me.Hide
            ParentObject.Show (vbModeless)
            'ParentObject.Top = SavedTop
            'ParentObject.Left = SavedLeft
            IHaveDoneMyJob = True
        End If
    End If
End Sub



Public Sub ManualInitialize()

    Call UserForm_Initialize

End Sub

Private Sub UserForm_Initialize()
    
    Dim i As Long
    Dim a As Object
    
    Set ControlCollection = New Collection
    
    'populate the ControlCollection
    For i = conFirstControlIndex To conLastControlIndex
        Set a = Controls("OptionButton" & i)
        ControlCollection.Add a, CStr(i)
    Next i
    
    
        
    
End Sub

Public Sub RefreshLayout()

    Dim i As Long
    Dim InjectionIndex As Long
    Dim tempOptionBox As Control
    Dim tempFullness As Double

    NumberOfInjections = SEC.Injections.Count
    
    'disable / enable required options
    For i = conFirstControlIndex To NumberOfInjections
    
        InjectionIndex = 1 + i - conFirstControlIndex
        
        With ControlCollection.Item(CStr(i))
            'e.g. OptionButton1.Enabled = True
            .Enabled = True
            .Visible = True
            'e.g. OptionButton1.Caption = 1: 4 mL
            .Caption = i & ": " & SEC.Injections.XData(InjectionIndex) & " mL"
        End With
        
    Next i
    
    For i = NumberOfInjections + 1 To conLastControlIndex
        With ControlCollection.Item(CStr(i))
            'e.g. OptionButton1.Enabled = False
            .Enabled = False
            .Visible = False
        End With
    Next i
    
    'how many of the grid members should be there
    tempFullness = NumberOfInjections / (conLastControlIndex - conFirstControlIndex + 1)
    
    'resize the form to fit exactly the enabled members
    Select Case tempFullness
        Case Is > 0.5
            Me.Width = 240
            Me.Height = 300
        Case Else
            Me.Width = 120
            Me.Height = 30 + 2 * tempFullness * 270
    End Select
    
    'select the default
    Controls("OptionButton" & SelectedInjection).Value = True
    
End Sub

Private Sub OptionButton1_Click(): SelectedInjection = 1: End Sub
Private Sub OptionButton2_Click(): SelectedInjection = 2: End Sub
Private Sub OptionButton3_Click(): SelectedInjection = 3: End Sub
Private Sub OptionButton4_Click(): SelectedInjection = 4: End Sub
Private Sub OptionButton5_Click(): SelectedInjection = 5: End Sub
Private Sub OptionButton6_Click(): SelectedInjection = 6: End Sub
Private Sub OptionButton7_Click(): SelectedInjection = 7: End Sub
Private Sub OptionButton8_Click(): SelectedInjection = 8: End Sub
Private Sub OptionButton9_Click(): SelectedInjection = 9: End Sub
Private Sub OptionButton10_Click(): SelectedInjection = 10: End Sub
Private Sub OptionButton11_Click(): SelectedInjection = 11: End Sub
Private Sub OptionButton12_Click(): SelectedInjection = 12: End Sub
Private Sub OptionButton13_Click(): SelectedInjection = 13: End Sub
Private Sub OptionButton14_Click(): SelectedInjection = 14: End Sub
Private Sub OptionButton15_Click(): SelectedInjection = 15: End Sub
Private Sub OptionButton16_Click(): SelectedInjection = 16: End Sub
Private Sub OptionButton17_Click(): SelectedInjection = 17: End Sub
Private Sub OptionButton18_Click(): SelectedInjection = 18: End Sub
Private Sub OptionButton19_Click(): SelectedInjection = 19: End Sub
Private Sub OptionButton20_Click(): SelectedInjection = 20: End Sub
Private Sub OptionButton21_Click(): SelectedInjection = 21: End Sub
Private Sub OptionButton22_Click(): SelectedInjection = 22: End Sub
Private Sub OptionButton23_Click(): SelectedInjection = 23: End Sub
Private Sub OptionButton24_Click(): SelectedInjection = 24: End Sub
Private Sub OptionButton25_Click(): SelectedInjection = 25: End Sub
Private Sub OptionButton26_Click(): SelectedInjection = 26: End Sub
Private Sub OptionButton27_Click(): SelectedInjection = 27: End Sub
Private Sub OptionButton28_Click(): SelectedInjection = 28: End Sub
Private Sub OptionButton29_Click(): SelectedInjection = 29: End Sub
Private Sub OptionButton30_Click(): SelectedInjection = 30: End Sub



Private Sub UserForm_Terminate()
    
    Set SEC = Nothing
    Set ParentObject = Nothing
    Set ControlCollection = Nothing

End Sub
