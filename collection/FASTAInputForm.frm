VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FASTAInputForm 
   Caption         =   "FASTA Input"
   ClientHeight    =   5400
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9288.001
   OleObjectBlob   =   "FASTAInputForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FASTAInputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub B_OutputCase_Click()

Const CaptionD = "Output UPPERCASE"
Const CaptionE = "Output lowercase"
Const CaptionF = "Preserve case"

Select Case B_OutputCase.Caption
    Case CaptionD: B_OutputCase.Caption = CaptionE
    Case CaptionE: B_OutputCase.Caption = CaptionF
    Case CaptionF: B_OutputCase.Caption = CaptionD
    Case Else: B_OutputCase.Caption = CaptionD
End Select

End Sub

Private Sub B_OutputType_Click()

Const CaptionA = "Output Sequences Only"
Const CaptionB = "Output Sequences and Headers"
Const CaptionC = "Output Headers Only"

Select Case B_OutputType.Caption
    Case CaptionA: B_OutputType.Caption = CaptionB
    Case CaptionB: B_OutputType.Caption = CaptionC
    Case CaptionC: B_OutputType.Caption = CaptionA
    Case Else: B_OutputType.Caption = CaptionA
End Select

End Sub




Private Sub CommandButton1_Enter()
FASTAInputForm.Hide
End Sub

