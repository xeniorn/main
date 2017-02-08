Attribute VB_Name = "XmodProteins"
Option Explicit

'****************************************************************************************************
Private Function AACharge(pKa As Double, pH As Double, Species As Variant) As Double

'====================================================================================================
'Calculates the charge of a particular acidic or basic residue, needed for theoretical pI calculation
'Juraj Ahel, 2015-02-02, for theoretical pI calculation
'Last update 2015-02-02
'====================================================================================================
'Species 0 = acid, 1 = base

Dim ChargeSign As Long
Dim Charge As Double

If Species = 0 Then ChargeSign = -1 Else ChargeSign = 1

If pH = pKa Then Charge = 0 Else Charge = ChargeSign / (1 + 10 ^ (ChargeSign * (pH - pKa)))

AACharge = Charge

End Function
'****************************************************************************************************
Function Theoretical_pI(ProteinSequence As String, Optional pKaValues As String = "ProtParam") As Double

'====================================================================================================
'Calculates theoretical pI from a given protein sequence
'Juraj Ahel, 2015-02-02, for the table of constructs so I don't ever have to go to ProtParam
'Last update 2015-02-02
'====================================================================================================
'Requires AACharge() and StringCharCount()
'Results not same as in ProtParam, generally slightly smaller. Check for exact pKa values used in ProtParam!

'aminoacid representations: 1=D; 2=E; 3=R; 4=K; 5=H; 6=C; 7=Y; 8=Cterm; 9=Nterm

Dim AminoAcids
Dim AASpecies

AminoAcids = Array("D", "E", "R", "K", "H", "C", "Y")
AASpecies = Array(0, 0, 1, 1, 1, 0, 0, 0, 1) '0 is acid, 1 is base

Dim AACounts(1 To 9) As Long
Dim PartialCharges(1 To 9) As Double
Dim pH As Double, TotalCharge As Double
Dim i As Long
Dim pHl As Double, pHh As Double

Dim pKa(1 To 9) As Double

Select Case UCase(pKaValues)
    Case "WIKIPEDIA": pKa(1) = 3.9: pKa(2) = 4.3: pKa(3) = 12.01: pKa(4) = 10.5: pKa(5) = 6.08: pKa(6) = 8.28: pKa(7) = 10.1: pKa(8) = 3.7: pKa(9) = 8.2
    Case "PROTPARAM": pKa(1) = 4.05: pKa(2) = 4.45: pKa(3) = 12: pKa(4) = 10: pKa(5) = 5.98: pKa(6) = 9: pKa(7) = 10: pKa(8) = 3.55: pKa(9) = 7.4
    Case Else: pH = 0: GoTo 10
End Select

For i = 1 To 7
    AACounts(i) = StringCharCount(ProteinSequence, AminoAcids(i - 1))
Next i

AACounts(8) = 1: AACounts(9) = 1

pH = 0
pHl = 0
pHh = 14

Do
    TotalCharge = 0
    
    For i = 1 To 9
        PartialCharges(i) = AACharge(pKa(i), pH, AASpecies(i - 1))
        TotalCharge = TotalCharge + AACounts(i) * PartialCharges(i)
    Next i
    
    If Abs(TotalCharge) < 0.01 Then GoTo 10
    
    If TotalCharge > 0 Then
        pHl = pH
        pH = (pH + pHh) / 2
    Else
        pHh = pH
        pH = (pH + pHl) / 2
    End If
    
Loop

10 Theoretical_pI = Round(pH, 1)
    
End Function

