Attribute VB_Name = "modChromatography"
Option Explicit

Public Type AxisDataType
    Label As String
    Unit As String
    Description As String
End Type

Sub CallUnicornForm()

    Dim a As AxisDataType
    
    Const conDefaultSheetName As String = "Unicorn"
    
    Const file1 As String = "Y:\Juraj\FPLC\Mpa\Pur160427 Mpa 5c SeMet SEC1.res"
    Const file2 As String = "C:\Users\juraj.ahel\Downloads\sample1.res"
    Const file3 As String = "C:\temp\CLA24\UNICORN\Local\Fil\Default\Result\Markus\Manual run 1.res"
    Const file4 As String = "C:\Users\juraj.ahel\Desktop\Pur160121 Mpa 5c(SeMet) SEC S6 16 70.res"
    Const file5 As String = "Y:\Juraj\FPLC\JA testMethod Binary (1462569897)001.res"
    Const file6 As String = "C:\temp\Juraj Ahel FPLC\Gimli\Default\Result\Juraj\Ettan\Ettan Mpa\Ettan150601E P12.res"
    Const file7 As String = "C:\Users\juraj.ahel\Desktop\New folder (2)\Pur160205 SEC3 IMAC1 IEX.zip"
    Const file8 As String = "Y:\Personal Folders\Juraj\FPLC\Mpa\Pur160427 Mpa 5c SeMet SEC1.res"
    
    'CreateSheetFromName (conDefaultSheetName)
    
    Dim tmpFrm As frmCreateChromatograms
    
    Set tmpFrm = New frmCreateChromatograms
    
    With tmpFrm

        '.DefaultStartFolder = "Y:\Juraj\FPLC\Mpa\"
        .DefaultFileName = file8
        
        .ManualInitialize
        
        'modeless allows interaction with sheet while it's on!
        'Modal mode should only be on when user input would actually interfere with the actions
        .Show (vbModeless)
    
        '.Hide
        
    End With
    
    
End Sub

'*******************************************************************************
Function GetDefaultSettings(Optional ByVal CurveTypeName As String = "") As clsSeriesFormatSettings

'===============================================================================
'defines the settings class for some curve types
'required by frmCreateChromatograms
'Juraj Ahel, 2016-05-22
'Last update 2016-05-24
'===============================================================================
    
    Set GetDefaultSettings = New clsSeriesFormatSettings
    
    With GetDefaultSettings
        Select Case UCase(CurveTypeName)
            Case "UV1"
                'this is the default
            Case "UV2"
                
            Case "UV3"
            
            Case "COND"
                .LineColor = RGB(237, 125, 49)
                .LineWeight = xlHairline
            Case "CONC"
                .LineColor = RGB(147, 149, 152)
                .LineWeight = xlHairline
                .LineSmooth = False
            Case Else
                .LineWeight = xlHairline
                'nothing
        End Select
    End With
    
End Function


Sub UnicornWrapper()

    Const file1 As String = "Y:\Juraj\FPLC\Mpa\Pur160427 Mpa 5c SeMet SEC1.res"
    Const file2 As String = "C:\Users\juraj.ahel\Downloads\sample1.res"
    Const file3 As String = "C:\temp\CLA24\UNICORN\Local\Fil\Default\Result\Markus\Manual run 1.res"
    Const file4 As String = "Y:\Juraj\FPLC\Mpa\Pur160427 Mpa 5c SeMet IMAC1 HisTrap5.res"
    Const file5 As String = "Y:\Juraj\FPLC\JA testMethod Binary (1462569897)001.res"
    Const file6 As String = "C:\temp\Juraj Ahel FPLC\Gimli\Default\Result\Juraj\Ettan\Ettan Mpa\Ettan150601E P12.res"
    Const file7 As String = "C:\Users\juraj.ahel\Desktop\test1.txt"
    
    Dim InputFilename As String
    
    InputFilename = file1
    
    
    Dim SEC As clsSizeExclusionChromatography
    
    Set SEC = New clsSizeExclusionChromatography
    
    Call SEC.ImportFile(InputFilename, "UNICORN3")


End Sub



Sub tempOutput(Headers As Collection)

    Dim ctempHeader As clsUnicorn3Header
    
    Dim i As Long
    Dim j As Long
    
    Dim tempOutArray(1 To 1, 1 To 6)
    
    Dim outrange As Excel.Range
        
    For i = 1 To Headers.Count
        
        Set ctempHeader = Headers.Item(i)
        tempOutArray(1, 1) = ctempHeader.MagicID
        tempOutArray(1, 2) = CStr(ctempHeader.Name)
        tempOutArray(1, 3) = ctempHeader.DataSize
        tempOutArray(1, 4) = ctempHeader.DataOffsetToNext
        tempOutArray(1, 5) = ctempHeader.DataAddress
        tempOutArray(1, 6) = ctempHeader.OffsetMetaToData
        
        Set outrange = Excel.Range(Cells(i + 1, 1), Cells(i + 1, 6))
        
        outrange.Value = tempOutArray
        
        For j = 1 To 6
            tempOutArray(1, j) = ""
        Next j
        
    Next i
        
        
        Set ctempHeader = Nothing

End Sub


Sub TestCSV()

    Dim aaa As clsComparativeChromatography
    
    Dim sera As String
    
    Dim conFileDialog As FileDialog
    Dim tempFile As String
    Dim FileName As String

    Set conFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With conFileDialog
    
        .AllowMultiSelect = False
        
        .InitialFileName = "E:\PhD\Mpa\160322-24 MALS data\MALS data UV.txt"
                
        .Show
        
        If .SelectedItems.Count > 0 Then
        
            tempFile = .SelectedItems.Item(1)
            
            'MsgBox ("Selected item: " & .SelectedItems.Item(1))
            
            FileName = tempFile
        
        Else
        
            
            FileName = vbEmpty
            
        End If
                
        
    End With
        
    
    Set aaa = New clsComparativeChromatography
    
    Call aaa.ImportTable(FileName)
    
    
    If MsgBox("Normalize data?", vbQuestion + vbYesNo) = vbYes Then
        Call aaa.NormalizeAll(2, 6)
    End If
    
    Call aaa.ThinData(0.025)
    
    Call aaa.tempDraw
    

End Sub

