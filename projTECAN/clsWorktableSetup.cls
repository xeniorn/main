VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorktableSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2015-12-xx
'Last update 2016-03-07
'====================================================================================================
Option Explicit

Const conAllowedPlateLocationsNumber As Byte = 6

Private pAllowedLocations() As String


Private pContainters As Collection

Private Sub Class_Initialize()
    
    Dim i As Integer

    Set pContainters = New Collection
    
    'Grid position 05 is the first plate container carrier in TECAN (2016-03)
    'Grid position 17 is the second plate container carrier in TECAN (2016-03)
    '01-03 are sites 1-3 on the carrier in the standard tecan 3-slot landscape plate carrier
        
    ReDim pAllowedLocations(1 To conAllowedPlateLocationsNumber, 1 To 2)
    pAllowedLocations(1, 1) = "12_00"
    pAllowedLocations(2, 1) = "12_01"
    pAllowedLocations(3, 1) = "12_02"
    pAllowedLocations(4, 1) = "18_00"
    pAllowedLocations(5, 1) = "18_01"
    pAllowedLocations(6, 1) = "18_02"
    
    For i = 1 To conAllowedPlateLocationsNumber
        pAllowedLocations(i, 2) = "free"
    Next i
    

End Sub

Private Sub Class_Terminate()

    Set pContainters = Nothing

End Sub

Private Sub ErrorReport(ErrorNumber As Integer, ErrorString As String)

    Err.Raise vbError + ErrorNumber, , ErrorString

End Sub

Public Property Get Containers() As Collection

    Set Containers = pContainters

End Property

Private Sub LegalCheck(LocationCode As String, Optional SetStatus As String = "")

    Dim i As Integer
    Dim IAmLegal As Boolean
    
    IAmLegal = False
    i = 0
    Do
        i = i + 1
        If LocationCode = pAllowedLocations(i, 1) Then
            IAmLegal = True
        End If
    Loop Until (i >= conAllowedPlateLocationsNumber) Or IAmLegal
    
    If Not IAmLegal Then
        ErrorReport 10, "clsWorktableSetup: A plate location is illegal. Check grid / site."
    Else
        Select Case UCase(SetStatus)
            Case "FREE"
                pAllowedLocations(i, 2) = "free"
            Case "FULL"
                pAllowedLocations(i, 2) = "full"
            Case Else
                'do nothing
        End Select
    End If

End Sub

Private Function LocationCode(Grid As Integer, Site As Integer) As String

    LocationCode = Format(Grid, "00") & "_" & Format(Site, "00")

End Function

Private Sub RefreshStatus()

    Dim tempCode As String
    Dim i As Integer
            
    For i = 1 To Containers.Count
        tempCode = LocationCode(Containers.Item(i).Grid, Containers.Item(i).Site)
        LegalCheck tempCode, "FuLL"
    Next i

End Sub

'****************************************************************************************************
Public Sub Locate(InputSample As clsSample)
'====================================================================================================
'locates a particular DNA sample inside any Worktable Container (source, not target!) on the Worktable
'annotates directly in the sample object
'Juraj Ahel, 2015-12-xx
'Last update 2016-03-15
'====================================================================================================
    
Dim i As Integer
Dim tempFound As Boolean
Dim tempcLocation As clsWorktableSampleLocation
Dim tempcContainer As clsWorktableContainer

'Set tempcLocation = New clsWorktableSampleLocation

tempFound = False

For i = 1 To Containers.Count
    
    If TypeOf Containers.Item(i) Is clsWorktableContainer Then
        
        Set tempcContainer = Containers.Item(i)
        Set tempcLocation = tempcContainer.LocateSample(InputSample)
            
        If Not (tempcLocation Is Nothing) Then
            If tempcLocation.Unique Then
                
                If Not tempFound Then
                    tempFound = True
                Else
                    ErrorReport 1, "Sample name ambiguity: more than one sample with one name in inputs (" & InputSample.Name & ")"
                    Set tempcLocation = Nothing
                End If
                
            End If
        End If
        
    End If
    
Next i

If tempcLocation Is Nothing Then
    ' sample wasn't found at all
    ErrorReport 2, "Sample name not found in inputs: (" & InputSample.Name & ")"
Else
    Set InputSample.Location = tempcLocation 'NOT GOOD, IT HAS MULTIPLE LOCATIONS! (input, output...), possible multiple of both
End If
    
Set tempcLocation = Nothing
Set InputSample = Nothing
    
End Sub

'****************************************************************************************************
Public Sub FindFreeLocationFor(VirtualPlate As clsWorktableSeqPlate)
'====================================================================================================
'Quite self-describing
'Juraj Ahel, 2016-03-07
'Last update 2016-03-07
'====================================================================================================

    Dim i As Integer
    Dim IFoundIt As Boolean
    Dim tempCode As String
    Dim tempGrid As Integer
    Dim tempSite As Integer
    
    RefreshStatus
    
    IFoundIt = False
    
    Do
        i = i + 1
        If UCase(pAllowedLocations(i, 2)) = "FREE" Then
            IFoundIt = True
        End If
    Loop Until IFoundIt Or (i >= conAllowedPlateLocationsNumber)
    
    If IFoundIt Then
        'code is of format "00_00" = "grid_site"
        tempCode = pAllowedLocations(i, 1)
        tempGrid = CInt(Left(tempCode, 2))
        tempSite = CInt(Right(tempCode, 2))
        VirtualPlate.DefineSite tempGrid, tempSite
        
    Else
        ErrorReport 11, "clsWorktableSetup: Tried to add another sequencing plate, but there are no more free sites."
    End If
    
    
End Sub
'****************************************************************************************************
Sub AddSequencingPlate(InputPlate As clsWorktableSeqPlate)
'====================================================================================================
'
'Juraj Ahel, 2016-03-07
'Last update 2016-03-07
'====================================================================================================
    
    Dim tempCode As String
    
    Containers.Add InputPlate
    
    RefreshStatus
    
End Sub

'****************************************************************************************************
Sub DefineSequencingPlatesFromVirtual(VirtualPlate As clsVirtualSequencingPlate)
'====================================================================================================
'creates physical plates from the virtual plate (split to 12x8 arrays, define location on worktable)
'Juraj Ahel, 2016-03-07
'Last update 2016-03-07
'====================================================================================================

Dim NumberOfPlates As Integer
Dim i As Integer
Dim tempcSeqPlate As clsWorktableSeqPlate

NumberOfPlates = VirtualPlate.NumberOfPlates

For i = 1 To NumberOfPlates
    
    Set tempcSeqPlate = VirtualPlate.ExtractPlate(1)
    
    FindFreeLocationFor tempcSeqPlate
    
    AddSequencingPlate tempcSeqPlate
    
    Set tempcSeqPlate = Nothing
    
Next i

End Sub

'****************************************************************************************************
Public Property Get SequencingPlates() As Collection
'====================================================================================================
'
'Juraj Ahel, 2016-03-09
'Last update 2016-03-15
'====================================================================================================

    Dim i As Integer
    Dim tempCollection As Collection
    Dim SequencingPlatesNumber As Integer
    
    Set tempCollection = New Collection

'[PART]define the output locations - read from worktable
    SequencingPlatesNumber = 0
    With Containers
        For i = 1 To .Count
            If VarType(.Item(i)) = vbObject Then
                If TypeOf .Item(i) Is clsWorktableSeqPlate Then
                    SequencingPlatesNumber = SequencingPlatesNumber + 1
                    tempCollection.Add .Item(i)
                End If
            Else
                ErrorReportGlobal 11, "clsWorktableSetup: Wrong type of variable loaded in Worktable container collection."
            End If
        Next i
    End With
    
    Set SequencingPlates = tempCollection
    
    Set tempCollection = Nothing

End Property

