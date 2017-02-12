VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGibsonAssembly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'
'Juraj Ahel, 2017-01-21
'
'====================================================================================================

Option Explicit

Private Const conClassName As String = "clsClassTemplate"

Private Const conRequiredRows As Long = 5

'[Var Declaration]
Private pLigations As VBA.Collection
Private pPCRs As VBA.Collection

Private pFragmentNumber As Long

Private pAssemblyName As String

Private GibsonDone As Boolean
Private PCRDone As Boolean
Private ReactionsImported As Boolean

Private pFinalAssembly As clsDNA

'[Object Initialization]
Private Sub Class_Initialize()

    
    Set pLigations = New VBA.Collection
    Set pPCRs = New VBA.Collection
    Set pFinalAssembly = New clsDNA
    
    GibsonDone = False
    PCRDone = False
    ReactionsImported = False

End Sub

'[Object Dereferencing]
Private Sub Class_Terminate()

    Set pLigations = Nothing
    
    Set pPCRs = Nothing
    Set pFinalAssembly = Nothing

End Sub

Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    'If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    If Len(ErrorString) = 0 Then
        Err.Raise vbError + ErrorNumber, conClassName, ErrorString
    Else
        Err.Raise vbError + ErrorNumber, conClassName
    End If

End Sub

'[Property Gets and Lets and Sets]

Public Property Get PCRs() As VBA.Collection
    Set PCRs = pPCRs
End Property

Public Property Get PCR(ByVal i As Long) As clsPCR
    If Not PCRs Is Nothing Then
        If PCRs.Count >= i Then
            Set PCR = PCRs.Item(i)
        End If
    End If
End Property


Public Property Get SourceDNA(ByVal i As Long) As clsDNA
    Dim tPCR As clsPCR
    Set tPCR = PCRs.Item(i)
    Set SourceDNA = tPCR.SourceDNA
    Set tPCR = Nothing
End Property

Public Property Set SourceDNA(ByVal i As Long, ByVal inp_SourceDNA As clsDNA)
    Dim tPCR As clsPCR
    Set tPCR = PCRs.Item(i)
    Set tPCR.SourceDNA = inp_SourceDNA
End Property

Public Property Get Ligations() As VBA.Collection
    Set Ligations = pLigations
End Property

Public Property Set Ligations(ByVal inp_Ligations As VBA.Collection)
    Set pLigations = inp_Ligations
End Property

Public Property Get FinalDNA(ByVal i As Long) As clsDNA
    Dim tPCR As clsPCR
    Set tPCR = PCRs.Item(i)
    Set FinalDNA = tPCR.FinalDNA
    Set tPCR = Nothing
End Property

Public Property Set FinalDNA(ByVal i As Long, ByVal inp_FinalDNA As clsDNA)
    Dim tPCR As clsPCR
    Set tPCR = PCRs.Item(i)
    Set tPCR.FinalDNA = inp_FinalDNA
End Property

Public Property Get ForwardPrimer(ByVal i As Long) As clsDNA
    Dim tPCR As clsPCR
    Set tPCR = PCRs.Item(i)
    Set ForwardPrimer = tPCR.ForwardPrimer
    Set tPCR = Nothing
End Property

Public Property Set ForwardPrimer(ByVal i As Long, ByVal inp_ForwardPrimer As clsDNA)
    Dim tPCR As clsPCR
    Set tPCR = PCRs.Item(i)
    Set tPCR.ForwardPrimer = inp_ForwardPrimer
End Property

Public Property Get ReversePrimer(ByVal i As Long) As clsDNA
    Dim tPCR As clsPCR
    Set tPCR = PCRs.Item(i)
    Set ReversePrimer = tPCR.ReversePrimer
    Set tPCR = Nothing
End Property

Public Property Set ReversePrimer(ByVal i As Long, ByVal inp_ReversePrimer As clsDNA)
    Dim tPCR As clsPCR
    Set tPCR = PCRs.Item(i)
    Set tPCR.ReversePrimer = inp_ReversePrimer
End Property

Public Property Get FragmentNumber() As Long
    FragmentNumber = pFragmentNumber
End Property

Public Property Let FragmentNumber(ByVal inp_FragmentNumber As Long)
    pFragmentNumber = inp_FragmentNumber
End Property

Public Property Get FinalSequence() As String
    FinalSequence = pFinalAssembly.Sequence
End Property

Public Property Get Name() As String
    Name = FinalAssembly.Name
End Property

Public Property Let Name(ByVal S As String)
    FinalAssembly.Name = S
End Property

Public Property Get FinalAssembly() As clsDNA
    Set FinalAssembly = pFinalAssembly
End Property

Private Sub ResetCalculation()
    
    Set pPCRs = New VBA.Collection
    Set pLigations = New VBA.Collection
    
End Sub

Private Sub DoCalculation()
    
    'HACKED UP - WILL HAVE ISSUES WITH REPETITIVE OVERLAPS THAT ARE NOT REAL ISSUES (e.g. if an overlap also exists elsewhere in the sequence)
    
    
    Dim tDNA As clsDNA
    Dim tGibson1 As clsGibsonSingleLigation
    Dim tGibson2 As clsGibsonSingleLigation
    
    Dim CurrentFragment As Long
    Dim NextFragment As Long
    Dim PreviousFragment As Long
    
    'only if there are actual inputs already present
    If ReactionsImported Then
        
        Set tDNA = New clsDNA
        
        For CurrentFragment = 1 To FragmentNumber
        
            'define previous and next fragments - assuming circularity
                If CurrentFragment = 1 Then
                    PreviousFragment = FragmentNumber
                Else
                    PreviousFragment = CurrentFragment - 1
                End If
                
                If CurrentFragment = FragmentNumber Then
                    NextFragment = 1
                Else
                    NextFragment = CurrentFragment + 1
                End If
            
            'define final Gibson fragments
                Set tGibson1 = Ligations.Item(PreviousFragment)
                Set tGibson2 = Ligations.Item(CurrentFragment)
                Set tDNA = PCR(CurrentFragment).FinalDNA
                If tDNA.Name = "" Then
                    tDNA.Name = "GibsonFragment_" & Format(CurrentFragment, "00")
                End If
                tDNA.Sequence = PCRSimulate(FinalSequence, tGibson1.Overlap, DNAReverseComplement(tGibson2.Overlap), Circular:=True)
                Set PCR(CurrentFragment).FinalDNA = tDNA
                
            'Do PCR
                
                With PCR(CurrentFragment)
                    
                    Set tColl = PCRGetFragmentFromTemplate( _
                        .FinalDNA.Sequence, _
                        .SourceDNA.Sequence, _
                        62, _
                        15, _
                        .SourceDNA.Circular, _
                        MaxExtension:=75)
                        
                    Set .ForwardPrimer = tDNA.DefineNew(.FinalDNA.Name & "_f", tColl.Item("F"))
                    Set .ReversePrimer = tDNA.DefineNew(.FinalDNA.Name & "_r", tColl.Item("R"))
                    
                End With
                        
        Next CurrentFragment
        
        GibsonDone = True
        PCRDone = True
        
    End If
        

End Sub

Public Sub Yes()

    Call DoCalculation

End Sub

Public Function ConfirmGibson() As Boolean

    Dim i As Long
    Dim tColl As VBA.Collection
    Dim DNA() As String
    
    Dim fseq As String
    
    ReDim DNA(1 To FragmentNumber)
    
    Set tColl = New VBA.Collection
    
    For i = 1 To FragmentNumber
        tColl.Add PCRWithOverhangs(PCR(i).SourceDNA.Sequence, PCR(i).ForwardPrimer.Sequence, PCR(i).ReversePrimer.Sequence)
        DNA(i) = tColl(i)
    Next i
    
    fseq = DNAGibsonLigation(DNA)
    
    If DNAEqual(fseq, FinalSequence, True, True) Then ConfirmGibson = True
        
    
End Function

Public Sub ImportCollection(ByVal InputCollection As VBA.Collection)

    
    Dim tDNA As clsDNA
    Dim tGibson As clsGibsonSingleLigation
    
    Dim tempSequence As String
    
    Dim CurrentFragment As Long
    Dim NextFragment As Long
    Dim PreviousFragment As Long
    
    Dim tInputPrev As clsGibsonInput
    Dim tInputCurr As clsGibsonInput
    Dim tInputNext As clsGibsonInput
    
    Dim tPCR As clsPCR
    
    Dim i As Long
    
    If InputCollection Is Nothing Then
        Exit Sub
        Call ErrorReport(1, "Bad Input")
    End If
    
    If InputCollection.Count = 0 Then
        Exit Sub
        Call ErrorReport(1, "Bad Input")
    End If
    
    For i = 1 To InputCollection.Count
        If Not TypeOf InputCollection.Item(i) Is clsGibsonInput Then
            Exit Sub
            Call ErrorReport(1, "Bad Input")
        End If
    Next i
    
    
    FragmentNumber = InputCollection.Count
    
    ResetCalculation
    
    For CurrentFragment = 1 To FragmentNumber
        
        'define previous and next fragments - assuming circularity
            If CurrentFragment = 1 Then
                PreviousFragment = FragmentNumber
            Else
                PreviousFragment = CurrentFragment - 1
            End If
            
            If CurrentFragment = FragmentNumber Then
                NextFragment = 1
            Else
                NextFragment = CurrentFragment + 1
            End If
            
            Set tInputCurr = InputCollection.Item(CurrentFragment)
            Set tInputPrev = InputCollection.Item(PreviousFragment)
            Set tInputNext = InputCollection.Item(NextFragment)
            
            Set tPCR = New clsPCR
            PCRs.Add tPCR
            With tPCR
                Set .SourceDNA = tInputCurr.Source
                .FinalDNA.Name = tInputCurr.Name
            End With
            
        '2 Gibson Ligation
            Set tGibson = New clsGibsonSingleLigation
            With tGibson
                .FirstSequence = tInputCurr.Fragment.Sequence
                .InsertedSequence = tInputCurr.InsertAfter.Sequence & tInputNext.InsertBefore.Sequence
                .SecondSequence = tInputNext.Fragment.Sequence
                .ForbiddenRegions = tInputCurr.ForbiddenRegions
            End With
            
            Ligations.Add tGibson, CStr(CurrentFragment)
                    
         '3 define final sequence
            tempSequence = tempSequence & _
                tInputCurr.InsertBefore.Sequence & _
                tInputCurr.Fragment.Sequence & _
                tInputCurr.InsertAfter.Sequence
                
                    
    Next CurrentFragment
    
    
    Set tDNA = New clsDNA
        If tDNA.Name = "" Then
            tDNA.Name = "GibsonAssembly"
        End If
        tDNA.Sequence = tempSequence
    
    Set pFinalAssembly = tDNA
    
    ReactionsImported = True
    
    Set tDNA = Nothing
    Set tGibson = Nothing
    
    Set tInputCurr = Nothing
    Set tInputNext = Nothing
    Set tInputPrev = Nothing


End Sub

Public Sub ImportArray(ByVal InputArray As Variant)
    
    Const conDNASource As Long = 1
    Const conInsertBefore As Long = 2
    Const conFragment As Long = 3
    Const conInsertAfter As Long = 4
    Const conForbiddenRegions As Long = 5
    
    '1 DNA Source
    '2 Insert before fragment
    '3 Fragment
    '4 Insert after fragment
    '5 ForbiddenOverlapLocation
    
    Dim tDNA As clsDNA
    Dim tGibson As clsGibsonSingleLigation
    Dim tPCR As clsPCR
    
    Dim CurrentFragment As Long
    Dim NextFragment As Long
    Dim PreviousFragment As Long
    
    Dim tempSequence As String
    
    ResetCalculation
    
    FragmentNumber = 1 + UBound(InputArray, 2) - LBound(InputArray, 2)
    
    For CurrentFragment = LBound(InputArray, 2) To UBound(InputArray, 2)
        
        'define previous and next fragments - assuming circularity
            If CurrentFragment = LBound(InputArray, 2) Then
                PreviousFragment = UBound(InputArray, 2)
            Else
                PreviousFragment = CurrentFragment - 1
            End If
            
            If CurrentFragment = UBound(InputArray, 2) Then
                NextFragment = LBound(InputArray, 2)
            Else
                NextFragment = CurrentFragment + 1
            End If
        
        '1 DNA Source
            Set tDNA = New clsDNA
            With tDNA
                .Name = "Source_" & Format(CurrentFragment, "00")
                If Len(.Sequence) = 0 Then
                    .Sequence = DNAParseTextInput(InputArray(conFragment, CurrentFragment))
                    .Circular = False
                Else
                    .Sequence = DNAParseTextInput(InputArray(conDNASource, CurrentFragment))
                    .Circular = True
                End If
            End With
            
            Set tPCR = New clsPCR
            PCRs.Add tPCR
            Set PCR(CurrentFragment).SourceDNA = tDNA
            
        '2 Gibson Ligation
            Set tGibson = New clsGibsonSingleLigation
            With tGibson
                .FirstSequence = DNAParseTextInput(InputArray(conFragment, CurrentFragment))
                .InsertedSequence = DNAParseTextInput(InputArray(conInsertAfter, CurrentFragment)) & DNAParseTextInput(InputArray(conInsertBefore, NextFragment))
                .SecondSequence = DNAParseTextInput(InputArray(conFragment, NextFragment))
                .ForbiddenRegions = InputArray(conForbiddenRegions, CurrentFragment)
            End With
            
            Ligations.Add tGibson, CStr(CurrentFragment)
                    
         '3 define final sequence
            tempSequence = tempSequence & _
                DNAParseTextInput(InputArray(conInsertBefore, CurrentFragment)) & _
                DNAParseTextInput(InputArray(conFragment, CurrentFragment)) & _
                DNAParseTextInput(InputArray(conInsertAfter, CurrentFragment))
                
                    
    Next CurrentFragment
    
    
    Set tDNA = New clsDNA
        tDNA.Name = "GibsonAssembly"
        tDNA.Sequence = tempSequence
    
    Set pFinalAssembly = tDNA
    
    ReactionsImported = True
    
    Set tDNA = Nothing
    Set tGibson = Nothing
    
End Sub

Public Sub ImportRange(ByVal inputRange As Excel.Range)
    
    If inputRange.Rows.Count = conRequiredRows Then
        Call ImportArray(inputRange.Value2)
    End If
        
End Sub


