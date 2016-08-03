VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPipette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSample As clsSample
Private pVolume As Double
'Private pCount As Integer
Private pDirty As Boolean

Public Event CleanedWhileNotEmpty(SampleName As String, Volume As Double)

Private Sub Class_Initialize()
    
    Set pSample = New clsSample

End Sub

Private Sub Class_Terminate()

    Set pSample = Nothing

End Sub



Public Property Get IsDirty() As Boolean
    IsDirty = pDirty
End Property

Public Property Let IsDirty(State As Boolean)
    pDirty = State
End Property

Public Property Get Used() As Boolean
    'If Count > 0 Or Volume > 0 Then Used = True Else Used = False
    If Volume > 0 Then Used = True Else Used = False
End Property

Public Property Get Volume() As Double
    Volume = pVolume
End Property

Public Property Let Volume(N As Double)
    pVolume = N
    If Volume > 0 Then IsDirty = True
End Property

'Public Property Get Count() As Integer
'    Count = pCount
'End Property

'Public Property Let Count(N As Integer)
'    pCount = N
'End Property

Public Property Get Sample() As clsSample
    Set Sample = pSample
End Property

Public Property Set Sample(N As clsSample)
    Set Sample = N
End Property

Public Sub CleanMe()
    
    If Volume > 0 Then
        RaiseEvent CleanedWhileNotEmpty(Sample.Name, Volume)
    End If
    
    Volume = 0
    'Count = 0
    IsDirty = False
    Set pSample = New clsSample
End Sub

Public Sub DecreaseVolume(i As Double)
    
    Select Case i
        Case Is <= Volume
            Volume = Volume - i
        Case Else
            ErrorReportGlobal 69, "clsPipette: Tried to pipette out more than availabe Volume."
    End Select
    
End Sub


'Public Sub DecreaseCount(i As Integer)
'
'    Select Case i
'        Case Is <= Count
'            Count = Count - i
'        Case Else
'            ErrorReportGlobal 69, "clsPipette: Tried to pipette out more than count."
'    End Select
'
'    If Count = 0 Then
'        Volume = 0
'        Set Sample = Nothing
'    End If
'
'End Sub

