VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const conSecondsInADay = 86400

Private pStart As Double
Private pEnd As Double

Private pTimerRunning As Boolean

Public Property Let TimeStart(ByVal a As Double)
    pStart = a
End Property

Public Property Let TimeEnd(ByVal a As Double)
    pEnd = a
End Property

Public Property Get Duration() As Double
    Duration = pEnd - pStart
    If Duration < 0 Then Duration = conSecondsInADay + Duration
End Property

Public Sub StartTimer()
    TimeStart = Timer
End Sub

Public Sub EndTimer()
    TimeEnd = Timer
End Sub

Public Sub Trigger()
    If pTimerRunning = False Then
        StartTimer
        pTimerRunning = True
    Else
        EndTimer
        pTimerRunning = False
    End If
End Sub

