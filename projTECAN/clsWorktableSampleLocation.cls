VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorktableSampleLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pContainer As clsWorktableContainer
Private pCoordinate As Integer
Private pUnique As Boolean

Private Sub Class_Initialize()

pUnique = True

End Sub

Private Sub Class_Terminate()

Set pContainer = Nothing

End Sub



Public Property Get Container() As clsWorktableContainer
    Set Container = pContainer
End Property

Public Property Set Container(c As clsWorktableContainer)
    Set pContainer = c
    
End Property

Public Property Get WellCoordinate() As Integer
    WellCoordinate = pCoordinate
End Property

Public Property Let WellCoordinate(c As Integer)
    pCoordinate = c
End Property

Public Property Get Grid() As Integer
    Grid = Container.Grid
End Property

Public Property Get Site() As Integer
    Site = Container.Site
End Property

Public Property Get Unique() As Boolean
    Unique = pUnique
End Property

Public Property Let Unique(c As Boolean)
    pUnique = Unique
End Property



