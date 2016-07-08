VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTest1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String

Implements IContainerSample


Public Property Set IContainerSample_Location(RHS As clsWorktableSampleLocation)

End Property

Public Property Get IContainerSample_Location() As clsWorktableSampleLocation

End Property

Public Property Let IContainerSample_Name(RHS As String)
    pName = IContainerSample_Name
End Property

Public Property Get IContainerSample_Name() As String
    IContainerSample_Name = pName
End Property

Public Property Get IAmTest1() As String

    IAmTest1 = "IAmTest1!"

End Property



