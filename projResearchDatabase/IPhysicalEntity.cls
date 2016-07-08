VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPhysicalEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This is an interface all samples, gels, crystals, ...
'Juraj Ahel, 2016-06-28
'Last update 2016-06-28
'====================================================================================================
Option Explicit

Public Property Get Name() As String
End Property

Public Property Let Name(inpName As String)
End Property

Public Property Get EntityType() As String
End Property

Public Property Let EntityType(inpEntityType As String)
End Property
