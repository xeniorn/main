VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IChromExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This is an interface all Chromatography experiments must implement
'Juraj Ahel, 2016-05-08
'Last update 2016-05-08
'====================================================================================================
Option Explicit

Implements IExperiment


Public Sub AddChromatogram(ByVal Graph As clsChromatogram)

End Sub

Public Function GetChromatogram(Optional ByVal Index As Variant = 1) As clsChromatogram

End Function

Private Property Get IExperiment_ExpDate() As String

End Property

Private Property Get IExperiment_ExpName() As String

End Property

Private Property Get IExperiment_ExpType() As String

End Property

Private Property Get IExperiment_Scientist() As String

End Property
