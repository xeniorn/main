VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IResearchDatabaseEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This is an interface all elements of my research database / logbook
'Juraj Ahel, 2016-06-28
'Last update 2016-06-28
'====================================================================================================
Option Explicit

Public Property Get ThingRelationships() As clsRelationships
End Property

Public Property Get ExperimentRelationships() As clsRelationships
End Property

Public Property Get EntityType() As String
End Property

Public Property Get EntityPath() As String
End Property
