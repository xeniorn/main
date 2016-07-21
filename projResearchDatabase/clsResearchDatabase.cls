VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResearchDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2016-06-28, for this and that purpose
'Last update 2016-06-28
'====================================================================================================
Option Explicit

Private Const conClassName As String = "clsResearchDatabase"

'[Var Declaration]
Private pDatabase As VBA.Collection

Public Event NameClash(Name As String)

'[Object Initialization]
Private Sub Class_Initialize()

    Set pDatabase = New VBA.Collection

End Sub

'[Object Dereferencing]
Private Sub Class_Terminate()

    Set pDatabase = Nothing

End Sub

Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    'If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Debug.Print (conClassName & ":" & vbError + ErrorNumber & ":" & ErrorString)
    
    If Len(ErrorString) = 0 Then
        Err.Raise vbError + ErrorNumber, conClassName, ErrorString
    Else
        Err.Raise vbError + ErrorNumber, conClassName
    End If

End Sub

'[Property Gets and Lets and Sets]
Public Property Get Database() As VBA.Collection
    Set Database = pDatabase
End Property

Public Property Set Database(inp_Database As VBA.Collection)
    Set pDatabase = inp_Database
End Property


Public Sub LoadFile(ByVal Filename As String)

End Sub

Private Function EntryExists(ByVal Name As String) As Boolean
    EntryExists = IsElementOf(Name, Database)
End Function

Public Sub AddNewElement( _
    ByVal Name As String, _
    ByVal EntityType As String)
    
    Dim tDatabaseEntry As clsResearchDatabaseObject
    Dim tResearchObj As IResearchObject
    Dim tObj As Object
    
    If EntryExists(Name) Then
        Debug.Print ("Entry already exists in database (" & Name & ")")
        RaiseEvent NameClash(Name)
        Exit Sub
    Else
        Select Case UCase(EntityType)
        
            Case "1", "PHYSICAL"
                Set tResearchObj = New clsPhysicalObject
                Set tObj = New clsGenericPhysicalEntity
                
            Case "2", "EXPERIMENT"
                Set tResearchObj = New clsExperimentObject
                Set tObj = New clsGenericExperiment
                
            Case Else
                Call ErrorReport(, "Unsupported or invalid entry type")
                
        End Select
    End If
                
        Call tResearchObj.Add(tObj)
        tResearchObj.Name = Name
        
        Set tDatabaseEntry = New clsResearchDatabaseObject
        Set tDatabaseEntry.ActualObject = tResearchObj
            
        Database.Add tDatabaseEntry, Name
            
        Set tObj = Nothing
        Set tResearchObj = Nothing
        Set tDatabaseEntry = Nothing
        
End Sub
