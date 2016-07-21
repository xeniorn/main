VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClassTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2016-03-07, for this and that purpose
'Last update 2016-03-12
'====================================================================================================
Option Explicit

'This is an example Class, using all the basic elements I deem important for classes to have in general
'not all elements need to be used for each class, of course, but it must absolutely have initialization
'for object parameters if it has them, and termination of them to prevent memory leaks, an error reporting
'method so it allows for robust and modular error reporting from within the class, and ideally the clone method,
'which allows creating a new instance of the object that is identical to the parent

Private pName As String
Private pCollectionObject As Collection

Private Sub Class_Initialize()

    Set pCollectionObject = New Collection

End Sub

Private Sub Class_Terminate()

    Set pCollectionObject = Nothing

End Sub

Public Sub CloneFrom(c As Collection)

    'Needed for the Clone function
    
    pName = c.Item("pName")
    pCollectionObject = c.Item("pCollectionObject")

End Sub

Public Function Clone() As clsClassTemplate

    Dim ClassParameters As Collection
    
    Set ClassParameters = New Collection
    
    'Construct a collection with all the parameters that define the class,
    'using the variable names as keys, to ensure against accidents.
    'It unfortunately involves another public sub, CloneFrom, which is necessarily
    'public as it needs to be called from within the original instance of the class.
    'Alternatives possibly exist, but there are no "clean" solutions according to StackExchange (20160120)
    ClassParameters.Add pName, "pName"
    ClassParameters.Add pCollectionObject, "pCollectionObject"
    
    Set Clone = New clsClassTemplate
    
    Clone.CloneFrom ClassParameters
    
    Set ClassParemeters = Nothing

End Function

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(N As String)
    pName = N
End Property

Public Property Get Container() As Collection
    Set Container = pCollectionObject
End Property

Public Property Set Container(N As Collection)
    Set pCollectionObject = N
End Property

Private Sub ErrorReport(ErrorNumber As Integer, ErrorString As String)

    Err.Raise vbError + ErrorNumber, , ErrorString

End Sub
