VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "A00_clsClassTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'====================================================================================================
'This class does briefly this and that
'Juraj Ahel, 2016-MM-DD, for this and that purpose
'Last update 2016-MM-DD
'====================================================================================================
Option Explicit

Private Const conClassName As String = "clsClassTemplate"

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

Private Sub ErrorReport(Optional ByVal ErrorNumber As Long = 0, Optional ByVal ErrorString As String = 0)

    Const conDefaultErrorN As Long = 1
    Const conDefaultError As String = "An undocumented error has occured."
    
    If ErrorNumber = 0 Then ErrorNumber = conDefaultErrorN
    
    If Len(ErrorString) = 0 Then ErrorString = conDefaultError
    
    Err.Raise vbError + ErrorNumber, conClassName, ErrorString

End Sub

