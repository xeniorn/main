VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Compare(TestClass1 As IContainerSample, TestClass2 As IContainerSample) As Boolean

    Compare = (TestClass1.Name = TestClass2.Name)

End Function
