Attribute VB_Name = "test"
Option Explicit

Public i As Long

Sub kkk()

    Dim a As Collection
    Dim b As String
    
    b = "item1"
    
    Set a = New Collection
    
    a.Add b, b
    a.Add 1, "1"
    a.Remove b
    a.Remove b

End Sub

Sub aaaaa()

    Dim a As String
    Dim FilePath As String
    
    FilePath = "C:\Users\juraj.ahel\Desktop\New folder (2)\Pur160205 SEC3 IMAC1 IEX.zip"
    
    a = "aaa"
    
    Dim chrom As clsGeneralizedChromatography
    Set chrom = New clsGeneralizedChromatography
    
    Call chrom.ImportFile(FilePath, "UNICORN6")
    
    'uni6.ImportUNI6ZipFile (FilePath)

End Sub



Sub test1()

    Dim testString As String
    
    Dim result As VBA.Collection
    
    Dim unpak As clsBinaryUnpacker
    
    Dim teststr As String
    
    teststr = "This is a test."
    
    Call WriteBinaryFile(teststr)
    
    'testString = String(64, Chr(255))
    
    testString = ReadBinaryFile("C:\Users\juraj.ahel\Desktop\New folder (2)\CoordinateData.Volumes")
    
    testString = Right(testString, Len(testString) - 35)
    testString = Left(testString, Len(testString) - 1)
    
    Set unpak = New clsBinaryUnpacker
    
    'Set result = unpak.UnpackBinaryData("<f", testString, 35)
    Set result = unpak.UnpackBinaryData("<f", testString, 0)
    
End Sub

Sub testyx()

Dim a() As Variant
Dim b(1 To 1) As String

b(1) = "aaa"

a = b


End Sub

Sub TESTAR()

Dim inputfile As String

Dim final As String

inputfile = "C:\Users\juraj.ahel\Desktop\New folder (2)\Pur160205 SEC3 IMAC1 IEX.zip"

final = FileSystem_Unzip(inputfile)

End Sub

Sub testx()

    Dim xDoc As MSXML2.DOMDocument
    Dim XML As String
        
    XML = "<root><person><name><first>Juraj</first><last>Ahel</last></name> </person> <person> <name>No Name </name></person></root> "
    
    Dim oXml As MSXML2.DOMDocument60
    Set oXml = New MSXML2.DOMDocument60
    
    
    Dim tempNodes As IXMLDOMNodeList, tempNode As IXMLDOMNode, tempnode2 As IXMLDOMNode
    
    Dim test
    
    oXml.LoadXML XML
    
    Dim oSeqNodes, oSeqNode As IXMLDOMNode

    Set oSeqNodes = oXml.SelectNodes("//root/person")
    Set test = oXml.SelectSingleNode("//root/person")
    If oSeqNodes.Length = 0 Then
       'show some message
    Else
        For Each oSeqNode In oSeqNodes
             Debug.Print oSeqNode.SelectSingleNode("name").Text
             
             Set tempNodes = oSeqNode.SelectNodes("name")
             Set tempNode = tempNodes.Item(0)
             Set tempnode2 = tempNode.SelectSingleNode("first")
             
             
             Debug.Print tempnode2.Text
        Next
        
        'oSeqNode.SelectNodes ("name/first")
        
    End If
    
    


End Sub

Sub testa()

    Dim main1 As MainForm1
    
    Set main1 = New MainForm1
    
    main1.Show (vbModeless)

End Sub

Sub testb()

    Dim main1 As frmCreateChromatograms
    
    Set main1 = New frmCreateChromatograms
    
    main1.Show (vbModeless)

End Sub

Sub testc()

Dim i As Long

    For i = 0 To frmMetadata.Controls.Count - 1
    Debug.Print frmMetadata.Controls(i).Name
    frmMetadata.Controls(i).Name = "x" & frmMetadata.Controls(i).Name
    Next i

End Sub
