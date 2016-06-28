Attribute VB_Name = "OLDCODE20160628"
Option Explicit

'****************************************************************************************************
Function ConstructVector(InsertSeq As String, VectorSeq As String, StartIndex As Long, EndIndex As Long, ForwardStrand As Boolean) As String

Dim N As Long: N = Len(VectorSeq)

If Not ForwardStrand Then
    VectorSeq = DNAReverseComplement(VectorSeq)
    StartIndex = N - StartIndex + 1
    EndIndex = N - EndIndex + 1
End If

Dim tempString(1 To 3) As String
Dim tempOutput As String

tempString(1) = SubSequenceSelect(VectorSeq, 1, StartIndex)
tempString(2) = InsertSeq
tempString(3) = SubSequenceSelect(VectorSeq, EndIndex, N)


tempOutput = Join(tempString, "")

If Not ForwardStrand Then tempOutput = DNAReverseComplement(tempOutput)

ConstructVector = tempOutput

End Function



'****************************************************************************************************
Function ConstructVectorParse(insert1 As Variant, vector1 As Variant, si1 As Variant, ei1 As Variant, forr1 As Variant) As String

Dim Vector As String, insert As String, si As Long, ei As Long, forr As Boolean

forr = CBool(forr1)

Vector = CStr(vector1)

insert = CStr(insert1)

si = CInt(si1)
ei = CInt(ei1)

ConstructVectorParse = ConstructVector(insert, Vector, si, ei, forr)

End Function
