Attribute VB_Name = "XmodBinary"
Option Explicit



'from http://stackoverflow.com/questions/15782705/convert-8-bytes-array-into-double
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)

Private Function BytesToDbl(ByRef Bytes() As Byte) As Double
    Dim D As Double
    CopyMemory D, Bytes(0), LenB(D)
    BytesToDbl = D
End Function

Private Function BytesToSingle(ByRef Bytes() As Byte) As Single
    Dim D As Single
    CopyMemory D, Bytes(0), LenB(D)
    BytesToSingle = D
End Function


'*******************************************************************************
Function ReverseBytes(ByVal InputVar As String)
'===============================================================================
'For a given binary string, reverses the byte order (reverse character order)
'Juraj Ahel, 2016-05-07, for reading binary files
'Last update 2016-05-08
'===============================================================================

    Dim i As Long
         
    Dim tempOut() As String
    
    Dim StringLength As Long
    
    StringLength = Len(InputVar)
    
    ReDim tempOut(0 To StringLength - 1)
        
    For i = 0 To StringLength - 1
        
        tempOut(i) = Mid(InputVar, StringLength - i, 1)
        
    Next i
    
    ReverseBytes = Join(tempOut, "")

End Function

'*******************************************************************************
Function BinaryStringToInt32( _
        ByVal InputVar As String, _
        Optional ByVal LittleEndian As Boolean = True _
        ) As Long
'===============================================================================
'
'Juraj Ahel, 2016-05-06, for reading binary files
'Last update 2016-06-09
'===============================================================================
'TODO: implement this more neatly, considering VBA strings are made out of 2-byte
'characters (ASCII + 00 (unicode))...
    
    Const TypeByteLength As Long = 4
    
    Dim ByteLength As Byte
    'Dim ByteArray(1 To 4) As Byte
    Dim tempval As Long ', tempval1 As Long
        
    ByteLength = Len(InputVar)
    
    If Not LittleEndian Then
    
        InputVar = ReverseBytes(InputVar)
        
    End If
    
    If ByteLength = TypeByteLength Then
                
        tempval = CLng("&H" & AsHex(InputVar, ""))
        'tempval1 = CInt(InputVar)
    
    Else
    
        Call Err.Raise(1001, , "4 bytes required for a long")
        
    End If
    
    BinaryStringToInt32 = tempval

End Function


'*******************************************************************************
Function BinaryStringToDouble( _
        ByVal InputVar As String, _
        Optional ByVal LittleEndian As Boolean = True _
                ) As Double
''Optional ByVal LittleEndian As Boolean = True _
'===============================================================================
'
'Juraj Ahel, 2016-05-10, for reading binary files
'Last update 2016-06-09
'===============================================================================
    
    Const TypeByteLength As Long = 8
    
    Dim ByteLength As Byte
    'Dim ByteArray(1 To 4) As Byte
    Dim tempBytes(0 To 7) As Byte
    
    Dim i As Long
        
    ByteLength = Len(InputVar)
    
    If Not LittleEndian Then
   
        InputVar = ReverseBytes(InputVar)
   
    End If
    
    If ByteLength = TypeByteLength Then
                
        For i = 0 To 7
            tempBytes(i) = Asc(Mid(InputVar, i + 1, 1))
        Next i
        
        'tempBytes = Inputvar
        
        'tempval = CDbl("&H" & AsHex(Inputvar, ""))
        'tempval1 = CInt(InputVar)
    
    Else
    
        Call Err.Raise(1001, , "8 bytes required for a double precision float")
        
    End If
    
    '0E74DA40A70DB43F
    '3FB40DA740DA740E
    
    'CopyMemory BinaryStringToDouble, tempBytes(0), LenB(BinaryStringToDouble)
    BinaryStringToDouble = BytesToDbl(tempBytes)

End Function

'*******************************************************************************
Function BinaryStringToSingle( _
        ByVal InputVar As String, _
        Optional ByVal LittleEndian As Boolean = True _
                ) As Single
''Optional ByVal LittleEndian As Boolean = True _
'===============================================================================
'
'Juraj Ahel, 2016-06-09, for reading binary files
'Last update 2016-06-09
'===============================================================================
    
    Const TypeByteLength As Long = 4
    
    Dim ByteLength As Byte
    'Dim ByteArray(1 To 4) As Byte
    Dim tempBytes(0 To 3) As Byte
    
    Dim i As Long
        
    ByteLength = Len(InputVar)
    
    If Not LittleEndian Then
   
        InputVar = ReverseBytes(InputVar)
   
    End If
    
    If ByteLength = TypeByteLength Then
                
        For i = 0 To 3
            tempBytes(i) = Asc(Mid(InputVar, i + 1, 1))
        Next i
        
        'tempBytes = Inputvar
        
        'tempval = CDbl("&H" & AsHex(Inputvar, ""))
        'tempval1 = CInt(InputVar)
    
    Else
    
        Call Err.Raise(1001, , "4 bytes required for a single precision float")
        
    End If
    
    '0E74DA40A70DB43F
    '3FB40DA740DA740E
    
    'CopyMemory BinaryStringToDouble, tempBytes(0), LenB(BinaryStringToDouble)
    BinaryStringToSingle = BytesToSingle(tempBytes)

End Function

'Function HexToBinary(ByVal Inputvar As String)



'end Function

Function HexToString(ByVal InputVar As String, Optional SourceSpacer As String = "") As String

'===============================================================================
'Converts a string representing a hex number (one you get e.g. by reading a file in binary mode)
'to a string of characters with the byte values represented by hex-value pairs
'Juraj Ahel, 2016-05-10, for reading binary files
'Last update 2016-05-10
'===============================================================================
'TODO: detect spacer

    Dim i As Long
    
    Dim StringLength As Long
    
    'offset from byte to byte, counting the byte length and the spacer length
    Dim DataOffset1 As Long
    
    Dim NumberOfChar As Long
        
    DataOffset1 = Len(SourceSpacer) + 2
       
    Dim tempChar() As String
    
    StringLength = Len(InputVar)
               
    NumberOfChar = (StringLength + Len(SourceSpacer)) \ DataOffset1
               
    ReDim tempChar(1 To NumberOfChar)
           
    i = 0
       
    If (StringLength - 2) Mod DataOffset1 = 0 Then
            
        For i = 1 To NumberOfChar
        
            tempChar(i) = Chr(CDbl("&H" & Mid(InputVar, 1 + DataOffset1 * (i - 1), 2)))
                    
        Next i
    
        HexToString = Join(tempChar, "")
        
    Else

        Call Err.Raise(1001, , "Input data length isn't a multiple of (2 + spacer length)")

    End If

End Function

Function AsHex(InputVar As String, Optional Spacer As String = " ") As String

'===============================================================================
'Converts a string (one you get e.g. by reading a file in binary mode)
'to hex characters (separated by spaces by default), in the original order
'Juraj Ahel, 2016-05-06, for reading binary files
'Last update 2016-05-09
'===============================================================================
    
    'Const conHexFormat As String = "00"
    
    Dim i As Long
    
    Dim StringLength As Long
    
    'Dim tempOut As String
       
    Dim tempChar() As String
    
    StringLength = Len(InputVar)
           
    ReDim tempChar(1 To StringLength)
           
    i = 0
       
    'tempOut = ""
        
    For i = 1 To StringLength
        
        tempChar(i) = Hex(Asc(Mid(InputVar, i, 1)))
        
        If Len(tempChar(i)) = 1 Then
            tempChar(i) = "0" & tempChar(i)
        End If
        
        'tempOut = tempOut & Spacer & Format(Hex(Asc(Mid(InputVar, i, 1))), conHexFormat)
        'tempOut = Join(tempChar, Spacer)
        
        
    Next i

    AsHex = Join(tempChar, Spacer)
        
End Function

