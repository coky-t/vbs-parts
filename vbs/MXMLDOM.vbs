'Attribute VB_Name = "MXMLDOM"
Option Explicit

'
' Copyright (c) 2020 Koki Takeyama
'
' Permission is hereby granted, free of charge, to any person obtaining
' a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation
' the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the
' Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included
' in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'

'
' Microsoft XML, vX.X
' - MSXML2.DOMDocument
'

Private XMLDOM
Private BinBase64
Private BinHex

'
' --- MSXML2.DOMDocument ---
'

'
' GetXMLDOM
' - Returns a MSXML2.DOMDocument object.
'

Public Function GetXMLDOM()
    'Static XMLDOM
    If IsEmpty(XMLDOM) Then
        Set XMLDOM = CreateObject("MSXML2.DOMDocument")
    End If
    Set GetXMLDOM = XMLDOM
End Function

'
' GetBinBase64
' - Returns the IXMLDOMElement object with bin.base64 datatype.
'

Public Function GetBinBase64()
    'Static BinBase64
    If IsEmpty(BinBase64) Then
        Set BinBase64 = GetXMLDOM().createElement("tmp")
        BinBase64.DataType = "bin.base64"
    End If
    Set GetBinBase64 = BinBase64
End Function


'
' GetBinHex
' - Returns the IXMLDOMElement object with bin.hex datatype.
'

Public Function GetBinHex()
    'Static BinHex
    If IsEmpty(BinHex) Then
        Set BinHex = GetXMLDOM().createElement("tmp")
        BinHex.DataType = "bin.hex"
    End If
    Set GetBinHex = BinHex
End Function


'
' --- Base64 ---
'

'
' GetBase64TextFromBinary
' - Return the base64-encoded data.
'

'
' Binary:
'   Required. A Variant that contains an array of bytes.
'

Public Function GetBase64TextFromBinary(Binary)
    On Error Resume Next
    
    With GetBinBase64()
        .nodeTypedValue = Binary
        GetBase64TextFromBinary = .Text
    End With
End Function

'
' GetBinaryFromBase64Text
' - Return the resulting data.
'

'
' Base64Text:
'   Required. A String that contains a base64-encoded data.
'

Public Function GetBinaryFromBase64Text(Base64Text)
    On Error Resume Next
    
    With GetBinBase64()
        .Text = Base64Text
        GetBinaryFromBase64Text = .nodeTypedValue
    End With
End Function

'
' --- Hex ---
'

'
' GetHexTextFromBinary
' - Return the hex-text data.
'

'
' Binary:
'   Required. A Variant that contains an array of bytes.
'

Public Function GetHexTextFromBinary(Binary)
    On Error Resume Next
    
    With GetBinHex()
        .nodeTypedValue = Binary
        GetHexTextFromBinary = .Text
    End With
End Function

'
' GetBinaryFromHexText
' - Return the resulting data.
'

'
' HexText:
'   Required. A String that contains a hex-text data.
'

Public Function GetBinaryFromHexText(HexText)
    On Error Resume Next
    
    With GetBinHex()
        .Text = HexText
        GetBinaryFromHexText = .nodeTypedValue
    End With
End Function

'
' --- Test ---
'

Private Sub Test_HexAndBase64()
    Dim HexText
    HexText = GetTestHexText()
    
    Dim Binary
    Binary = GetBinaryFromHexText(HexText)
    Debug_Print_StringB Binary
    
    Dim Base64Text
    Base64Text = GetBase64TextFromBinary(Binary)
    Debug_Print Base64Text
    
    Binary = GetBinaryFromBase64Text(Base64Text)
    Debug_Print_StringB Binary
    
    HexText = GetHexTextFromBinary(Binary)
    Debug_Print HexText
End Sub

Private Function GetTestHexText()
    Dim HexText
    Dim Index
    For Index = 0 To 15
        HexText = HexText & "0" & Hex(Index)
    Next
    For Index = 16 To 255
        HexText = HexText & Hex(Index)
    Next
    GetTestHexText = HexText
End Function

Private Sub Debug_Print_StringB(StringB)
    Dim Text
    Dim Index1
    Dim Index2
    For Index1 = 1 To LenB(StringB) Step 16
        For Index2 = Index1 To MinL(Index1 + 15, LenB(StringB))
            Text = _
                Text & _
                Right("0" & Hex(AscB(MidB(StringB, Index2, 1))), 2) & " "
        Next
        Text = Text & vbNewLine
    Next
    
    Debug_Print "-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --"
    Debug_Print Text
    Debug_Print "-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --"
End Sub

Private Function MinL(Value1, Value2)
    If Value1 < Value2 Then
        MinL = Value1
    Else
        MinL = Value2
    End If
End Function

Private Sub Debug_Print(Str)
    WScript.Echo Str
End Sub