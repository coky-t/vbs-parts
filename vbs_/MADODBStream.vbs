'Attribute VB_Name = "MADODBStream"
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
' Microsoft ActiveX Data Objects X.X Library
' - ADODB.Stream
'

'
' === TextFile ===
'

'
' ReadTextFileW
' - Reads an entire file and returns the resulting string (Unicode).
'
' ReadTextFileA
' - Reads an entire file and returns the resulting string (ASCII).
'
' ReadTextFileUTF8
' - Reads an entire file and returns the resulting string (UTF-8).
'

'
' FileName:
'   Required. A String value that contains the name of a file.
'   FileName can contain any valid path and name in UNC format.
'

Public Function MADODBStream_ReadTextFileW(FileName)
    MADODBStream_ReadTextFileW = _
        MADODBStream_ReadTextFile(FileName, "unicode")
End Function

Public Function MADODBStream_ReadTextFileA(FileName)
    MADODBStream_ReadTextFileA = _
        MADODBStream_ReadTextFile(FileName, "iso-8859-1")
End Function

Public Function MADODBStream_ReadTextFileUTF8(FileName)
    MADODBStream_ReadTextFileUTF8 = _
        MADODBStream_ReadTextFile(FileName, "utf-8")
End Function

Public Function MADODBStream_ReadTextFile( _
    FileName, _
    Charset)
    
    If FileName = "" Then Exit Function
    
    MADODBStream_ReadTextFile = LoadFromFileAndReadText(FileName, Charset)
End Function

'
' WriteTextFileW
' - Writes a specified string (Unicode) to a file.
'
' WriteTextFileA
' - Writes a specified string (ASCII) to a file.
'
' WriteTextFileUTF8
' - Writes a specified string (UTF-8) to a file.
'

'
' FileName:
'   Required. A String value that contains the fully-qualified name of
'   the file to which the contents of the Stream will be saved.
'   You can save to any valid local location, or any location you have
'   access to via a UNC value.
'
' Text:
'   Required. A String value that contains the text in characters to be
'   written.
'

Public Sub MADODBStream_WriteTextFileW(FileName, Text)
    MADODBStream_WriteTextFile FileName, Text, "unicode"
End Sub

Public Sub MADODBStream_WriteTextFileA(FileName, Text)
    MADODBStream_WriteTextFile FileName, Text, "iso-8859-1"
End Sub

Public Sub MADODBStream_WriteTextFileUTF8(FileName, Text)
    MADODBStream_WriteTextFile FileName, Text, "utf-8"
End Sub

Public Sub MADODBStream_WriteTextFile( _
    FileName, _
    Text, _
    Charset)
    
    If FileName = "" Then Exit Sub
    
    WriteTextAndSaveToFile FileName, Text, Charset
End Sub

'
' --- TextFile ---
'

'
' LoadFromFileAndReadText
' - Reads an entire file and returns the resulting string.
'

'
' FileName:
'   Required. A String value that contains the name of a file.
'   FileName can contain any valid path and name in UNC format.
'
' Charset:
'   Optional. A String value that specifies the character set into
'   which the contents of the Stream will be translated.
'   The default value is Unicode.
'   Allowed values are typical strings passed over the interface as
'   Internet character set names (for example, "iso-8859-1", "Windows-1252",
'   and so on).
'   For a list of the character set names that are known by a system,
'   see the subkeys of HKEY_CLASSES_ROOT\MIME\Database\Charset
'   in the Windows Registry.
'

Public Function LoadFromFileAndReadText( _
    FileName, _
    Charset)
    
    On Error Resume Next
    
    With CreateObject("ADODB.Stream")
        If Charset <> "" Then .Charset = Charset
        .Open
        .LoadFromFile FileName
        LoadFromFileAndReadText = .ReadText
        .Close
    End With
End Function

'
' WriteTextAndSaveToFile
' - Writes a specified string to a file.
'

'
' FileName:
'   Required. A String value that contains the fully-qualified name of
'   the file to which the contents of the Stream will be saved.
'   You can save to any valid local location, or any location you have
'   access to via a UNC value.
'
' Text:
'   Required. A String value that contains the text in characters to be
'   written.
'
' Charset:
'   Optional. A String value that specifies the character set into
'   which the contents of the Stream will be translated.
'   The default value is Unicode.
'   Allowed values are typical strings passed over the interface as
'   Internet character set names (for example, "iso-8859-1", "Windows-1252",
'   and so on).
'   For a list of the character set names that are known by a system,
'   see the subkeys of HKEY_CLASSES_ROOT\MIME\Database\Charset
'   in the Windows Registry.
'

Public Sub WriteTextAndSaveToFile( _
    FileName, _
    Text, _
    Charset)
    
    On Error Resume Next
    
    With CreateObject("ADODB.Stream")
        If Charset <> "" Then .Charset = Charset
        .Open
        .WriteText Text
        .SaveToFile FileName, 2 'ADODB.adSaveCreateOverWrite
        .Close
    End With
End Sub

'
' === BinaryFile ===
'

'
' ReadBinaryFile
' - Reads an entire file and returns the resulting data.
'

'
' FileName:
'   Required. A String value that contains the name of a file.
'   FileName can contain any valid path and name in UNC format.
'

Public Function MADODBStream_ReadBinaryFile(FileName)
    If FileName = "" Then Exit Function
    MADODBStream_ReadBinaryFile = LoadFromFileAndRead(FileName)
End Function

'
' WriteBinaryFile
' - Writes a binary data to a file.
'

'
' FileName:
'   Required. A String value that contains the fully-qualified name of
'   the file to which the contents of the Stream will be saved.
'   You can save to any valid local location, or any location you have
'   access to via a UNC value.
'
' Buffer:
'   Required. A Variant that contains an array of bytes to be written.
'

Public Sub MADODBStream_WriteBinaryFile(FileName, Buffer)
    If FileName = "" Then Exit Sub
    'WriteAndSaveToFile FileName, Buffer
    
    Dim Buf
    Dim Index
    For Index = 1 to LenB(Buffer)
        Buf = Buf & ChrW(AscB(MidB(Buffer, Index, 1)))
    Next
    MADODBStream_WriteTextFileA FileName, Buf
End Sub

'
' --- BinaryFile ---
'

'
' LoadFromFileAndRead
' - Reads an entire file and returns the resulting data.
'

'
' FileName:
'   Required. A String value that contains the name of a file.
'   FileName can contain any valid path and name in UNC format.
'

Public Function LoadFromFileAndRead(FileName)
    On Error Resume Next
    
    With CreateObject("ADODB.Stream")
        .Type = 1 'ADODB.adTypeBinary
        .Open
        .LoadFromFile FileName
        LoadFromFileAndRead = .Read
        .Close
    End With
End Function

'
' WriteAndSaveToFile
' - Writes a binary data to a file.
'

'
' FileName:
'   Required. A String value that contains the fully-qualified name of
'   the file to which the contents of the Stream will be saved.
'   You can save to any valid local location, or any location you have
'   access to via a UNC value.
'
' Buffer:
'   Required. A Variant that contains an array of bytes to be written.
'

Public Sub WriteAndSaveToFile( _
    FileName, _
    Buffer)
    
    On Error Resume Next
    
    With CreateObject("ADODB.Stream")
        .Type = 1 'ADODB.adTypeBinary
        .Open
        .Write Buffer
        .SaveToFile FileName, 2 'ADODB.adSaveCreateOverWrite
        .Close
    End With
End Sub

'
' --- Test ---
'

Private Sub Test_MADODBStream_TextFileW()
    Dim FileName
    FileName = MADODBStream_GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Text
    
    Text = "WriteTextFileW" & vbNewLine
    MADODBStream_WriteTextFileW FileName, Text
    Text = MADODBStream_ReadTextFileW(FileName)
    MADODBStream_Debug_Print Text
End Sub

Private Sub Test_MADODBStream_TextFileA()
    Dim FileName
    FileName = MADODBStream_GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Text
    
    Text = "WriteTextFileA" & vbNewLine
    MADODBStream_WriteTextFileA FileName, Text
    Text = MADODBStream_ReadTextFileA(FileName)
    MADODBStream_Debug_Print Text
End Sub

Private Sub Test_MADODBStream_TextFileUTF8()
    Dim FileName
    FileName = MADODBStream_GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Text
    
    Text = "WriteTextFileUTF8" & vbNewLine
    MADODBStream_WriteTextFileUTF8 FileName, Text
    Text = MADODBStream_ReadTextFileUTF8(FileName)
    MADODBStream_Debug_Print Text
End Sub

Private Sub Test_MADODBStream_BinaryFile()
    Dim FileName
    FileName = MADODBStream_GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Buffer
    Dim Index
    For Index = 0 To 255
        Buffer = Buffer & ChrB(Index)
    Next
    
    MADODBStream_WriteBinaryFile FileName, Buffer
    
    Dim Data
    Data = MADODBStream_ReadBinaryFile(FileName)
    
    Dim Text
    Dim Index1
    For Index1 = 1 To LenB(Data) Step 16
        Dim Index2
        For Index2 = Index1 To Index1 + 15
            Text = _
                Text & Right("0" & Hex(AscB(MidB(Data, Index2, 1))), 2) & " "
        Next
        Text = Text & vbNewLine
    Next
    
    MADODBStream_Debug_Print Text
End Sub

Private Function MADODBStream_GetSaveAsFileName()
    MADODBStream_GetSaveAsFileName = InputBox("GetSaveAsFileName")
End Function

Private Sub MADODBStream_Debug_Print(Str)
    WScript.Echo Str
End Sub
