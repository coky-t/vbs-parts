'Attribute VB_Name = "Test_MFileSystem"
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
' --- Test ---
'

Public Sub Test_TextFileW()
    Test_TextFileW_Core _
        "testw.txt", _
        "WriteTextFileW" & vbNewLine, _
        "AppendTextFileW" & vbNewLine
End Sub

Public Sub Test_TextFileA()
    Test_TextFileA_Core _
        "testa.txt", _
        "WriteTextFileA" & vbNewLine, _
        "AppendTextFileA" & vbNewLine
End Sub

'
' --- Test Core ---
'

Public Sub Test_TextFileW_Core( _
    FileName, _
    Text1, _
    Text2)
    
    Dim Text
    
    MFileSystem_WriteTextFileW FileName, Text1
    Text = MFileSystem_ReadTextFileW(FileName)
    Debug_Print Text
    
    MFileSystem_AppendTextFileW FileName, Text2
    Text = MFileSystem_ReadTextFileW(FileName)
    Debug_Print Text
End Sub

Public Sub Test_TextFileA_Core( _
    FileName, _
    Text1, _
    Text2)
    
    Dim Text
    
    MFileSystem_WriteTextFileA FileName, Text1
    Text = MFileSystem_ReadTextFileA(FileName)
    Debug_Print Text
    
    MFileSystem_AppendTextFileA FileName, Text2
    Text = MFileSystem_ReadTextFileA(FileName)
    Debug_Print Text
End Sub
