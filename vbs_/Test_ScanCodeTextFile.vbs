'Attribute VB_Name = "Test_ScanCodeTextFile"
Option Explicit

'
' Copyright (c) 2025 Koki Takeyama
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

Sub Test_SaveScanCodeTextFile()
    Dim OutputFilePath
    OutputFilePath = "C:\work\data\scancode-licensedb-text.txt"
    
    ' https://github.com/aboutcode-org/scancode-licensedb/tree/main/docs
    Dim ScanCodeTextDirPath
    ScanCodeTextDirPath = "C:\work\data\scancode-licensedb\docs"
    
    Test_SaveScanCodeTextFile_Core _
        OutputFilePath, ScanCodeTextDirPath
End Sub

'
' --- Test Core ---
'

Sub Test_SaveScanCodeTextFile_Core( _
    OutputFilePath, DirPath)
    
    Dim OutputText
    
    Dim Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File
    For Each File In Folder.Files
        If GetFileSystemObject().GetExtensionName(File.Path) = "json" Then
        If Not File.Name = "index.json" Then
            Debug_Print File.Name
            
            Dim FileText
            FileText = MADODBStream_ReadTextFileUTF8(File.Path)
            'Debug_Print FileText
            
            Dim JsonObj
            Set JsonObj = ParseJsonText(FileText)
            
            Dim JsonKeys
            Set JsonKeys = GetJsonKeys(JsonObj)
            
            Dim JsonKeysLength
            JsonKeysLength = GetJsonKeysLength(JsonKeys)
            
            If JsonKeyExists(JsonKeys, JsonKeysLength, "text") Then
                Dim Key
                Dim Text
                Key = GetJsonItemValue(JsonObj, "key")
                Text = GetJsonItemValue(JsonObj, "text")
                'Debug_Print "--- key: " & Key
                'Debug_Print "--- text: " & Text
            
                OutputText = OutputText & _
                    "<pre name=""" & Key & """>" & _
                    ReplaceChars(Text) & "</pre>" & vbCrLf
            End If
        End If
        End If
    Next
    
    MADODBStream_WriteTextFileUTF8 OutputFilePath, OutputText, True
    Debug_Print "... Done."
End Sub

Private Function JsonKeyExists( _
    JsonKeys, JsonKeysLength, Key)
    Dim Index
    For Index = 0 To JsonKeysLength - 1
        If CStr(GetJsonKeysItem(JsonKeys, Index)) = Key Then
            JsonKeyExists = True
            Exit Function
        End If
    Next
    JsonKeyExists = False
End Function

Private Function ReplaceChars(Str)
    Dim Temp
    Temp = Str
    
    Temp = Replace(Temp, "&", "&amp;")
    Temp = Replace(Temp, ">", "&gt;")
    Temp = Replace(Temp, "<", "&lt;")
    'Temp = Replace(Temp, vbCrLf, "<br>")
    'Temp = Replace(Temp, vbLf, "<br>")
    
    ReplaceChars = Temp
End Function
