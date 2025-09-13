'Attribute VB_Name = "MHTMLDocument"
Option Explicit

'
' Copyright (c) 2020,2022 Koki Takeyama
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
' Microsoft HTML Object Library
' - MSHTML.HTMLDocument
' - htmlfile
'

'#Const UseCallByName = False

Private HTMLDocument

'
' --- MSHTML.HTMLDocument ---
'

'
' GetHTMLDocumentForJson
' - Returns a MSHTML.HTMLDocument object.
'

Public Function GetHTMLDocumentForJson()
    'Static HTMLDocument
    If IsEmpty(HTMLDocument) Then
        Set HTMLDocument = CreateObject("htmlfile")
        With HTMLDocument
            .write "<meta http-equiv=""x-ua-compatible"" content=""IE=11""/>"
            
            '
            ' JsonText to JsonObject
            '
            .write _
                "<script>document.ParseJsonText=function (JsonText) { " & _
                "return eval('(' + JsonText + ')'); }</script>"
                
            '
            ' JsonObject to JsonItems
            '
            .write _
                "<script>document.GetKeys=function (JsonObj) { " & _
                "var keys = []; " & _
                "for (var key in JsonObj) { keys.push(key); } " & _
                "return keys; }</script>"
'#If UseCallByName Then
'            ' nop
'#Else
            .write _
                "<script>document.GetItem=function (obj, i) { " & _
                "return obj[i]; }</script>"
'            .write _
'                "<script>document.GetLength=function (obj) { " & _
'                "return obj.length; }</script>"
'#End If
            .write _
                "<script>document.IsJsonArray=function (obj) { " & _
                "return Object.prototype.toString.call(obj) === " & _
                "'[object Array]'; }</script>"
                
            '
            ' JsonObject to JsonText
            '
            .write _
                "<script>document.GetJsonText=function (obj) { " & _
                "return document.parentWindow.JSON.stringify(obj); }</script>"
                
            '
            ' VbsObject to JsonObject
            '
            .write _
                "<script>document.NewJsonArray=function () { " & _
                "return new Array; }</script>"
            .write _
                "<script>document.AddJsonArrayItem=" & _
                "function (arr, item) { " & _
                "arr.push(item);" & _
                "return arr; }</script>"
            .write _
                "<script>document.NewJsonDictionary=function () { " & _
                "var dic = {}; " & _
                "return dic; }</script>"
            .write _
                "<script>document.AddJsonDictionaryItem=" & _
                "function (dic, key, item) { " & _
                "dic[key] = item;" & _
                "return dic; }</script>"
        End With
    End If
    Set GetHTMLDocumentForJson = HTMLDocument
End Function

'
' === Json ===
'

'
' ParseJsonText
' - Returns a JSON object.
'

'
' JsonText:
'   Required. String expression that identifies JSON data.
'

Public Function ParseJsonText(JsonText)
    On Error Resume Next
    
'#If UseCallByName Then
'    Set ParseJsonText = _
'        CallByName( _
'            GetHTMLDocumentForJson(), _
'            "ParseJsonText", _
'            VbMethod, _
'            JsonText)
'#Else
    Set ParseJsonText = GetHTMLDocumentForJson().ParseJsonText(JsonText)
'#End If
End Function

'
' GetJsonKeys
' - Returns an array containing all existing keys in a JSON object.
'

'
' JsonObject:
'   Required. The name of a JSON object
'

Public Function GetJsonKeys(JsonObject)
    On Error Resume Next
    
'#If UseCallByName Then
'    Set GetJsonKeys = _
'        CallByName( _
'            GetHTMLDocumentForJson(), _
'            "GetKeys", _
'            VbMethod, _
'            JsonObject)
'#Else
    Set GetJsonKeys = GetHTMLDocumentForJson().GetKeys(JsonObject)
'#End If
End Function

'
' GetJsonKeysLength
' - Returns the number of elements in a JSON keys array.
'

'
' JsonKeys:
'   Required. The name of a JSON keys array object
'

Public Function GetJsonKeysLength(JsonKeys)
'#If UseCallByName Then
'    GetJsonKeysLength = CallByName(JsonKeys, "length", VbGet)
'#Else
'    GetJsonKeysLength = GetHTMLDocumentForJson().GetLength(JsonKeys)
    GetJsonKeysLength = JsonKeys.Length
'#End If
End Function

'
' GetJsonKeysItem
' - Returns an item of JSON  object.
'

'
' JsonKeys:
'   Required. The name of a JSON keys array object
'
' Index:
'   Required. Index associated with the item being retrieved.
'

Public Function GetJsonKeysItem(JsonKeys, Index)
'#If UseCallByName Then
'    GetJsonKeysItem = CallByName(JsonKeys, Index, VbGet)
'#Else
    GetJsonKeysItem = GetHTMLDocumentForJson().GetItem(JsonKeys, Index)
'#End If
End Function

'
' IsJsonItemObject
' - Returns a Boolean value indicating whether an item of JSON object
'   represents an object variable.
'

'
' JsonObject:
'   Required. The name of a JSON object
'
' Key:
'   Required. Key associated with the item being retrieved.
'

Public Function IsJsonItemObject( _
    JsonObject, _
    Key)
    
'#If UseCallByName Then
'    IsJsonItemObject = IsObject(CallByName(JsonObject, Key, VbGet))
'#Else
    IsJsonItemObject = _
        IsObject(GetHTMLDocumentForJson().GetItem(JsonObject, Key))
'#End If
End Function

'
' GetJsonItemValue
' - Returns an item of JSON object.
'

'
' JsonObject:
'   Required. The name of a JSON object
'
' Key:
'   Required. Key associated with the item being retrieved.
'

Public Function GetJsonItemValue( _
    JsonObject, _
    Key)
    
'#If UseCallByName Then
'    GetJsonItemValue = CallByName(JsonObject, Key, VbGet)
'#Else
    GetJsonItemValue = GetHTMLDocumentForJson().GetItem(JsonObject, Key)
'#End If
End Function

'
' GetJsonItemObject
' - Returns an item of JSON object.
'

'
' JsonObject:
'   Required. The name of a JSON object
'
' Key:
'   Required. Key associated with the item being retrieved.
'

Public Function GetJsonItemObject( _
    JsonObject, _
    Key)
    
'#If UseCallByName Then
'    Set GetJsonItemObject = CallByName(JsonObject, Key, VbGet)
'#Else
    Set GetJsonItemObject = GetHTMLDocumentForJson().GetItem(JsonObject, Key)
'#End If
End Function

'
' IsJsonArray
' - Returns a Boolean value indicating whether JSON object is an instance of
'   Array object.
'

'
' JsonObject:
'   Required. The name of a JSON object
'

Public Function IsJsonArray(JsonObject)
'#If UseCallByName Then
'    IsJsonArray = _
'        CallByName( _
'            GetHTMLDocumentForJson(), _
'            "IsJsonArray", _
'            VbMethod, _
'            JsonObject)
'#Else
    IsJsonArray = GetHTMLDocumentForJson().IsJsonArray(JsonObject)
'#End If
End Function

'
' GetJsonText
' - Returns string expression that identifies JSON data.
'

'
' JsonObject:
'   Required. The name of a JSON object
'

Public Function GetJsonText(JsonObject)
'#If UseCallByName Then
'    GetJsonText = _
'        CallByName( _
'            GetHTMLDocumentForJson(), _
'            "GetJsonText", _
'            VbMethod, _
'            JsonObject)
'#Else
    GetJsonText = GetHTMLDocumentForJson().GetJsonText(JsonObject)
'#End If
End Function

'
' VbsObject to JsonObject
'

Public Function NewJsonArray()
'#If UseCallByName Then
'    Set NewJsonArray = _
'        CallByName( _
'            GetHTMLDocumentForJson(), _
'            "NewJsonArray", _
'            VbMethod)
'#Else
    Set NewJsonArray = GetHTMLDocumentForJson().NewJsonArray()
'#End If
End Function

Public Function AddJsonArrayItem(JsonArray, Item)
'#If UseCallByName Then
'    Set AddJsonArrayItem = _
'        CallByName( _
'            GetHTMLDocumentForJson(), _
'            "AddJsonArrayItem", _
'            VbMethod, _
'            JsonArray, _
'            Item)
'#Else
    Set AddJsonArrayItem = _
        GetHTMLDocumentForJson().AddJsonArrayItem(JsonArray, Item)
'#End If
End Function

Public Function NewJsonDictionary()
'#If UseCallByName Then
'    Set NewJsonDictionary = _
'        CallByName( _
'            GetHTMLDocumentForJson(), _
'            "NewJsonDictionary", _
'            VbMethod)
'#Else
    Set NewJsonDictionary = GetHTMLDocumentForJson().NewJsonDictionary()
'#End If
End Function

Public Function AddJsonDictionaryItem(JsonDictionary, Key, Item)
'#If UseCallByName Then
'    Set AddJsonDictionaryItem = _
'        CallByName( _
'            GetHTMLDocumentForJson(), _
'            "AddJsonDictionaryItem", _
'            VbMethod, _
'            JsonDictionary, _
'            Key, _
'            Item)
'#Else
    Set AddJsonDictionaryItem = _
        GetHTMLDocumentForJson().AddJsonDictionaryItem( _
            JsonDictionary, Key, Item)
'#End If
End Function

'
' JsonObject to VbsObject
'

Public Function GetVbsObjectFromJsonObject(JsonObject)
    If IsJsonArray(JsonObject) Then
        Set GetVbsObjectFromJsonObject = _
            GetVbsCollectionFromJsonArray(JsonObject)
    Else
        Set GetVbsObjectFromJsonObject = _
            GetVbsDictionaryFromJsonDictionary(JsonObject)
    End If
End Function

Public Function GetVbsCollectionFromJsonArray(JsonArray)
    Dim VbsCollection
    Set VbsCollection = New Collection
    
    Dim Keys
    Set Keys = GetJsonKeys(JsonArray)
    
    Dim KeysLength
    KeysLength = GetJsonKeysLength(Keys)
    
    Dim Index
    Dim Key
    For Index = 0 To KeysLength - 1
        Key = GetJsonKeysItem(Keys, Index)
        If IsJsonItemObject(JsonArray, Key) Then
            Dim JsonItemObject
            Set JsonItemObject = GetJsonItemObject(JsonArray, Key)
            
            Dim VbsItemObject
            Set VbsItemObject = GetVbsObjectFromJsonObject(JsonItemObject)
            
            VbsCollection.PushBack VbsItemObject
        Else
            VbsCollection.PushBack GetJsonItemValue(JsonArray, Key)
        End If
    Next
    
    Set GetVbsCollectionFromJsonArray = VbsCollection
End Function

Public Function GetVbsDictionaryFromJsonDictionary(JsonDictionary)
    Dim VbsDictionary
    Set VbsDictionary = CreateObject("Scripting.Dictionary")
    
    Dim Keys
    Set Keys = GetJsonKeys(JsonDictionary)
    
    Dim KeysLength
    KeysLength = GetJsonKeysLength(Keys)
    
    Dim Index
    Dim Key
    For Index = 0 To KeysLength - 1
        Key = GetJsonKeysItem(Keys, Index)
        If IsJsonItemObject(JsonDictionary, Key) Then
            Dim JsonItemObject
            Set JsonItemObject = GetJsonItemObject(JsonDictionary, Key)
            
            Dim VbsItemObject
            Set VbsItemObject = GetVbsObjectFromJsonObject(JsonItemObject)
            
            VbsDictionary.Add Key, VbsItemObject
        Else
            VbsDictionary.Add Key, GetJsonItemValue(JsonDictionary, Key)
        End If
    Next
    
    Set GetVbsDictionaryFromJsonDictionary = VbsDictionary
End Function

'
' VbsObject to JsonObject
'

Public Function GetJsonObjectFromVbsObject(VbsObject)
    Select Case TypeName(VbsObject)
    
    Case "Collection"
        Set GetJsonObjectFromVbsObject = _
            GetJsonArrayFromVbsCollection(VbsObject)
        
    Case "Dictionary"
        Set GetJsonObjectFromVbsObject = _
            GetJsonDictionaryFromVbsDictionary(VbsObject)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetJsonArrayFromVbsCollection(VbsCollection)
    Dim JsonArray
    Set JsonArray = NewJsonArray()
    
    Dim Index
    For Index = 1 To VbsCollection.Count
        If IsObject(VbsCollection.Item(Index)) Then
            AddJsonArrayItem JsonArray, _
                GetJsonObjectFromVbsObject(VbsCollection.Item(Index))
        Else
            AddJsonArrayItem JsonArray, VbsCollection.Item(Index)
        End If
    Next
    
    Set GetJsonArrayFromVbsCollection = JsonArray
End Function

Public Function GetJsonDictionaryFromVbsDictionary(VbsDictionary)
    Dim JsonDictionary
    Set JsonDictionary = NewJsonDictionary()
    
    Dim VbsDicKeys
    VbsDicKeys = VbsDictionary.Keys
    
    Dim Index
    For Index = LBound(VbsDicKeys) To UBound(VbsDicKeys)
        Dim Key
        Key = VbsDicKeys(Index)
        If IsObject(VbsDictionary.Item(Key)) Then
            AddJsonDictionaryItem JsonDictionary, Key, _
                GetJsonObjectFromVbsObject(VbsDictionary.Item(Key))
        Else
            AddJsonDictionaryItem JsonDictionary, Key, VbsDictionary.Item(Key)
        End If
    Next
    
    Set GetJsonDictionaryFromVbsDictionary = JsonDictionary
End Function

'
' JsonText to VbsObject
'

Public Function GetVbsObjectFromJsonText(JsonText)
    Dim JsonObject
    Set JsonObject = ParseJsonText(JsonText)
    Set GetVbsObjectFromJsonText = GetVbsObjectFromJsonObject(JsonObject)
End Function

'
' VbsObject to JsonText
'

Public Function GetJsonTextFromVbsObject(VbsObject)
    Dim JsonObject
    Set JsonObject = GetJsonObjectFromVbsObject(VbsObject)
    GetJsonTextFromVbsObject = GetJsonText(JsonObject)
End Function
