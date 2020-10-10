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
' VBA/VB.NET Collection Like Class
'
' VBA - Collection object
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object
'
' VB.NET - Collection class
' https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.collection?view=netcore-3.1
'
' Scripting - Dictionary object
' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/x4k5wbx4(v=vs.84)
'

Class Collection
    
    Private dicItemFromIndex ' key: index, value: item
    Private dicKeyFromIndex  ' key: index, value: key
    Private dicIndexFromKey  ' key: key,   value: index
    
    Private Sub Class_Initialize
        Set dicItemFromIndex = CreateObject("Scripting.Dictionary")
        Set dicKeyFromIndex = CreateObject("Scripting.Dictionary")
        Set dicIndexFromKey = CreateObject("Scripting.Dictionary")
    End Sub
    
    Private Sub Class_Terminate
        Set dicItemFromIndex = Nothing
        Set dicKeyFromIndex = Nothing
        Set dicIndexFromKey = Nothing
    End Sub
    
    '
    ' VBA - Collection - Item method
    ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/item-method-visual-basic-for-applications
    '
    ' VB.NET - Collection - Item property
    ' https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.collection.item?view=netcore-3.1
    '
    ' Scripting - Dictionary - Item property
    ' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/84k9x471(v=vs.84)
    '
    
    Public Default Property Get Item(index_or_key)
        Dim index
        index = GetIndex(index_or_key)
        
        If IsObject(dicItemFromIndex.Item(index)) Then
            Set Item = dicItemFromIndex.Item(index)
        Else
            Item = dicItemFromIndex.Item(index)
        End If
    End Property
    
    '
    ' VBA - Collection - Count property
    ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/count-property-visual-basic-for-applications
    '
    ' VB.NET - Collection - Count property
    ' https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.collection.count?view=netcore-3.1
    '
    ' Scripting - Dictionary - Count property
    ' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/5t9h9579(v=vs.84)
    '
    
    Public Property Get Count
        Count = dicItemFromIndex.Count
    End Property
    
    '
    ' VBA - Collection - Add method
    ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-visual-basic-for-applications
    '
    ' VB.NET - Collection - Add method
    ' https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.collection.add?view=netcore-3.1
    '
    ' Scripting - Dictionary - Add method
    ' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/5h92h863(v=vs.84)
    '
    
    Public Sub PushBack(item_)
        Add item_, Empty, Empty, Empty
    End Sub
    
    Public Sub AddBefore(item_, before)
        Add item_, Empty, before, Empty
    End Sub
    
    Public Sub AddAfter(item_, after)
        Add item_, Empty, Empty, after
    End Sub
    
    Public Sub AddWithKey(item_, key)
        Add item_, key, Empty, Empty
    End Sub
    
    Public Sub AddWithKeyBefore(item_, key, before)
        Add item_, key, before, Empty
    End Sub
    
    Public Sub AddWithKeyAfter(item_, key, after)
        Add item_, key, Empty, after
    End Sub
    
    Public Sub Add(item_, key, before, after)
        Dim index
        index = GetIndexForAdd(before, after)
        
        dicItemFromIndex_Add index, item_
        dicKeyFromIndex_Add index, key
        dicIndexFromKey_Add key, index
    End Sub
    
    Private Function GetIndexForAdd(before, after)
        If IsIndex(before) Then
            GetIndexForAdd = before
        ElseIf IsKey(before) Then
            GetIndexForAdd = dicIndexFromKey.Item(before)
        ElseIf IsIndex(after) Then
            GetIndexForAdd = after + 1
        ElseIf IsKey(after) Then
            GetIndexForAdd = dicIndexFromKey.Item(after) + 1
        Else
            GetIndexForAdd = dicItemFromIndex.Count + 1
        End If
    End Function
    
    Private Sub dicItemFromIndex_Add(index, item_)
        Dim i
        For i = dicItemFromIndex.Count To index Step -1
            dicItemFromIndex.Key(i) = i + 1
        Next
        
        dicItemFromIndex.Add index, item_
    End Sub
    
    Private Sub dicKeyFromIndex_Add(index, key)
        Dim i
        For i = dicKeyFromIndex.Count To index Step -1
            dicKeyFromIndex.Key(i) = i + 1
        Next
        
        dicKeyFromIndex.Add index, key
    End Sub
    
    Private Sub dicIndexFromKey_Add(key, index)
        Dim i
        For i = index + 1 To dicKeyFromIndex.Count
            Dim k
            k = dicKeyFromIndex.Item(i)
            If IsKey(k) Then
                dicIndexFromKey.Item(k) = i
            End If
        Next
        
        If IsKey(key) Then
            dicIndexFromKey.Add key, index
        End If
    End Sub
    
    '
    ' VBA - Collection - Remove method
    ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/remove-method-visual-basic-for-applications
    '
    ' VB.NET - Collection - Remove method
    ' https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.collection.remove?view=netcore-3.1
    '
    ' Scripting - Dictionary - Remove method
    ' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/ywyayk03(v=vs.84)
    '
    ' Scripting - Dictionary - Key property
    ' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/1ex01tte(v=vs.84)
    '
    
    Public Sub Remove(index_or_key)
        Dim index
        index = GetIndex(index_or_key)
        
        Dim key
        key = GetKey(index_or_key)
        
        dicItemFromIndex_Remove index
        dicKeyFromIndex_Remove index
        dicIndexFromKey_Remove index, key
    End Sub
    
    Private Sub dicItemFromIndex_Remove(index)
        dicItemFromIndex.Remove index
        
        Dim i
        For i = index To dicItemFromIndex.Count
            dicItemFromIndex.Key(i + 1) = i
        Next
    End Sub
    
    Private Sub dicKeyFromIndex_Remove(index)
        dicKeyFromIndex.Remove index
        
        Dim i
        For i = index To dicKeyFromIndex.Count
            dicKeyFromIndex.Key(i + 1) = i
        Next
    End Sub
    
    Private Sub dicIndexFromKey_Remove(index, key)
        If IsKey(key) Then
            dicIndexFromKey.Remove key
        End If
        
        Dim i
        For i = index To dicKeyFromIndex.Count
            Dim k
            k = dicKeyFromIndex.Item(i)
            If IsKey(k) Then
                dicIndexFromKey.Item(k) = i
            End If
        Next
    End Sub
    
    '
    ' VB.NET - Collection - Clear method
    ' https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.collection.clear?view=netcore-3.1
    '
    ' Scripting - Dictionary - RemoveAll method
    ' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/45731e2w(v=vs.84)
    '
    
    Public Sub Clear
        dicItemFromIndex.RemoveAll
        dicKeyFromIndex.RemoveAll
        dicIndexFromKey.RemoveAll
    End Sub
    
    '
    ' VB.NET - Collection - Contains method
    ' https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.collection.contains?view=netcore-3.1
    '
    ' Scripting - Dictionary - Exists method
    ' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/57hdf10z(v=vs.84)
    '
    
    Public Function Contains(key)
        Contains = dicIndexFromKey.Exists(key)
    End Function
    
    '
    '
    '
    
    Private Function GetIndex(index_or_key)
        If IsIndex(index_or_key) Then
            GetIndex = index_or_key
        ElseIf IsKey(index_or_key) Then
            GetIndex = dicIndexFromKey.Item(index_or_key)
        End If
    End Function
    
    Private Function GetKey(index_or_key)
        If IsIndex(index_or_key) Then
            GetKey = dicKeyFromIndex.Item(index_or_key)
        ElseIf IsKey(index_or_key) Then
            GetKey = index_or_key
        End If
    End Function
    
    Private Function IsIndex(index_or_key)
        IsIndex = Not IsEmpty(index_or_key) And IsNumeric(index_or_key)
    End Function
    
    Private Function IsKey(index_or_key)
        If IsString(index_or_key) Then
            IsKey = (index_or_key <> "")
        Else
            IsKey = False
        End If
    End Function
    
    Private Function IsString(index_or_key)
        IsString = (TypeName(index_or_key) = "String")
    End Function
    
End Class
