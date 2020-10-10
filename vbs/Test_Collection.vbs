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

Public Sub Test_CollectionA1 ' PushBack, Clear
    Dim coll
    Set coll = New Collection
    Test_Collection_PushBack_Core coll, "item1"
    Test_Collection_PushBack_Core coll, "item2"
    Test_Collection_PushBack_Core coll, "item3"
    Test_Collection_PushBack_Core coll, "item4"
    Test_Collection_PushBack_Core coll, "item5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_Clear_Core coll
    Test_Collection_Count_Core coll
End Sub

Public Sub Test_CollectionA2_1 ' Remove first index's item
    Dim coll
    Set coll = New Collection
    Test_Collection_PushBack_Core coll, "item1"
    Test_Collection_PushBack_Core coll, "item2"
    Test_Collection_PushBack_Core coll, "item3"
    Test_Collection_PushBack_Core coll, "item4"
    Test_Collection_PushBack_Core coll, "item5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_Remove_Core coll, 1
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionA2_2 ' Remove third index's item
    Dim coll
    Set coll = New Collection
    Test_Collection_PushBack_Core coll, "item1"
    Test_Collection_PushBack_Core coll, "item2"
    Test_Collection_PushBack_Core coll, "item3"
    Test_Collection_PushBack_Core coll, "item4"
    Test_Collection_PushBack_Core coll, "item5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_Remove_Core coll, 3
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionA2_3 ' Remove last index's item
    Dim coll
    Set coll = New Collection
    Test_Collection_PushBack_Core coll, "item1"
    Test_Collection_PushBack_Core coll, "item2"
    Test_Collection_PushBack_Core coll, "item3"
    Test_Collection_PushBack_Core coll, "item4"
    Test_Collection_PushBack_Core coll, "item5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_Remove_Core coll, 5
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionA3_1 ' AddBefore first index's item
    Dim coll
    Set coll = New Collection
    Test_Collection_PushBack_Core coll, "item1"
    Test_Collection_PushBack_Core coll, "item2"
    Test_Collection_PushBack_Core coll, "item3"
    Test_Collection_PushBack_Core coll, "item4"
    Test_Collection_PushBack_Core coll, "item5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_AddBefore_Core coll, "itemX", 1
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionA3_2 ' AddBefore third index's item
    Dim coll
    Set coll = New Collection
    Test_Collection_PushBack_Core coll, "item1"
    Test_Collection_PushBack_Core coll, "item2"
    Test_Collection_PushBack_Core coll, "item3"
    Test_Collection_PushBack_Core coll, "item4"
    Test_Collection_PushBack_Core coll, "item5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_AddBefore_Core coll, "itemX", 3
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionA3_3 ' AddBefore last index's item
    Dim coll
    Set coll = New Collection
    Test_Collection_PushBack_Core coll, "item1"
    Test_Collection_PushBack_Core coll, "item2"
    Test_Collection_PushBack_Core coll, "item3"
    Test_Collection_PushBack_Core coll, "item4"
    Test_Collection_PushBack_Core coll, "item5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_AddBefore_Core coll, "itemX", 5
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionA4_1 ' AddAfter first index's item
    Dim coll
    Set coll = New Collection
    Test_Collection_PushBack_Core coll, "item1"
    Test_Collection_PushBack_Core coll, "item2"
    Test_Collection_PushBack_Core coll, "item3"
    Test_Collection_PushBack_Core coll, "item4"
    Test_Collection_PushBack_Core coll, "item5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_AddAfter_Core coll, "itemX", 1
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionA4_2 ' AddAfter third index's item
    Dim coll
    Set coll = New Collection
    Test_Collection_PushBack_Core coll, "item1"
    Test_Collection_PushBack_Core coll, "item2"
    Test_Collection_PushBack_Core coll, "item3"
    Test_Collection_PushBack_Core coll, "item4"
    Test_Collection_PushBack_Core coll, "item5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_AddAfter_Core coll, "itemX", 3
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionA4_3 ' AddAfter last index's item
    Dim coll
    Set coll = New Collection
    Test_Collection_PushBack_Core coll, "item1"
    Test_Collection_PushBack_Core coll, "item2"
    Test_Collection_PushBack_Core coll, "item3"
    Test_Collection_PushBack_Core coll, "item4"
    Test_Collection_PushBack_Core coll, "item5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_AddAfter_Core coll, "itemX", 5
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionB1 ' AddWithKey, Item(key), Clear
    Dim coll
    Set coll = New Collection
    Test_Collection_AddWithKey_Core coll, "item1", "key1"
    Test_Collection_AddWithKey_Core coll, "item2", "key2"
    Test_Collection_AddWithKey_Core coll, "item3", "key3"
    Test_Collection_AddWithKey_Core coll, "item4", "key4"
    Test_Collection_AddWithKey_Core coll, "item5", "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_Item_Core coll, "key1"
    Test_Collection_Item_Core coll, "key2"
    Test_Collection_Item_Core coll, "key3"
    Test_Collection_Item_Core coll, "key4"
    Test_Collection_Item_Core coll, "key5"
    
    Debug_Print_Separator
    
    Test_Collection_Clear_Core coll
    Test_Collection_Count_Core coll
End Sub

Public Sub Test_CollectionB2_1 ' Remove first key's Item, Contains
    Dim coll
    Set coll = New Collection
    Test_Collection_AddWithKey_Core coll, "item1", "key1"
    Test_Collection_AddWithKey_Core coll, "item2", "key2"
    Test_Collection_AddWithKey_Core coll, "item3", "key3"
    Test_Collection_AddWithKey_Core coll, "item4", "key4"
    Test_Collection_AddWithKey_Core coll, "item5", "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_Remove_Core coll, "key1"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_Contains_Core coll, "key1"
    Test_Collection_Contains_Core coll, "key2"
    Test_Collection_Contains_Core coll, "key3"
    Test_Collection_Contains_Core coll, "key4"
    Test_Collection_Contains_Core coll, "key5"
End Sub

Public Sub Test_CollectionB2_2 ' Remove third key's item, Contains
    Dim coll
    Set coll = New Collection
    Test_Collection_AddWithKey_Core coll, "item1", "key1"
    Test_Collection_AddWithKey_Core coll, "item2", "key2"
    Test_Collection_AddWithKey_Core coll, "item3", "key3"
    Test_Collection_AddWithKey_Core coll, "item4", "key4"
    Test_Collection_AddWithKey_Core coll, "item5", "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_Remove_Core coll, "key3"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_Contains_Core coll, "key1"
    Test_Collection_Contains_Core coll, "key2"
    Test_Collection_Contains_Core coll, "key3"
    Test_Collection_Contains_Core coll, "key4"
    Test_Collection_Contains_Core coll, "key5"
End Sub

Public Sub Test_CollectionB2_3 ' Remove last key's item, Contains
    Dim coll
    Set coll = New Collection
    Test_Collection_AddWithKey_Core coll, "item1", "key1"
    Test_Collection_AddWithKey_Core coll, "item2", "key2"
    Test_Collection_AddWithKey_Core coll, "item3", "key3"
    Test_Collection_AddWithKey_Core coll, "item4", "key4"
    Test_Collection_AddWithKey_Core coll, "item5", "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_Remove_Core coll, "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_Contains_Core coll, "key1"
    Test_Collection_Contains_Core coll, "key2"
    Test_Collection_Contains_Core coll, "key3"
    Test_Collection_Contains_Core coll, "key4"
    Test_Collection_Contains_Core coll, "key5"
End Sub

Public Sub Test_CollectionB3_1 ' AddWithKeyBefore first key's item
    Dim coll
    Set coll = New Collection
    Test_Collection_AddWithKey_Core coll, "item1", "key1"
    Test_Collection_AddWithKey_Core coll, "item2", "key2"
    Test_Collection_AddWithKey_Core coll, "item3", "key3"
    Test_Collection_AddWithKey_Core coll, "item4", "key4"
    Test_Collection_AddWithKey_Core coll, "item5", "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_AddWithKeyBefore_Core coll, "itemX", "keyX", "key1"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionB3_2 ' AddWithKeyBefore third key's item
    Dim coll
    Set coll = New Collection
    Test_Collection_AddWithKey_Core coll, "item1", "key1"
    Test_Collection_AddWithKey_Core coll, "item2", "key2"
    Test_Collection_AddWithKey_Core coll, "item3", "key3"
    Test_Collection_AddWithKey_Core coll, "item4", "key4"
    Test_Collection_AddWithKey_Core coll, "item5", "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_AddWithKeyBefore_Core coll, "itemX", "keyX", "key3"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionB3_3 ' AddWithKeyBefore last key's item
    Dim coll
    Set coll = New Collection
    Test_Collection_AddWithKey_Core coll, "item1", "key1"
    Test_Collection_AddWithKey_Core coll, "item2", "key2"
    Test_Collection_AddWithKey_Core coll, "item3", "key3"
    Test_Collection_AddWithKey_Core coll, "item4", "key4"
    Test_Collection_AddWithKey_Core coll, "item5", "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_AddWithKeyBefore_Core coll, "itemX", "keyX", "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionB4_1 ' AddWithKeyAfter first key's item
    Dim coll
    Set coll = New Collection
    Test_Collection_AddWithKey_Core coll, "item1", "key1"
    Test_Collection_AddWithKey_Core coll, "item2", "key2"
    Test_Collection_AddWithKey_Core coll, "item3", "key3"
    Test_Collection_AddWithKey_Core coll, "item4", "key4"
    Test_Collection_AddWithKey_Core coll, "item5", "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_AddWithKeyAfter_Core coll, "itemX", "keyX", "key1"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionB4_2 ' AddWithKeyAfter third key's item
    Dim coll
    Set coll = New Collection
    Test_Collection_AddWithKey_Core coll, "item1", "key1"
    Test_Collection_AddWithKey_Core coll, "item2", "key2"
    Test_Collection_AddWithKey_Core coll, "item3", "key3"
    Test_Collection_AddWithKey_Core coll, "item4", "key4"
    Test_Collection_AddWithKey_Core coll, "item5", "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_AddWithKeyAfter_Core coll, "itemX", "keyX", "key3"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Public Sub Test_CollectionB4_3 ' AddWithKeyAfter last key's item
    Dim coll
    Set coll = New Collection
    Test_Collection_AddWithKey_Core coll, "item1", "key1"
    Test_Collection_AddWithKey_Core coll, "item2", "key2"
    Test_Collection_AddWithKey_Core coll, "item3", "key3"
    Test_Collection_AddWithKey_Core coll, "item4", "key4"
    Test_Collection_AddWithKey_Core coll, "item5", "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
    
    Debug_Print_Separator
    
    Test_Collection_AddWithKeyAfter_Core coll, "itemX", "keyX", "key5"
    Test_Collection_Count_Core coll
    Debug_Print_Collection_Items coll
End Sub

Class Dummy
    Public elem1
End Class

Public Sub Test_CollectionC1 ' PushBack object
    Dim coll
    Set coll = New Collection
    
    Dim dummy1
    Dim dummy2
    Dim dummy3
    Dim dummy4
    Dim dummy5
    Set dummy1 = New Dummy
    Set dummy2 = New Dummy
    Set dummy3 = New Dummy
    Set dummy4 = New Dummy
    Set dummy5 = New Dummy
    dummy1.elem1 = "dummy1-elem1"
    dummy2.elem1 = "dummy2-elem1"
    dummy3.elem1 = "dummy3-elem1"
    dummy4.elem1 = "dummy4-elem1"
    dummy5.elem1 = "dummy5-elem1"
    Test_Collection_PushBack_Dummy_Core coll, dummy1
    Test_Collection_PushBack_Dummy_Core coll, dummy2
    Test_Collection_PushBack_Dummy_Core coll, dummy3
    Test_Collection_PushBack_Dummy_Core coll, dummy4
    Test_Collection_PushBack_Dummy_Core coll, dummy5
    Test_Collection_Count_Core coll
    Debug_Print_Collection_DummyItems coll
End Sub

Private Sub Test_Collection_PushBack_Dummy_Core(coll, item)
    Debug_Print "Collection.PushBack " & item.elem1
    coll.PushBack item
End Sub

Private Sub Debug_Print_Collection_DummyItems(coll)
    Dim i
    For i = 1 To coll.Count
        Debug_Print "Collection.Item(" & i & ").elem1: " & coll.Item(i).elem1
    Next
End Sub

'
' --- Test Core ---
'

Public Sub Test_Collection_Item_Core(coll, index_or_key)
    Debug_Print _
        "Collection.Item(" & index_or_key & "): " & coll.Item(index_or_key)
End Sub

Public Sub Test_Collection_Count_Core(coll)
    Debug_Print "Collection.Count: " & coll.Count
End Sub

Public Sub Test_Collection_PushBack_Core(coll, item)
    Debug_Print "Collection.PushBack " & item
    coll.PushBack item
End Sub

Public Sub Test_Collection_AddBefore_Core(coll, item, before)
    Debug_Print "Collection.AddBefore " & item & ", " & before
    coll.AddBefore item, before
End Sub

Public Sub Test_Collection_AddAfter_Core(coll, item, after)
    Debug_Print "Collection.AddAfter " & item & ", " & after
    coll.AddAfter item, after
End Sub

Public Sub Test_Collection_AddWithKey_Core(coll, item, key)
    Debug_Print "Collection.AddWithKey " & item & ", " & key
    coll.AddWithKey item, key
End Sub

Public Sub Test_Collection_AddWithKeyBefore_Core(coll, item, key, before)
    Debug_Print _
        "Collection.AddWithKeyBefore " & item & ", " & key & ", " & before
    coll.AddWithKeyBefore item, key, before
End Sub

Public Sub Test_Collection_AddWithKeyAfter_Core(coll, item, key, after)
    Debug_Print _
        "Collection.AddWithKeyAfter " & item & ", " & key & ", " & after
    coll.AddWithKeyAfter item, key, after
End Sub

Public Sub Test_Collection_Add_Core(coll, item, key, before, after)
    Debug_Print _
        "Collection.Add " & item & ", " & key & ", " & before & ", " & after
    coll.Add item, key, before, after
End Sub

Public Sub Test_Collection_Remove_Core(coll, index_or_key)
    Debug_Print "Collection.Remove " & index_or_key
    coll.Remove index_or_key
End Sub

Public Sub Test_Collection_Clear_Core(coll)
    Debug_Print "Collection.Clear"
    coll.Clear
End Sub

Public Sub Test_Collection_Contains_Core(coll, key)
    Debug_Print "Collection.Contains(" & key & "): " & coll.Contains(key)
End Sub

Public Sub Debug_Print_Collection_Items(coll)
    Dim i
    For i = 1 To coll.Count
        Debug_Print "Collection.Item(" & i & "): " & coll.Item(i)
    Next
End Sub

Public Sub Debug_Print_Separator
    Debug_Print "-----"
End Sub
