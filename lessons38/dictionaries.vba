'sheet "find"
'object dict belongs not to excel but to pc file system
'that is why we need to add library: microsoft scripting runtime (dict is one of its elements)
'to do this go to tools->references
'we can now use intelliSence with this library
'odjects can belong to third party libraries. We need to add them this way (Outlook etc)
'dict has the following methods: add,remove,removeAll,count,exists,keys,items,compareMode
'dict similar to collections. but in collections we cannot change added items. But we can do this in dict
'keys and items are stored in two separate arrays

Sub Urok33_1()
    
    'declare dict var
    Dim myDict As New Scripting.Dictionary
    Dim wsSheet As Worksheet
    
    Set wsSheet = ThisWorkbook.Worksheets("find")
    
    
    'COMPAREMODE upper-lower case sensitive. Should be used before adding keys
    myDict.CompareMode = TextCompare    'case insensitive: pens = Pens will reassign value
    myDict.CompareMode = BinaryCompare  'case sensitive (by dafault): pens <> Pens will add both
    
    
    'ADD
    'fill the dict. Keys should be unique.  Values can be objects,arrays,collections,dictionarues
    myDict.Add Key:="pens", Item:=800
    myDict.Add "pencils", Item:=wsSheet.Range("C39")
    myDict.Add "pins", 200
    myDict.Add "markers", 200
    myDict("books") = 10
    
    Debug.Print myDict("pens")
    Debug.Print myDict("pencils").Address(False, False)
    Debug.Print myDict("books")
    
    'this will not work because we didnt assign range to it:
    'Debug.Print myDict("pins").Value

    'reassign items
    myDict("pins") = 300
    Debug.Print myDict("pins")
    
    'EXISTS     We dont have it in collections and we use On Error Go To instead there
    If myDict.Exists("books") = True Then
        Debug.Print "sold " & myDict("books") & "pcs"
    Else
        Debug.Print "Key [books] is not in dict"
    End If
    
    
    'COUNT
    Debug.Print "articles in stock: " & myDict.Count
    
    'REMOVE
    'myDict.Remove ("books")
    
    'REMOVE ALL
    'myDict.RemoveAll
    
    'KEYS or ITEMS get keys or items from array
    Debug.Print myDict.Keys(1)
    Debug.Print myDict.Items(4)
    
End Sub

'using for loop on dict

Sub Urok33_2()
    
    'declare dict var
    Dim myDict As New Scripting.Dictionary
    Dim intCounter As Integer
    
    
    myDict.Add Key:="pens", Item:=800
    myDict.Add "pencils", 1000
    myDict.Add "pins", 200
    myDict.Add "markers", 200
    
    For intCounter = 0 To myDict.Count - 1
        Debug.Print myDict.Keys(intCounter), myDict.Items(intCounter)
    Next intCounter
    
End Sub

