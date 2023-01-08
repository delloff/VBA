'values in collections are not mutable
'collections are similar to arrays and we can choose between them depending on situation
'see dicts, they are similar
'see On Error GoTo for unique collections


Sub Urok22_1()

    'declaring and creating a collection
    Dim someColl As New Collection
    
    'collections have 4 methods: add, remove, count, item
    
    'add
    someColl.Add "Test" 'will be index 1
    someColl.Add 10     'will be index 2 etc
    someColl.Add 22.22
    someColl.Add "ABc1"
    someColl.Add ""
    
    'item
    Debug.Print someColl.Item(3)
    
    'outputting a value from collection by its index, similar to item method
    Debug.Print someColl(3)
    
    'count (similar to len() in python)
    Debug.Print someColl.Count
    
    'remove
    someColl.Remove (5)
    Debug.Print someColl.Count
    
    
End Sub



Sub Urok22_2()

    'declaring and creating a collection
    Dim someColl As New Collection
    
    someColl.Add "Test"
    someColl.Add 10
    someColl.Add 22.22
    someColl.Add ""
    someColl.Add "ABc1"
    
    'using for loop with count
    For Counter = 1 To someColl.Count
        Debug.Print someColl(Counter)
    Next Counter
End Sub

'add method

Sub Urok22_3()

    'declaring and creating a collection
    Dim someColl As New Collection
    
    'add value, key(should be unique & string. If no key is given it uses indexes as keys by default)
    someColl.Add 1, "first index"
    
    'placing value in collection (before/after)
    someColl.Add 2, "second index", 1
    someColl.Add 3, "third index", , 1
    
    Debug.Print someColl("second index")    'get value by its key
    Debug.Print someColl.Item(2)            'get value by its index
    
    Debug.Print "------"
    For Counter = 1 To someColl.Count
        Debug.Print someColl(Counter)
    Next Counter
   
End Sub

