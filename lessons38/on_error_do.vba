'What to do if there is an error
'On Error GoTo 0   (by defalt)
'On Error GoTo [mark]
'On Error Resume Next
'On Error GoTo -1


'On Error Go To 0

Sub Urok36_1()

    Debug.Print 1
    Debug.Print 2
    Debug.Print 3 / 0
    Debug.Print 4

End Sub


'On Error Resume Next (skipping error)
'see also Unique Collections

Sub Urok36_2()

    On Error Resume Next

    Debug.Print 1
    Debug.Print 2
    Debug.Print 3 / 0
    Debug.Print 4

End Sub


'On Error Go to [line]

Sub Urok36_3()

    On Error GoTo errorHandling

    Debug.Print 1
    Debug.Print 2
    Debug.Print 3 / 0
    Debug.Print 4
    
    Exit Sub    'if there are no errors (or if we comment the line with ZeroDivision etc)
    
errorHandling:
    MsgBox "Error occured, please check"

End Sub


'On Error GoTo -1

Sub Urok36_4()

    On Error GoTo errorHandling

    Debug.Print 1
    Debug.Print 2
    Debug.Print 3 / 0
    Debug.Print 4
    
    Exit Sub    'if there are no errors (or if we comment the line with ZeroDivision etc)
    
errorHandling:
    On Error GoTo -1
    On Error GoTo errorHandling2
    Debug.Print "Error occured, please check"
    Debug.Print 5
    Debug.Print 6 / 0
    Exit Sub
    
errorHandling2:
    Debug.Print "Error occured, please check"
    Debug.Print 7
    
End Sub

'sheet multiarr1
'get uniques (countries) using collection
'additional HW: create dropdown menu with unique countries using forms

Sub getUniques()

    Dim collUniques As New Collection
    Dim dataColumn As Range     'not whole column but range with countries
    Dim lastRow As Long
    Dim countryChecked As Range
    
    
    'getting last row (1 045 576) in column with countries in this part:
    'lastRow = ThisWorkbook.Worksheets("MultiArr1").Range("E" & ThisWorkbook.Worksheets("MultiArr1").Rows.Count)
    'then go up to last used row in this part and get the row number using .Row:
    '.end(xlUp).Row
    'so write it all together
    lastRow = _
    ThisWorkbook.Worksheets("MultiArr1").Range("E" & ThisWorkbook.Worksheets("MultiArr1").Rows.Count).End(xlUp).Row
    
    'now link
    Set dataColumn = ThisWorkbook.Worksheets("MultiArr1").Range("E2:E" & lastRow)
    
    'iterate through countries and skip countries that already added to collection
    On Error Resume Next
    For Each countryChecked In dataColumn
        'add value (current cell value), key. In this case one country is key & value
        'as keys should be str only, we will be safe if we use "cstr" (convert to str) with keys
        collUniques.Add countryChecked.Value, CStr(countryChecked.Value)
    Next countryChecked
    On Error GoTo 0
    
    
    Dim i As Long
    
    For i = 1 To collUniques.Count
        Debug.Print collUniques(i)
    Next i
    
End Sub


