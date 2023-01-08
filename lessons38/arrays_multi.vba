
'MULTI DIMENSIONAL ARRAYS

Sub Urok27_1()

'    'simple variable
'    Dim normalVar As String
'
'    normalVar = Worksheets("MultiArr1").Cells(1, 1)
'    'MsgBox normalVar
'
'    'one dimensional array
'    Dim oneDimArr(1 To 3) As String
'
'    oneDimArr(1) = Worksheets("MultiArr1").Cells(3, 1)
'    oneDimArr(2) = Worksheets("MultiArr1").Cells(3, 2)
'    oneDimArr(3) = Worksheets("MultiArr1").Cells(3, 3)
'    'Debug.Print oneDimArr(1) & oneDimArr(2) & oneDimArr(3)
'
'
'   'two dimensional array - twoDimArr(rows, columns) like 2d table
    'if we use for loop we need to use rows and cols numbers from excel in brackets:
    Dim twoDimArr(5 To 7, 1 To 3) As String
'
'    twoDimArr(1, 1) = Worksheets("MultiArr1").Cells(5, 1)
'    twoDimArr(1, 2) = Worksheets("MultiArr1").Cells(5, 2)
'    twoDimArr(1, 3) = Worksheets("MultiArr1").Cells(5, 3)
'    twoDimArr(2, 1) = Worksheets("MultiArr1").Cells(6, 1)
'    twoDimArr(2, 2) = Worksheets("MultiArr1").Cells(6, 2)
'    twoDimArr(2, 3) = Worksheets("MultiArr1").Cells(6, 3)
'    twoDimArr(3, 1) = Worksheets("MultiArr1").Cells(7, 1)
'    twoDimArr(3, 2) = Worksheets("MultiArr1").Cells(7, 2)
'    twoDimArr(3, 3) = Worksheets("MultiArr1").Cells(7, 3)
'    Debug.Print twoDimArr(2, 2)
    
    'assigning values to array in the above example takes too long. It is better to use for loop
    
    Dim cellChecked As Range
    For Each cellChecked In Worksheets("MultiArr1").Range("A5:C7")
        twoDimArr(cellChecked.Row, cellChecked.Column) = cellChecked
    Next cellChecked
    Debug.Print twoDimArr(6, 2)

End Sub

        
'ALTERNATIVE 2D

Sub Urok27_2()

Dim twoDimArr(5 To 7, 1 To 3) As String
Dim i As Integer, j As Integer

    For i = 5 To 7
        For j = 1 To 3
            twoDimArr(i, j) = Cells(i, j).Value
        Next j
    Next i

Debug.Print twoDimArr(6, 2)

End Sub



'3d array
Sub Urok27_3()

    Dim threeDimArr(9 To 11, 5 To 7, 1 To 3) As String     'threeDimArr(Worsheets, Rows, Columns)
    'threeDimArr(1, 1, 1) = Worksheets("MultiArr").Cells(5, 1) ...
    Dim cellChecked As Range
    Dim wsCounter As Long
    
    For wsCounter = 9 To 11
        For Each cellChecked In Worksheets(wsCounter).Range("A5:C7")
            threeDimArr(wsCounter, cellChecked.Row, cellChecked.Column) = cellChecked
        Next cellChecked
    Next wsCounter
    
    Debug.Print threeDimArr(10, 6, 2)
End Sub


'Putting values in array and then to the worksheet

Sub Urok27_4()
    
    Dim threeDimArr(9 To 11, 5 To 7, 1 To 3) As String
    Dim rowCounter As Integer, colCounter As Integer
    Dim lastRow As Long
    
    'saving to array data from excel
    Dim cellChecked As Range
    For wsCounter = 9 To 11
        For Each cellChecked In Worksheets(wsCounter).Range("A5:C7")
            threeDimArr(wsCounter, cellChecked.Row, cellChecked.Column) = cellChecked
        Next cellChecked
    Next wsCounter
    
    'saving data from array to database (excel)
    For wsCounter = 9 To 11
        For rowCounter = 5 To 7
            For colCounter = 1 To 3
                lastRow = Worksheets("List").Range("A50").End(xlUp).Row     'A50 means the capacity of our database
                Worksheets("List").Range("A" & lastRow + 1) = wsCounter     'finding last row and put value on next row
                Worksheets("List").Range("B" & lastRow + 1) = rowCounter
                Worksheets("List").Range("C" & lastRow + 1) = colCounter
                Worksheets("List").Range("D" & lastRow + 1) = threeDimArr(wsCounter, rowCounter, colCounter)
            Next colCounter
        Next rowCounter
    Next wsCounter
    Worksheets("List").Range("D7") = threeDimArr(10, 6, 2)
    
    'trick
    Dim autoArr
    autoArr = Worksheets("List").Cells(1, 1).CurrentRegion
    Debug.Print autoArr(13, 4)
    
    
End Sub




