Sub MoveValuesToColumnK()
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = Cells(Rows.Count, "J").End(xlUp).Row ' Find the last row in column J
    
    For i = 2 To lastRow ' Start from row 2 assuming row 1 is the header
        If InStr(1, Cells(i, "J"), "Заявка", vbTextCompare) > 0 Then
            Cells(i, "K") = Cells(i, "J") ' Copy the value to column K in the same row
            Cells(i, "J") = "" ' Clear the value in column J
        End If
    Next i
End Sub
