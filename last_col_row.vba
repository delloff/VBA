

Sub Urok11()

    MsgBox ThisWorkbook.Worksheets(1).Range("A1:C5").Cells.Count
    MsgBox ThisWorkbook.Worksheets.Count
    MsgBox Application.Workbooks.Count

End Sub


Sub Urok11_2()
    
    Range("A26").End(xlUp).Select
    
End Sub


'find last used row

Sub Urok11_3()

    Dim lastrow As Long
    
    'count last row in col A, go up to last used row and output int index
    lastrow = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox lastrow

End Sub


'find last used col

Sub Urok11_4()

    Dim lastcol As Long
    
    'finding last used col in second row, return its index
    lastcol = Worksheets(1).Cells(2, Columns.Count).End(xlToLeft).Column
    MsgBox lastcol

End Sub


Sub s_to_num()

With Range("H9")
.NumberFormat = "General"
.Value = .Value
End With

End Sub
