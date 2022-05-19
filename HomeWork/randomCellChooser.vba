Sub randomCellChooser()

Dim randomCell As Single, Cell As Range, countCells As Long, randomVal As Long, i As Long, x As Long, y As Long, countRows As Long, countColums As Long

countCells = Range("E8:Y32").Rows.Count * Range("E8:Y32").Columns.Count '525
countRows = Range("E8:Y32").Rows.Count
countColums = Range("E8:Y32").Columns.Count

Randomize
randomVal = Int((countCells * Rnd) + 1)




For Each Cell In Range("E8:Y32")
    randomCell = Rnd
    If randomCell > 0 And randomCell < 0.01 And Application.WorksheetFunction.CountIf(Range("E8:Y32"), "x") < 10 Then Cell.Value = "x"


Next

End Sub
