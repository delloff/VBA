Private Sub Worksheet_Change(ByVal Target As Range)
    Dim dataSheet As Worksheet
    Dim intersectRange As Range
    Dim cell As Range

    Set dataSheet = ThisWorkbook.Worksheets("data")
    Set intersectRange = Intersect(Target, Me.Range("H20:H" & Me.Cells(Me.Rows.Count, "H").End(xlUp).Row))

    If intersectRange Is Nothing Then Exit Sub

    For Each cell In intersectRange
        If cell.Row >= 20 Then
            If Not IsError(Application.Match(cell.Value, dataSheet.Columns("A"), 0)) Then
                cell.Offset(0, 1).Value = "kurana"
            ElseIf Not IsError(Application.Match(cell.Value, dataSheet.Columns("D"), 0)) Then
                cell.Offset(0, 1).Value = "komvad"
            Else
                cell.Offset(0, 1).Value = ""
            End If
        End If
    Next cell

End Sub
