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



Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range
    Dim cell As Range
    
    ' Set the range object to the column L data
    Set rng = Me.Range("L:L")
    
    ' Loop through each cell in the changed range
    For Each cell In Target
        ' Check if the changed cell is in column L
        If Not Intersect(cell, rng) Is Nothing Then
            Select Case cell.value
                Case "sony1", "sony2", "sony3"
                    cell.value = "sony"
                Case "philips1", "philips2", "philips3"
                    cell.value = "philips"
                Case "redmi1", "redmi2", "redmi3", "REDMI"
                    cell.value = "redmi"
            End Select
        End If
    Next cell
End Sub
