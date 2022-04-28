Sub notes()
    Dim note As Range
    For Each note In Range("C12:C16")
        If note = 3 Then
            note.Offset(0, 1) = "óäîâëåòâîðèòåëüíî"
        ElseIf note = 4 Then
            note.Offset(0, 1) = "õîðîøî"
        ElseIf note = 5 Then
            note.Offset(0, 1) = "îòëè÷íî"
        End If
    Next note

End Sub
