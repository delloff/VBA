Private Sub Worksheet_Change(ByVal Target As Range)
    Dim currentUser As String
    Dim authorizedUsers As String
    Dim cell As Range

    currentUser = Environ("USERNAME")
    authorizedUsers = "K271,P14,J19"

    For Each cell In Target.Cells
        If Not Intersect(cell, Me.Range("C3:C29500")) Is Nothing Then
            'allow changes to column C (commment field) for all users
            'do nothing
        ElseIf Not Intersect(cell, Me.Range("A1:AA1048576")) Is Nothing Then 'restrict changes to columns other than C to unauthorized users
            If InStr(1, "," & authorizedUsers & ",", "," & currentUser & ",", vbTextCompare) = 0 Then '0 means user not found
                Application.EnableEvents = False

                If Target.CountLarge >= 1 Then 'show error message
                    MsgBox "You are not authorized to change this cell."
                End If

                Application.Undo
                Application.EnableEvents = True
                Exit For
            End If
        End If
    Next cell
End Sub
