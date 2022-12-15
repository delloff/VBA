
Sub Urok9()
    If Range("B3") = 3 Then
        MsgBox "three"
        Range("B3").Font.Bold = True
        Range("B3").Font.Color = vbRed
    End If
    
End Sub


Sub Urok9_1()
    Dim cellname As Range
    
    For Each cellname In Range("B1:B6")
        If cellname = 5 Then
            MsgBox cellname.Offset(0, -1) & ": Five"
        
        ElseIf cellname = 4 Then
            MsgBox cellname.Offset(0, -1) & ": Four"
        
        ElseIf cellname = 3 Then
            MsgBox cellname.Offset(0, -1) & ": Three"
        
        Else:
            MsgBox "?"
        
        End If
    Next cellname

End Sub


Sub Urok9_2()
    Dim cellname As Range
       
    For Each cellname In Range("A1:A5")
        If cellname = "Ivan" And cellname.Font.Bold = True Then
            MsgBox "OK!"
        Else: MsgBox "not bold & not Ivan"
        End If
           
    Next cellname

End Sub


Sub Urok9_3()
    Dim cellname As Range
       
    For Each cellname In Range("B1:B5")
        If cellname = 5 Then
            cellname.Offset(0, 1) = "excellent"
        ElseIf cellname = 4 Then
            cellname.Offset(0, 1) = "good"
        Else:
            cellname.Offset(0, 1) = "satisfactory"
                      
        End If
    Next cellname

End Sub


