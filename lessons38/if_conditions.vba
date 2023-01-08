Sub Urok9()
    If Range("B3") = 3 Then
        MsgBox "three"
        Range("B3").Font.Bold = True
        Range("B3").Font.Color = vbRed
    End If
    
End Sub


Sub Urok9_1()
    Dim cellName As Range
    
    For Each cellName In Range("B1:B6")
        If cellName = 5 Then
            MsgBox cellName.Offset(0, -1) & ": Five"
        
        ElseIf cellName = 4 Then
            MsgBox cellName.Offset(0, -1) & ": Four"
        
        ElseIf cellName = 3 Then
            MsgBox cellName.Offset(0, -1) & ": Three"
        
        Else:
            MsgBox "?"
        
        End If
    Next cellName

End Sub


Sub Urok9_2()
    Dim cellName As Range
       
    For Each cellName In Range("A1:A5")
        If cellName = "Ivan" And cellName.Font.Bold = True Then
            MsgBox "OK!"
        Else: MsgBox "not bold & not Ivan"
        End If
           
    Next cellName

End Sub


Sub Urok9_3()
    Dim cellName As Range
       
    For Each cellName In Range("B1:B5")
        If cellName = 5 Then
            cellName.Offset(0, 1) = "excellent"
        ElseIf cellName = 4 Then
            cellName.Offset(0, 1) = "good"
        Else:
            cellName.Offset(0, 1) = "satisfactory"
                      
        End If
    Next cellName

End Sub
