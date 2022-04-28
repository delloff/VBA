
Sub learnDoLoop_I()

Dim checker As String
checker = "ok"

Do While checker = "ok"
    checker = InputBox("enter 'ok' to continue")
Loop
      
End Sub



Sub learnDoLoop_II()

Dim checker As String

Do
    checker = InputBox("enter 'ok' to continue")
    Loop While checker = "ok"
      
End Sub




Sub learnDoLoop_Until_III()

Dim checker As String

Do Until checker = "ok"
    checker = InputBox("enter 'ok' to quit")
Loop
      
End Sub



Sub learnDoLoop_Until_IV()

Dim checker As String

Do
    checker = InputBox("enter 'ok' to quit")
Loop Until checker = "ok"
      
End Sub



Sub learnExitDo()

Dim i As Long
i = 1

Do Until i >= 300
    ThisWorkbook.Worksheets(1).Range("G" & i) = i
    i = i + 1
    
    If ThisWorkbook.Worksheets(1).Range("G" & i) = "STOP" Then
    Exit Do
        
        
    End If
Loop

End Sub




Sub homeWork_II()

Dim checker As String


i = 1

Do
    checker = InputBox("enter 'ok' to quit" & vbCr & "loop number: " & i)
    i = i + 1
Loop Until checker = "ok"

End Sub

