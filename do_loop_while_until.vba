
'DO     this cycle is used with methods while/until


' DO WHILE ... LOOP

Sub Urok10_1()

Dim checker As String
    checker = "OK"

    Do While checker = "OK"
        checker = InputBox("Enter 'OK' to repeat the cycle")

    Loop

End Sub

' DO ... LOOP WHILE ...

Sub Urok10_2()

Dim checker As String

    Do
        checker = InputBox("Enter 'OK' to stop the cycle")
    
    Loop While checker <> "OK"

End Sub

'DO LOOP UNTIL (opposite to while)

Sub Urok10_3()

Dim checker As String

    Do Until checker = "OK"
    
        checker = InputBox("Do not Enter 'OK' to stop the cycle")
    
    Loop

End Sub

'DO ... LOOP UNTILL ...
Sub Urok10_4()

Dim checker As String

    Do
        checker = InputBox("Do not Enter 'OK' to stop the cycle")
    
    Loop Until checker = "OK"

End Sub


'EXIT DO

Sub Urok10_5()

Dim i As Long
i = 1

    Do Until i >= 35
        ThisWorkbook.Worksheets("Urok9").Range("H" & i) = i
        i = i + 1
        
        If ThisWorkbook.Worksheets("Urok9").Range("H" & i) = "Stop" Then
            Exit Do
        End If
    Loop

End Sub

'HOMEWORK

Sub Urok10_6()

Dim checker As String
Dim i As Long

i = 1
    Do
        checker = InputBox("Do not Enter 'OK' to stop the cycle. Cycle: " & i)
        i = i + 1
    Loop Until checker = "OK"

End Sub
