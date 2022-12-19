
'CALL (call certain function from this or different module)

Sub Urok13()

'    Call Urok13_1
'    Call arrays.Urok12_3

End Sub


Sub Urok13_1()

    MsgBox "hi"

End Sub

Sub Urok13_3()

    Dim rangeToCount As Range
    Dim notEmptyCells As Long
    Dim cellChecked As Range
    
    Set rangeToCount = Range("B2:E5")
    For Each cellChecked In rangeToCount
        If cellChecked.Value <> "" Then
            notEmptyCells = notEmptyCells + 1
        End If
        
    Next cellChecked
    
    MsgBox notEmptyCells
        
End Sub

