'https://analystcave.com

'FOR        'for counter 1 to 10

Sub Urok_8()

    Dim i As Long
    
    For i = 1 To 10
'        ThisWorkbook.Worksheets("Sheet1").cells(i, 1) = i
        ThisWorkbook.Worksheets("Sheet1").Range("A" & i) = i
    Next i
    
End Sub



'FOR EACH       'For Each Iterator in Items

Sub Urok8_1()

    Dim cellName As Range
    Dim i As Long
    
    i = 1
    
    For Each cellName In Range("A1:C3")
        cellName.Value = i
        i = i + 1
    Next cellName

End Sub


Sub Urok8_2()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        MsgBox ws.name
    Next ws
End Sub

'looping backwards

Sub test()

Dim i As Long
 
    For i = 3 To 1 Step -1
       Debug.Print i
    Next i

End Sub




