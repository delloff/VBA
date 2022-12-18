
Sub Urok12_1()

Dim arr(0 To 5) As String
'arr(4) = 40    'index 4 holds 40
    
'    bordera of the array:
'    MsgBox LBound(arr)
'    MsgBox UBound(arr)

Dim i As Long

'putting values in array
For i = LBound(arr) To UBound(arr)
    arr(i) = ActiveSheet.Range("C" & i + 1)
Next i

'outputting values in msgbox
For i = LBound(arr) To UBound(arr)
    MsgBox arr(i)
Next i

End Sub

'DYNAMIC ARRAY

Sub urok12_2()

Dim dynarr()    'we dont declare borders
Dim lboundVar As Long
Dim uboundVar As Long

lboundVar = 1
uboundVar = 10
ReDim dynarr(lboundVar To uboundVar)    'declare borders to our dynArr()
Dim i As Long

For i = LBound(dynarr) To UBound(dynarr)
    MsgBox i
Next i

lboundVar = 5
uboundVar = 8

ReDim dynarr(lboundVar To uboundVar)
For i = LBound(dynarr) To UBound(dynarr)
    MsgBox i
Next i

End Sub


'HOMEWORK

Sub Urok12_3()

Dim dynarr()
Dim lboundVar As Long
Dim uboundVar As Long

first_ind = 1
last_ind = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
ReDim dynarr(first_ind To last_ind)
Dim i As Long

'putting values in array
For i = LBound(dynarr) To UBound(dynarr)
    dynarr(i) = ActiveSheet.Range("A" & i)
Next i


'outputting values in msgbox
For i = LBound(dynarr) To UBound(dynarr)
    MsgBox dynarr(i) & " - grade: " & ActiveSheet.Range("B" & i)
Next i

End Sub


