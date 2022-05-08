Option Explicit


Sub HWdynamicArrays()

Dim nameArr()
Dim gradeArr()

Dim lastRow As Long
lastRow = Worksheets(1).Range("B" & Rows.Count).End(xlUp).Row

Dim lboundVar As Long
Dim uboundVar As Long

lboundVar = 1
uboundVar = lastRow - 1 'exclude first row

ReDim nameArr(lboundVar To uboundVar)
ReDim gradeArr(lboundVar To uboundVar)


Dim i As Long
For i = LBound(nameArr) To UBound(nameArr)
    nameArr(i) = Range("B" & i + 1)
    gradeArr(i) = Range("C" & i + 1)
Next i


For i = LBound(nameArr) To UBound(nameArr)
    MsgBox nameArr(i) & " - Оценка: " & gradeArr(i)
Next i


End Sub
