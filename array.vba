Option Explicit

Sub learningArrays()

Dim testArray(1 To 5) As Long
'testArray(4) = 40
MsgBox LBound(testArray)    'Lower Bound
MsgBox UBound(testArray)    'Upper Bound


End Sub

Sub testArray_I()

Dim testArray(1 To 5) As Long
Dim i As Long

'записываем в массив значение ячеек
For i = LBound(testArray) To UBound(testArray)
    testArray(i) = ActiveSheet.Range("C" & i + 11)
Next i

'выводим в диологовом окне эти значения
For i = LBound(testArray) To UBound(testArray)
    MsgBox testArray(i)
Next i

End Sub


'динамичный массиф (определяет границы, которые нам нужны)

Sub dynamicArrays()

Dim dynArray()
Dim lboundVar As Long
Dim uboundVar As Long

lboundVar = 1
uboundVar = 3

ReDim dynArray(lboundVar To uboundVar)

Dim i As Integer
For i = LBound(dynArray) To UBound(dynArray)
    MsgBox i
Next i


'изменияем границы массива

lboundVar = 5
uboundVar = 7

ReDim dynArray(lboundVar To uboundVar)
MsgBox "Произошло изменение границ массива"
For i = UBound(dynArray) To LBound(dynArray) Step -1
    MsgBox i
Next i

End Sub
