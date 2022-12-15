Option Explicit

'Workbooks.Worksheets.Range("A1")


'WORKBOOKS:

'WORKBOOKS:

    'Wokbooks("les_main.xlsm")
    'Workbooks(1) - the order in which we opened the books (remembers by index)
    'ThisWorkbook
    'ActiveWorkbook - refers to the excel file which cells we last clicked


Sub Urok5()

'    MsgBox Workbooks("lessons_secondary").Name
'    MsgBox Workbooks(2).Name
'    MsgBox ThisWorkbook.Name
'    MsgBox ActiveWorkbook.Name

End Sub


'WORKSHEETS:

    'Worksheets("Sheet1")
    'Worksheets(1)
    'ActiveSheet


Sub Urok5_1()

'   MsgBox Workbooks(1).Worksheets("Sheet2").Name
'   MsgBox Workbooks(1).Worksheets(2).Name
'   MsgBox Workbooks(1).ActiveSheet.Name

'If you do not specify a specific workbook, then the default is ActiveWorkbook:
'    MsgBox Worksheets(2).Name

End Sub


'RANAGE:
'    Range ("A1")


Sub Urok5_2()

'    MsgBox Range("A1").Address
'    MsgBox Range("A1:D4").Address

End Sub

'CELLS


'Workbooks.Worksheets.Cells(1,1)

Sub Urok6()
    Dim rowIndex As Long, colIndex As Long, les_main_Sheet1 As Worksheet
    
    Set les_main_Sheet1 = Workbooks("les_main.xlsm").Worksheets("Sheet1")
    

    With les_main_Sheet1
        rowIndex = .cells(2, 2)
        colIndex = .cells(3, 2)
        
        les_main_Sheet1.cells(4, 3) = "test2"
    End With


End Sub

'MIXED METHOD (RANGE+CELLS)
Sub Urok6_1()

'    ThisWorkbook.Worksheets("Sheet1").Range("A1:E10") = 0
'    ThisWorkbook.Worksheets("Sheet1").Range(cells(1, 1), cells(10, 5)) = 0
    Workbooks("les_main.xlsm").Worksheets("Sheet1").Range(cells(1, 1), cells(10, 5)) = 0

End Sub

Sub Urok6_2()

'    ThisWorkbook.Worksheets("Sheet1").Range("A1:E10") = 1
'    ThisWorkbook.Worksheets("Sheet1").Range(cells(1, 1), cells(10, 5)) = 1
    Workbooks("les_main.xlsm").Worksheets("Sheet1").Range(cells(1, 1), cells(10, 5)) = 1

End Sub

'WITH - shortens repetitive code

Sub Urok7()

'    ThisWorkbook.Worksheets("Sheet1").Range("A1") = 1
'    ThisWorkbook.Worksheets("Sheet1").Range("A2") = 2
'    ThisWorkbook.Worksheets("Sheet1").Range("A3") = 3
    
    'we can cut down the above code to:
    With ThisWorkbook.Worksheets("Sheet1")
        Range("A1") = 1
        Range("A2") = 2
        Range("A3") = 3
    End With
    
End Sub

Sub Urok7_1()

    With ThisWorkbook.Worksheets("Sheet1").Range("A1").Offset(4, 4)
        .Font.Bold = True
        .Font.Color = vbRed
        .Value = 100
    End With
    
End Sub


'OFFSET (works also with negatives)

Sub Urok7_2()

    ThisWorkbook.Worksheets("Sheet1").Range("A1").Offset(1, 2) = 4  'C2

End Sub


