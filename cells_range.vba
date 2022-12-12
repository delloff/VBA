Option Explicit

'Workbooks.Worksheets.Range("A1")


'WORKBOOKS:

    'Wokbooks("lessons")
    'Workbooks(1) - ïîðÿäîê â êîòîðîì ìû îòêðûëè êíèãè (çàïîìèíàåò ïî èíäåêñàì)
    'ThisWorkbook
    'ActiveWorkbook - ññûëàåòñÿ íà ýêñåëü ôàéë íà ÿ÷åéêè êîòîðîãî ìû ïîñëåäíèé ðàç êëèêíóëè


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

'åñëè íå óêàçûâàòü êîíêðåòíóþ êíèãó, òîãäà ïî óìîë÷àíèþ áåðåòñÿ ActiveWorkbook:
'    MsgBox Worksheets(2).Name

End Sub


'RANAGE:
'    Range ("A1")


Sub Urok5_2()

'    MsgBox Range("A1").Address
'    MsgBox Range("A1:D4").Address

End Sub




