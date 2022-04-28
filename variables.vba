
Sub testSub()

    'BASIC VARIABLE
    Dim someText As String
    Dim someNumber As Long
    Dim someDate As Date

    someText = "Òåñòîâàÿ Èíôîðìàöèÿ"
    someNumber = 100
    someDate = "07.05.2022" 'dd.mm.yy

    MsgBox someDate



    'OBJECT VARIABLE
    Dim someWorkbook As Workbook
    Dim someworkSheet As Worksheet
    Dim someRange As Range


    Set someWorkbook = ThisWorkbook
    Set someworkSheet = Worksheets("List1")
    Set someRange = Range("C4")

'    MsgBox someWorkbook.Name
'    MsgBox someRange.Address

    Dim someNewDate As Date

    someNewDate = "29.04.2022"

    MsgBox someNewDate

End Sub
