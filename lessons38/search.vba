

Sub Urok26_1()

Dim rangeToSearch As Range
Set rangeToSearch = Range("A1:C3")

'the cell itself is saved in rangeToSearch, not its value. That is why we can use offset
Debug.Print rangeToSearch.Find("H")
Debug.Print rangeToSearch.Find("H").Address

End Sub



Sub Urok26_2()

Dim rangeToSearch As Range
Set rangeToSearch = Range("A1:C3")
Dim cellValue As Range
Set cellValue = rangeToSearch.Find("J")

    If cellValue Is Nothing Then
        MsgBox "value not found", Buttons:=vbInformation
    Else
        Debug.Print cellValue.Address
    End If

End Sub

'Multiple search

Sub Urok26_3()

Dim rangeToSearch As Range
Dim firstCellFound As Range
Dim cellName As Range

Set rangeToSearch = Range("A1:C3")
Set firstCellFound = rangeToSearch.Find("A", , , , xlByRows)    'xlByRows skipping 1st row?
Set cellName = firstCellFound

'MsgBox rangeToSearch.FindNext(firstCellFound).Address

If firstCellFound Is Nothing Then
    Exit Sub
End If


Dim allCellColl As New Collection

Do
    allCellColl.Add cellName.Address
    Set cellName = rangeToSearch.FindNext(cellName)
    
Loop While firstCellFound.Address <> cellName.Address

'output all found addresses
Dim Counter As Long
For Counter = 1 To allCellColl.Count
    MsgBox allCellColl(Counter)
Next Counter


End Sub

