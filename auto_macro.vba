Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

Dim testRange As Range
Dim myRange As Range

Set testRange = Range("G9:I11")
Set myRange = Selection

If Intersect(testRange, myRange) Is Nothing Then
    'Selection is NOT inside the range.

    MsgBox "Selection is outside the test range."

Else
    'Selection IS inside the range.

    MsgBox "Selection is inside the test range."
    MsgBox Target.Address
End If

End Sub

'------------------------------------------------------------------

Option Explicit


Private Sub Workbook_Open()

Sheets(1).Select
MsgBox "Hello Andrei"

End Sub



Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)

If Sh.Name = "colors" And Target.Interior.Color <> RGB(0, 255, 0) Then
    Target.Interior.Color = RGB(0, 255, 0)
Else:
    Target.Interior.Color = xlNone

End If
    

End Sub
