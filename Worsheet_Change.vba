
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'
'    Cells.Interior.ColorIndex = xlNone
'
'    Selection.EntireRow.Interior.ColorIndex = 40
'
'
'End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
Const WS_RANGE As String = "A1:C10"


On Error GoTo ws_exit:
Application.EnableEvents = False
If Not Intersect(Target, Me.Range(WS_RANGE)) Is Nothing Then
With Target
.Interior.ColorIndex = 38
End With
End If


ws_exit:
Application.EnableEvents = True
End Sub


'https://www.thesmallman.com/vba-worksheet-change-events#:~:text=In%20Excel%20a%20Worksheet%20Change,Red%20whenever%20the%20cell%20changes.

https://you.com/search?q=write+vba+code+to+automatically+update+one+excel+file+with+another+one&fromSearchBar=true&tbm=youchat&cid=c0_8aad94e2-b738-4234-aa89-aac7a416fc24
