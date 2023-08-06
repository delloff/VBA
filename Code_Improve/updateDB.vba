
'import data

Sub Urok28_1()
    
    Dim filePath As Variant
    Dim wbImportFile As Workbook
    
    'saving filepath to variable
    filePath = Application.GetOpenFilename '(Title:="Specify Data File", filefilter:="Excel Files(*.xls*),*xls*")
    
    'to avoid visual opening of excel file
    Application.ScreenUpdating = False
    
    If filePath <> False Then
        'creating variable for source workbook to work with it, save its data, open it
        Set wbImportFile = Workbooks.Open(filePath)
        
        'copy data from source worksheet
        wbImportFile.Worksheets("Sheet1").Range("C8:F29").Copy
        
        'paste this data to different workbook
        ThisWorkbook.Worksheets("List").Range("F2").PasteSpecial (xlPasteAll)
        wbImportFile.Close (False)
    Else
        MsgBox "please choose a file"
    End If
    
    Application.ScreenUpdating = True

End Sub


Sub clearData1()

    Range("F2:I25").Clear

End Sub
