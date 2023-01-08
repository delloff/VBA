
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


'Select multiple files

Sub Urok29_1()

    Dim fileName As Variant
    Dim Counter As Long
    Dim workBookProcessed As Workbook
    Dim i As Long
    i = 0
    
    
    'if multiselect is on, then an array is automaticly created (even if only one file is selected)
    fileName = Application.GetOpenFilename(MultiSelect:=True)
    
    'check if user selected any files
    If IsArray(fileName) Then
        Application.ScreenUpdating = False
        
        For Counter = LBound(fileName) To UBound(fileName)
            'open each workbook and save data to our variable
            Set workBookProcessed = Workbooks.Open(fileName(Counter))
            
            'do smth with open file
            workBookProcessed.Worksheets("Sheet1").Range("A1:C1").Copy
            
            ThisWorkbook.Worksheets("List").Range("A35").Offset(i, 0).PasteSpecial (xlPasteValues)  'paste 123
            i = i + 1
          
            workBookProcessed.Close (True)  'close the file after we have done everything we need
                 
        Next Counter
        
        Application.ScreenUpdating = True
        
    End If
        
End Sub


Sub clearData2()

    Range("A35:C38").Clear

End Sub



Sub formatCells()

    With Worksheets("List").Range("A35:C37")
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    
End Sub


