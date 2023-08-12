
'import data

Sub importDB()
    
    Dim filePath As Variant
    Dim wbImportFile As Workbook
    Dim wb As Workbook
    Set wb = Workbooks("access.xlsm")   'set filename for PT
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    

'STAMI

    'saving filepath to variable
    filePath = "D:\VBA\BI\Bestell" & ChrW(252) & "bersichtSTAMI.xlsx"

    wb.Worksheets("InforSTAMI").Cells.ClearContents
      
    If filePath <> False Then
        'creating variable for source workbook to work with it, open it, save its data
        On Error Resume Next
        Set wbImportFile = Workbooks.Open(filePath, ReadOnly:=True)
        On Error GoTo 0
        
        'copy data from source worksheet
        wbImportFile.Worksheets(1).UsedRange.Copy
        
        'paste this data to different workbook
        wb.Worksheets("InforSTAMI").Range("A1").PasteSpecial (xlPasteAll)
        wbImportFile.Close (False)
        
    End If
  

'STASA

    filePath = "D:\VBA\BI\Bestell" & ChrW(252) & "bersichtSTASA.xlsx"
    wb.Worksheets("InforSTASA").Cells.ClearContents
    
    If filePath <> False Then
        'creating variable for source workbook to work with it, save its data, open it
        On Error Resume Next
        Set wbImportFile = Workbooks.Open(filePath, ReadOnly:=True)
        On Error GoTo 0
        
        'copy data from source worksheet
        wbImportFile.Worksheets(1).UsedRange.Copy
        
        'paste this data to different workbook
        wb.Worksheets("InforSTASA").Range("A1").PasteSpecial (xlPasteAll)
        wbImportFile.Close (False)
    End If
    

' Clear the contents of the clipboard
    Dim clipboard As Object
    Set clipboard = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'Create DataObject object
    
    clipboard.SetText "" ' Set empty text to the clipboard
    clipboard.PutInClipboard ' Put the empty text in the clipboard
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
