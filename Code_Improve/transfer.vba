Sub transferZVZData()
    
    Dim zvzSheet As Worksheet
    Dim lastZvzRow, sum As Long
    
    Set zvzSheet = Workbooks("_ZVZ_PivotTable.xlsm").Worksheets("ZVZ")
    Set avzSheet = Workbooks("Elektromotor  komplett BD9.xlsx").Worksheets("AVZ")
    
    lastAvzRow = avzSheet.Cells(Rows.Count, "B").End(xlUp).Row
    lastMainRow = ThisWorkbook.Worksheets("GoTo").Cells(Rows.Count, "K").End(xlUp).Row
    
    'set range for zvz nums
    Set avzNums = ThisWorkbook.Worksheets("GoTo").Range("K9:K" & lastMainRow)
    Set avzRange = avzSheet.Range("B4:B" & lastAvzRow)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    'paste data in ZVZ
    For Each avzNum In avzNums
        Set targetCell = avzRange.find(avzNum)
        lastZvzRow = zvzSheet.Cells(Rows.Count, "I").End(xlUp).Row
        
        With targetCell
            .Offset(0, 5).Copy  'article
            zvzSheet.Cells(lastZvzRow + 1, "I").PasteSpecial Paste:=xlPasteValues
            .Offset(0, 8).Copy  'bennennung1
            zvzSheet.Cells(lastZvzRow + 1, "J").PasteSpecial Paste:=xlPasteValues
            .Offset(0, 16).Copy 'drawing
            zvzSheet.Cells(lastZvzRow + 1, "H").PasteSpecial Paste:=xlPasteValues
            .Offset(0, 28).Copy 'herrsteler
            zvzSheet.Cells(lastZvzRow + 1, "X").PasteSpecial Paste:=xlPasteValues
        End With
            
        'finding main article
        i = -1
        Set numLow = targetCell.Offset(0, 1)
        Set numUpper = targetCell.Offset(i, 1)
        
        Do Until numUpper < numLow
            Set numUpper = targetCell.Offset(i, 1)
            i = i - 1
        Loop
        numUpper.Offset(0, 4).Copy
        zvzSheet.Cells(lastZvzRow + 1, "G").PasteSpecial Paste:=xlPasteValues
        
        'calculating quantity
        If IsEmpty(numUpper.Offset(0, 3)) Then  'the cell for main article under 0 is always empty, we set it to 1 (when we calculate 1's)
            numUpper.Offset(0, 3) = 1
        End If
        sum = targetCell.Offset(0, 4) * numUpper.Offset(0, 3)
        zvzSheet.Cells(lastZvzRow + 1, "L") = sum
        zvzSheet.Cells(lastZvzRow + 1, "M").Value = "pc"
        
        'finding Baugruppe
        If numLow <> 1 Then
            Do Until numUpper = 1
                Set numUpper = targetCell.Offset(i, 1)
                i = i - 1
            Loop
        
            If numUpper.Offset(0, 8) = " " Or numUpper.Offset(0, 8) = "" Or IsEmpty(numLow.Offset(0, 8)) Then
                avzSheet.Range("K4").Copy
                zvzSheet.Cells(lastZvzRow + 1, "C").PasteSpecial Paste:=xlPasteValues
            Else
                numUpper.Offset(0, 8).Copy
                zvzSheet.Cells(lastZvzRow + 1, "C").PasteSpecial Paste:=xlPasteValues
            End If
        End If
            
        If numLow = 1 Then
            If numLow.Offset(0, 8) = " " Or numLow.Offset(0, 8) = "" Or IsEmpty(numLow.Offset(0, 8)) Then
                avzSheet.Range("K4").Copy
                zvzSheet.Cells(lastZvzRow + 1, "C").PasteSpecial Paste:=xlPasteValues
            ElseIf Not IsEmpty(numLow.Offset(0, 8)) And numLow.Offset(0, 8) <> avzSheet.Range("K4") Then
                avzSheet.Range("K4").Copy
                zvzSheet.Cells(lastZvzRow + 1, "C").PasteSpecial Paste:=xlPasteValues
            Else
                numLow.Offset(0, 8).Copy
                zvzSheet.Cells(lastZvzRow + 1, "C").PasteSpecial Paste:=xlPasteValues
            End If
        End If
        
    Next avzNum
    
'    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

End Sub
