
'Functions similar to procedures (Sub) but act like variables. See also Functions2
'Optional arguments with keyword 'Optional'. It can have a default value (=default value) and
'placed only after all required arguments

Function Revenue(price As Double, quantity As Double, Optional Discount As Double = 0) As Double

    Revenue = price * quantity * (1 - Discount)

End Function


Sub Revenue_sub()

Dim price As Double
Dim quantity As Double
Dim resultingRevenue As Double

price = Range("C5")
quantity = Range("D5")

Debug.Print Revenue(price, quantity, 0.05)

End Sub


Option Explicit

Function ABC_Analysis(rangeOfValues As Range, currentCell As Range, Optional ABC_Type As String)

'Îáúÿâëÿåì Array ïåðåìåííûå
Dim positionBeforeSortArray() As Long
Dim valuesArray() As Double
Dim percentageArray() As Double
Dim cumulativePercentageArray() As Double
Dim classArray() As String
'Îáúÿâëÿåì ïðî÷èå ðàáî÷èå ïåðåìåííûå
Dim valuesSum As Double
Dim cellChecked As Range
Dim Counter As Integer
Dim firstRowOfWorkedRangeInExcel As Long
Dim lastRowOfWorkedRangeInExcel As Long
Dim rowOfCurrentCell As Long

Call RegisterUDF                                                'Ðåãèñòðèðóåì ôóíêöèþ â áàçó äàííûõ ôóíêöèé Excel

'ABC-ÀÍÀËÈÇ
If Intersect(rangeOfValues, currentCell) Is Nothing Then            'Åñëè ÿ÷åéêà ñî çíà÷åíèåì, êîòîðîå ìû õîòèì êëàññèôèöèðîâàòü,
    ABC_Analysis = "Àíàëèçèðóåìîå çíà÷åíèå âíå îáëàñòè çíà÷åíèé!"   'íàõîäèòñÿ âíå îáëàñòè ñî âñåìè çíà÷åíèÿìè - òî çíà÷åíèå íå êëàññèôèöèðóåòñÿ"
                                                                    'Â ïðîòèâíîì æå ñëó÷àå íà÷èíàåì àíàëèç
Else:
    'Øàã1 - Èíèöèàëèçèðóåì âñå Arrays
    ReDim valuesArray(1 To rangeOfValues.Cells.Count)
    ReDim positionBeforeSortArray(1 To rangeOfValues.Cells.Count)
    ReDim classArray(1 To rangeOfValues.Cells.Count)
    ReDim percentageArray(1 To rangeOfValues.Cells.Count)
    ReDim cumulativePercentageArray(1 To rangeOfValues.Cells.Count)
    
    
    'Øàã 2 - Ñîõðàíÿåì äåíåæíûå çíà÷åíèÿ â ìàññèâ äëÿ çíà÷åíèé
    Counter = 1 'Íà÷èíàåì ñ èíèöèàëèçàöèè êàóíòåðà
    For Each cellChecked In rangeOfValues
        valuesArray(Counter) = cellChecked                          'Ñîõðàíÿåì ñïåðâà çíà÷åíèÿ
        positionBeforeSortArray(Counter) = Counter                  'È çàòåì èçíà÷àëüíóþ ïîçèöèþ êàæäîãî çíà÷åíèÿ â ñïèñêå
        Counter = Counter + 1                                       'Èíêðåìåíòèðóåì (ïîâûøàåì íà åäèíèöó) êàóíòåð
    Next cellChecked
    
    'Øàã 3 - Ñîðòèðóåì çíà÷åíèÿ â óáûâàþùåì ïîðÿäêå ïðè ïîìîùè ìåòîäà "Bubble Sort"
    Dim i As Long
    Dim j As Long
    Dim temp As Double
    
    For i = LBound(valuesArray) To UBound(valuesArray) - 1
        For j = i + 1 To UBound(valuesArray)
            If valuesArray(i) < valuesArray(j) Then             'i < j - Óáûâàþùèé ïîðÿäîê.  i > j - Âîçðàñòàþùèé ïîðÿäîê
                
                temp = valuesArray(j)                           'Ïåðåñîðòèðîâûâàåì çíà÷åíèÿ
                valuesArray(j) = valuesArray(i)
                valuesArray(i) = temp
                'Saving new positions
                temp = positionBeforeSortArray(j)               'Ïåðåñîðòèðîâûâàåì íà÷àëüíûå ïîçèöèè çíà÷åíèé
                positionBeforeSortArray(j) = positionBeforeSortArray(i)
                positionBeforeSortArray(i) = temp
                
            End If
        Next j
    Next i
    
    'Øàã 4 - Ñ÷èòàåì îáùóþ ñóììó çíà÷åíèé
    valuesSum = Application.Sum(rangeOfValues)
  
    'Øàã 5 - Ñîõðàíÿåì äîëè çíà÷åíèé â îòäåë¸ííûé ìàññèâ
    For Counter = LBound(valuesArray) To UBound(valuesArray)
        percentageArray(Counter) = valuesArray(Counter) / valuesSum
    Next Counter
    
    'Øàã 6 - Ñîõðàíÿåì íàêîïèòåëüíûå äîëè çíà÷åíèé â îòäåë¸ííûé ìàññèâ
    For Counter = LBound(valuesArray) To UBound(valuesArray)
        If Counter = 1 Then
            cumulativePercentageArray(Counter) = percentageArray(Counter)
        Else:
            cumulativePercentageArray(Counter) = cumulativePercentageArray(Counter - 1) + percentageArray(Counter)
        End If
    Next Counter
    
    
    
    'Øàã 7 - Ïðîâîäèì êëàññèôèêàöèþ è ñîõðàíÿåì çíà÷åíèÿ â ìàññèâ
    For Counter = LBound(valuesArray) To UBound(valuesArray)
        If WorksheetFunction.Round(cumulativePercentageArray(Counter), 2) <= 0.8 Then         'Äåöèìàëüíûé çíàê â VBA - òî÷êà, ïî àìåðèêàíñêèì ñòàíäàðòàì
            classArray(Counter) = "A"
        ElseIf WorksheetFunction.Round(cumulativePercentageArray(Counter), 2) <= 0.95 And WorksheetFunction.Round(cumulativePercentageArray(Counter), 2) > 0.8 Then
            classArray(Counter) = "B"
        Else:
            classArray(Counter) = "C"
        End If
    Next Counter
    
    'Øàã 8 - Íàõîäèì, ñ êàêîé ïî êàêóþ ñòðîêó â ðàáî÷åì ëèñòå íàõîäÿòñÿ êëàññèôèöèðóåìûå çíà÷åíèÿ
    firstRowOfWorkedRangeInExcel = rangeOfValues.Cells(1, 1).Row
    lastRowOfWorkedRangeInExcel = firstRowOfWorkedRangeInExcel + rangeOfValues.Rows.Count - 1
    rowOfCurrentCell = currentCell.Row
    
    
    'Øàã 9 - Âûäàåì èòîãîâûé êëàññ ïðîâåðÿåìîãî çíà÷åíèÿ ïðè ïîìîùè ôóíêöèè ABC_Analysis
    For Counter = LBound(positionBeforeSortArray) To UBound(positionBeforeSortArray)
        If positionBeforeSortArray(Counter) = rowOfCurrentCell - firstRowOfWorkedRangeInExcel + 1 Then
            If ABC_Type = "Letter" Then                           'Åñëè òðåòèé îïöèîíàëüíûé àðãóìåíò = "Letter", òîãäà ëèòåðàëüíûé ôîðìàò
                ABC_Analysis = classArray(Counter)
            ElseIf ABC_Type = "" Then                             'Åñëè òðåòèé îïöèîíàëüíûé àðãóìåíò ïóñò, òîãäà ëèòåðàëüíûé ôîðìàò ïî óìîë÷àíèþ
                ABC_Analysis = classArray(Counter)
            ElseIf ABC_Type = "Cumulative" Then                   'Åñëè òðåòèé îïöèîíàëüíûé àðãóìåíò = "Cumulative", òîãäà íàêîïèòåëüíàÿ äîëÿ
                ABC_Analysis = cumulativePercentageArray(Counter)
            Else:
                ABC_Analysis = CVErr(xlErrName)                   'Ïðè ëþáîì äðóãîì ïîêàçàòåëå òðåòüåãî àðãóìåíòà - îøèáêà!
            End If
        End If
    Next Counter
End If

End Function

'instead of function popups
Sub RegisterUDF()
'Ðåãèñòðèðóåì íàøó User-Defined Function
Dim strFunc As String                                             'Ïåðåìåííàÿ äëÿ èìåíè ôóíêöèè
Dim strDesc As String                                             'Ïåðåìåííàÿ äëÿ îïèñàíèÿ ôóíêöèè
Dim strArgs() As String                                           'Ïåðåìåííàÿ äëÿ îïèñàíèé àðãóìåíòîâ ôóíêöèè

    ReDim strArgs(1 To 3) 'Âåðõíÿÿ ãðàíèöà ìàññèâà ñîîòâåòñòâóåò êîëè÷åñòâó àðãóìåíòîâ íàøåé ôóíêöèè
    strFunc = "ABC_Analysis"
    strDesc = "Ïðèñâàèâàåò êëàññ çíà÷åíèþ â ðàìêàõ ABC-Àíàëèçà. Ðåçóëüòàò " & _
              "ìîæåò âûäàâàòüñÿ â ëèòåðàëüíîì âèäå, ëèáî â âèäå íàêîïèòåëüíîé äîëè. " _
              & "Ïðè èãíîðèðîâàíèè îïöèîíàëüíîãî òðåòüåãî àðãóìåíòà ïî óìîë÷àíèþ âûäàåòñÿ ëèòåðàëüíûé êëàññ."
                           
              
    strArgs(1) = "Îáëàñòü çíà÷åíèé ïîäëåæàùèõ ABC-Àíàëèçó"
    strArgs(2) = "Òåêóùåå çíà÷åíèå (ÿ÷åéêà äîëæíà ïðèíàäëåæàòü îáëàñòè èç ïåðâîãî àðãóìåíòà)"
    strArgs(3) = "Îïöèîíàëüíûé àðãóìåíò - [Letter] âûäàåò ëèòåðàëüíûé êëàññ, [Cumulative] - íàêîïèòåëüíóþ äîëþ. Ïðè èãíîðèðîâàíèè âûäàåòñÿ ëèòåðàëüíûé êëàññ."
    Application.MacroOptions Macro:=strFunc, _
                             Description:=strDesc, _
                             ArgumentDescriptions:=strArgs, _
                             Category:="My Custom Category"
End Sub


