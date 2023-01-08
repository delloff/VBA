

Sub main_procedure()

    Dim arrayToSort(1 To 5) As Long

    arrayToSort(1) = 5
    arrayToSort(2) = 2
    arrayToSort(3) = 4
    arrayToSort(4) = 1
    arrayToSort(5) = 3
    
    'call bubble sort procedure and send data to it using arrayToSort
    Call bubbleSortAux(arrayToSort)
    

End Sub

'launch bubbleSortAux from main_procedure()

Sub bubbleSortAux(list() As Long)

    Dim First As Long
    Dim Last As Long
    Dim i As Long
    Dim j As Long
    Dim temp As Long
    
    First = LBound(list)
    Last = UBound(list)
    
    'two loops one inside the other
    For i = First To Last - 1
        For j = i + 1 To Last
            If list(i) > list(j) Then   'change > with < to determine order
                temp = list(j)
                list(j) = list(i)
                list(i) = temp
            End If
        Next j
    Next i

    For i = 1 To UBound(list)
        Debug.Print list(i)
    Next i

End Sub


'------------------------------------------------------------
'Bubble Sort in Python
'
'def bubbleSort(array):
'
'    for i in range(len(array)-1):
'        for j in range(i+1, (len(array))):
'            if array[i] > array[j]:
'                temp = array[j]
'                array[j] = array[i]
'                array[i] = temp
'    return array
'
'
'Print (bubbleSort([5,2,4,1,3]))

'-----------------------------------------------------------


'BubbleSort HW

Sub fillingTheArray()
    
    'declare array
    Dim Arr() As Long
    Dim Counter As Long
    
    'set the len() of array (XFD1 means the last cell in the 1st row)
    ReDim Arr(1 To Range("XFD1").End(xlToLeft).Column - 1)
    
    'get values into array
    For Counter = 1 To UBound(Arr)
        Arr(Counter) = Cells(1, Counter + 1)
    Next Counter
    
    'Sending the array to a helper procedure
    Call demonstrativeBubbleSort(Arr)
        
End Sub


Sub demonstrativeBubbleSort(listToSort() As Long)
    'Declaring the necessary variables
    Dim First As Long, Last As Long
    Dim i As Long, j As Long
    Dim temp As String
    
    'Determining the first and last indexes of an array
    First = LBound(listToSort)
    Last = UBound(listToSort)
    
    'clearing line 18
    With Range("A18").EntireRow.Cells
        .Clear
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.PatternColorIndex = xlAutomatic
        .RowHeight = 56.25
        .Font.Size = 48
    End With
    
    Range("A18").Font.Size = 18
    Range("A18").WrapText = True
    
    'Filling in the values in the initial order
    For i = First To Last
        Cells(18, i + 1) = listToSort(i)
    Next i
    
    'Creating a graph variable (we will learn in the next video!)
    Dim bubbleSortChart As ChartObject
    Set bubbleSortChart = ThisWorkbook.Worksheets(7).ChartObjects("bubbleSortChart")
    
    'Changing an area with graph data
    bubbleSortChart.Chart.SetSourceData Source:=Range("'BSHW'!B18:" & Cells(18, Last + 1).Address)
    
    'Selecting a standard color for all graph columns
    bubbleSortChart.Chart.FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(91, 155, 213)
    
    'drawing boundary chart for area with sorted values
    With Range("A18:" & Cells(18, Last + 1).Address)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With

    Range("A18") = "Îòñîðòèðîâàííûå çíà÷åíèÿ"
        
    'beginnig the Bubble Sort
    For i = First + 1 To Last - 1 + 1
            
        
        'Change the color of the checking column in the graph to red
        bubbleSortChart.Chart.FullSeriesCollection(1).Points(i - 1).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        'Change the color of the checked cell to red
        Cells(18, i).Interior.Color = 255
        
            For j = i + 1 To Last + 1
                
                'Change the color of the comparison column in the graph to gray
                bubbleSortChart.Chart.FullSeriesCollection(1).Points(j - 1).Format.Fill.ForeColor.RGB = RGB(191, 191, 191)
                'Ìåíÿåì öâåò ñðàâíèâàåìîé ÿ÷åéêè íà ñåðûé
                Cells(18, j).Interior.TintAndShade = -0.149998474074526
                
                'Sort by ascending. If i is greater than j, then swap
                If Cells(18, i) > Cells(18, j) Then
                    temp = Cells(18, j)
                    Cells(18, j) = Cells(18, i)
                    Cells(18, i) = temp
                End If
                                           
                'Returning the standard colors of the compared columns in the graphics and cells in the worksheet
                Cells(18, j).Interior.ThemeColor = xlThemeColorDark1
                bubbleSortChart.Chart.FullSeriesCollection(1).Points(j - 1).Format.Fill.ForeColor.RGB = RGB(91, 155, 213)
            
            Next j
            
        'Returning the standard colors of the check column in the graph and the cell in the worksheet
        Cells(18, i).Interior.PatternColorIndex = xlAutomatic
        Cells(18, i).Interior.ThemeColor = xlThemeColorDark1
        
        'Return blue color for all columns
        bubbleSortChart.Chart.FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(91, 155, 213)
        
    Next i
    
    'Final message
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Çíà÷åíèÿ ïðîñîðòèðîâàíû!", vbOKOnly + vbInformation, "Bubble Sort")
    
    
End Sub


