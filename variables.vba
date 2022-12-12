Option Explicit

'ÏÅÐÅÌÅÍÍÛÅ È ÒÈÏÛ ÄÀÍÍÛÕ

Sub Urok3()

    'BASIC VARIABLES
    Dim someText As String
    Dim num1 As Long
    Dim someDate As Date
    
    someText = "Èòîã: "
    num1 = 100
    someDate = "11.12.2022"
    
    Dim num2 As Long
    
    num2 = (10 + 10) * 5 + num1
    
    MsgBox someText & num2
    
    
'    ïåðåìåííûíå ïåðåçàïèñûâàþòñÿ êàê â ïèòîíå:
    num2 = 4
    MsgBox num2
    
End Sub
    
Sub Urok3_4()

    'OBJECT VARIABLE
    Dim someWorkbook As Workbook
    Dim someWorksheet As Worksheet
    Dim someRange As Range
    
    Set someWorkbook = ThisWorkbook
    Set someWorksheet = Worksheets("Sheet1")
    Set someRange = Range("C3")
    
    MsgBox someRange.Address
    
    
        
End Sub



'ÒÈÏÛ ÄÀÍÍÛÕ:
'Integer,Long - äëÿ öåëûõ ÷èñåë. Ìîæíî èñïîëüçîâàòü âñåãäà Long
'Double - äëÿ ÷èñåë ñ çàïÿòîé, òàê êàê Long áóäåò èõ îêðóãëÿòü
'String
'Boolean
'Date


Sub Urok4()

    Dim num1 As Long
    num1 = 2.7
    
    MsgBox num1
    
End Sub
