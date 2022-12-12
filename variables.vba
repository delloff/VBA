Option Explicit

'VARIABLES AND DATA TYPES

Sub Urok3()

    'BASIC VARIABLES
    Dim someText As String
    Dim num1 As Long
    Dim someDate As Date
    
    someText = "Eoia: "
    num1 = 100
    someDate = "11.12.2022"
    
    Dim num2 As Long
    
    num2 = (10 + 10) * 5 + num1
    
    MsgBox someText & num2
    
    
'   variables are overwritten like in python:
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



'DATA TYPES:
'Integer,Long - for integers, long is more commonly used
'Double - for floating point numbers. If you use Long the number will be rounded
'String
'Boolean
'Date


Sub Urok4()

    Dim num1 As Long
    num1 = 2.7
    
    MsgBox num1
    
End Sub

