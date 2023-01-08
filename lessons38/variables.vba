'declaration section here
'we can declare public variables here to use the in all modules
'here we can also declare private variables only for this module using Dim or Private keyword:

'Dim somePrivateName as String
'is equal to:
'Private somePrivateName as String

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
    
    
    'variables are rewritable like in python:
    num2 = 4
    MsgBox num2
    'variable are cleared after procedure ends. Except private variables that are used in many procedures
    'we can however hold the value till the programm close using static keyword (see last procedure)
    
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
'Double - äëÿ ÷èñåë ñ çàïÿòîé, òàê êàê Long áóäåò èõ îêðóãëÿòü. Òàêæå äëÿ î÷åíü áîëüøèõ ÷èñåë
'String
'Boolean
'Date


Sub Urok4()

    Dim num1 As Long
    num1 = 2.7
    
    MsgBox num1
    
End Sub

'Const
Sub constTest()

'constants always stay the same. it has to be assigned a value when we declare it
Const aConst As Integer = 5

End Sub



Sub staticTest()

    Dim intDimVar As Integer
    Static intStaticVar As Integer
    
    'always output 1
    intDimVar = intDimVar + 1
    MsgBox "normal var " & intDimVar
    
    'it will increment +1 each time we run the procedure (add watch to see) until we stop/reset the procedure
    intStaticVar = intStaticVar + 1
    MsgBox "static var " & intStaticVar

End Sub

'Variables can be
'simple: Dim someVar as Smth (store an object)
'object variables (just link to object). Thus we can have many such variables linking to one single object

'simple variable
Sub Urok32_1()

    Dim someStr As String
    someStr = ThisWorkbook.Worksheets(1).Range("A1")
    
    MsgBox someStr

End Sub

'object variable
'What are OBJECTS in VBA?
'Objects are structured/grouped data.
'And we can perform certain tasks that are supported by those objects itself (properties).
'object types: cells, worksheets, workbooks, forms, collections, rows, columns etc.
'so we dont create those objects, just link variables to already existing object.
'we can also create our objects: Dim someObject as Object
'But when linking to a collection we first need to create it


'lets look at worksheet as an object


Sub Urok32_2()
    
    'first declare
    Dim wsWorksheet As Worksheet
    Dim wsWorksheet2 As Worksheet
    
    'then create link to it using set
    Set wsWorksheet = ThisWorkbook.Worksheets(1)
    Set wsWorksheet2 = wsWorksheet
    
    'now we can perform different tasks on it depending on its properties
    MsgBox wsWorksheet.name
    MsgBox wsWorksheet2.name
    
    
End Sub


'lets look at Collection as an object

Sub Urok32_3()

    'declare
    Dim testColl As Collection
    
    'create Collection and link to it
    Set testColl = New Collection
    
    'it is better practice to merge above two lines in one line. Normally Set without New
    'Dim testColl As New Collection
    
    'perform different tasks on our Collection (adding values)
    testColl.Add 1
    testColl.Add 2
    testColl.Add 3
    
    'we can also remove links from such variables. Our variable becomes empty
    Set testColl = Nothing
    
End Sub


'create our own object

Sub Urok32_4()

    'declare
    Dim testObjectVar As Object
    
    'link to cell and use its methods
    Set testObjectVar = Range("A1")
    MsgBox testObjectVar.Address
    
    'unlink and then link it to different object, use its methods. However, it is not good to do it in code
    Set testObjectVar = Nothing
    Set testObjectVar = Worksheets(1)
    MsgBox testObjectVar.name

End Sub
