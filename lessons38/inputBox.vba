Option Explicit

    Const constInputBoxType_Formula As Integer = 0
    Const constInputBoxType_Number As Integer = 1
    Const constInputBoxType_Text As Integer = 2
    Const constInputBoxType_Boolean As Integer = 4
    Const constInputBoxType_Range As Integer = 8
    Const constInputBoxType_Error As Integer = 16
    Const constInputBoxType_Array As Integer = 64

'application input box (has more functionality)

Sub Urok30_1()

    'create variable where we store info from inputbox
    Dim name As String

    name = Application.InputBox("Enter Something")
    MsgBox ("You entered: " & name)

End Sub

'vba input box  (has less functionality)

Sub Urok30_2()

    'create variable where we store info from inputbox
    Dim name As String

    name = VBA.InputBox("Enter Something")
    MsgBox ("You entered: " & name)

End Sub

'testing constInputBoxType_Number
'you can enter number or click on any cell
'HOW to avoid eerror on cancel?
'Application.DisplayAlerts = False/True

Sub Urok30_3()  'with numbers

    Dim varNumberVariable As Variant    'because user can click on "cancel" and it will be false
    

    varNumberVariable = Application.InputBox(prompt:="Test number", Type:=constInputBoxType_Number)
    

    If TypeName(varNumberVariable) = "Boolean" And varNumberVariable = False Then
        MsgBox ("Cancelled")
    Else
        MsgBox ("You entered: " & varNumberVariable)
        
    End If
    
End Sub

Sub Urok30_4()  'with text

    Dim varTextVariable As Variant    'because user can click on "cancel" and it will be false
    

    varTextVariable = Application.InputBox(prompt:="Test text", Type:=constInputBoxType_Text)
    

    If TypeName(varTextVariable) = "Boolean" And varTextVariable = False Then
        MsgBox ("Cancelled")
    Else
        MsgBox ("You entered: " & varTextVariable)
        
    End If
    
End Sub

'it will throw error when we cancel.
'To avoid use: On Error Resume Next -> lines with error -> On Error GoTo 0

Sub Urok30_5()  'with range
    
    Dim rgInput As Range    'because user can click on "cancel" and it will be false
    
    On Error Resume Next
    Set rgInput = Application.InputBox("Test range", Type:=constInputBoxType_Range)
    On Error GoTo 0
    
    
    If rgInput Is Nothing Then
        MsgBox ("choosing range was cancelled")
    Else
        MsgBox ("You entered: " & rgInput.Address)
    End If
    
End Sub

'HW on sheet called "Find"

Sub Urok30_61()  'color the values in range
    
    Dim rgInput As Range    'because user can click on "cancel" and it will be false
    
    On Error Resume Next
    Set rgInput = Application.InputBox("Test range", Type:=constInputBoxType_Range)
    On Error GoTo 0
    
    Dim cellName As Range
    
    
    If rgInput Is Nothing Then
        MsgBox ("choosing range was cancelled")
    Else
        For Each cellName In rgInput
            If cellName.Value > 0 Then
                cellName.Font.Color = RGB(50, 205, 50)
            ElseIf cellName.Value < 0 Then
                cellName.Font.Color = RGB(255, 0, 0)
            Else
                cellName.Font.Color = RGB(115, 115, 115)
            End If
        Next cellName
    End If
    
End Sub

'HW on sheet called "Find"

Sub Urok_30_62()   'remove colors in range

    Dim rgInput As Range
    
    On Error Resume Next
    Set rgInput = Application.InputBox("Test range", Type:=constInputBoxType_Range)
    On Error GoTo 0
    
    Dim cellName As Range
    
    
    If rgInput Is Nothing Then
        MsgBox ("choosing range was cancelled")
    Else
        For Each cellName In rgInput
            cellName.Font.ColorIndex = xlAutomatic
        Next cellName
    End If
    
End Sub


