'the name of the module shoulbn't be msgBox becasue we have same names (module and excel function)


Sub Urok_21_1()

Dim userresponse As VbMsgBoxResult  'this variable holds user responses: vbOk = 1, vbCance = 2 etc (google it)

    If MsgBox("Hello Andrei", vbYesNo) = vbYes Then
        MsgBox "you enered Yes"
    Else
        MsgBox "You entered No"
    End If

End Sub



Sub Urok_21_2()

Dim userresponse As VbMsgBoxResult

userresponse = MsgBox("Hello Andrei", vbAbortRetryIgnore)

    If userresponse = vbAbort Then
        MsgBox "process aborted"
    ElseIf userresponse = vbRetry Then
        MsgBox "process retried"
    ElseIf userresponse = vbIgnore Then
        MsgBox "process ignored"
    End If

End Sub

'we can combine properties, one property of a each group. The order is not important

Sub Urok_21_3()

Dim userresponse As VbMsgBoxResult

userresponse = MsgBox("Testing MsgBox", vbYesNoCancel + vbInformation + vbMsgBoxHelpButton _
+ vbSystemModal, "Andrei testing message window")

End Sub


Sub helpFile()
    Dim lngQ As Long
    lngQ = MsgBox("Do you want to delete this workbook?", vbYesNo _
        + vbCritical + vbDefaultButton2 + vbMsgBoxHelpButton, _
        "Delete Workbook", "C:\Users\User\Desktop\help.txt", 1012)
End Sub
