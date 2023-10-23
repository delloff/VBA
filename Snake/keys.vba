Option Explicit

Sub bindKeys()
    Application.ScreenUpdating = False
    Application.OnKey "{LEFT}", "moveLeft"
    Application.OnKey "{UP}", "moveUp"
    Application.OnKey "{DOWN}", "moveDown"
    Application.OnKey "{RIGHT}", "moveRight"
    Application.OnKey "{ESC}", "StopGame"
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Private Sub moveLeft()
If GameStarted = True And OnMove = False And cinc <> 1 Then
    cinc = -1
    rinc = 0
    OnMove = True
End If
End Sub

Private Sub moveUp()
If TimerID = 0 Then
  GameStarted = True
  StartTimer
End If
If GameStarted = True And OnMove = False And rinc <> 1 Then
    cinc = 0
    rinc = -1
    OnMove = True
End If
End Sub

Private Sub moveRight()
If GameStarted = True And OnMove = False And cinc <> -1 Then
    cinc = 1
    rinc = 0
    OnMove = True
End If
End Sub

Private Sub moveDown()
If GameStarted = True And OnMove = False And rinc <> -1 Then
    cinc = 0
    rinc = 1
    OnMove = True
End If
End Sub

Sub freeKeys()
    Application.OnKey "{LEFT}"
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{ESC}"
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
