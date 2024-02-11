'Sub PreventScreenSaving()
'    Do
'        ' Simulate user activity here
'        Application.SendKeys "{NUMLOCK}"
'
'        ' Wait for 5 minutes
'        Application.Wait Now + TimeValue("00:05:00")
'    Loop
'End Sub
'


Option Explicit

Private Declare PtrSafe Function SetThreadExecutionState Lib "kernel32" (ByVal esFlags As Long) As Long
Private Const ES_CONTINUOUS As Long = &H80000000
Private Const ES_DISPLAY_REQUIRED As Long = &H2

Sub PreventScreenSaving()
    SetThreadExecutionState ES_CONTINUOUS Or ES_DISPLAY_REQUIRED
End Sub

Sub AllowScreenSaving()
    SetThreadExecutionState ES_CONTINUOUS
End Sub

