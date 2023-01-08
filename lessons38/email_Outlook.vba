'sheet "find"
'we can add libraries to work with third party programms
'microsoft scripting runtime (used for dict)
'to do this go to tools->references
'we can now use intelliSence with added library
'odjects can belong to third party libraries. We need to add them this way (Outlook etc)
'Erly binding: when we add library via Tools->References
'Late binding: when we work without it


'EARLY BINDING
'if used on different pc, outlook versions must match.
'As workaround I can write code using early binding, but change it to late binding when finished before distribution

Sub Urok34_1()
    
    '---------------------------EARLY BINDING-------------------------------
    'declare variable for outlook object(programm) & email object          '|
    Dim appOutlook As Outlook.Application                                  '|
    Dim newMail As Outlook.MailItem                                        '|
                                                                           '|
    'link declared variable to real object instance                        '|
    'we run outlook it pc memory                                           '|
    Set appOutlook = New Outlook.Application                               '|
    '-----------------------------------------------------------------------
    
    'create instance of email in pc memory and link to it variable appOutlook
    Set newMail = appOutlook.CreateItem(olMailItem)
    
    'output email
    newMail.display
    
    'adjust settings of email
    newMail.To = "johnny.dollar@outlook.com"
    newMail.Subject = "hello from VBA"
    newMail.Body = "message after disabling outlook warning message"
    
    'sending emai
'    newMail.Send
'
'    'close without sending and saving
'    newMail.Close olDiscard
'
'    'clearing variables
'    Set appOutlook = Nothing
'    Set newMail = Nothing

End Sub


'LATE BINDING (turn off Outlook library to work)
'if used on different pc, outlook versions CAN BE different
'intelliSence not supported


Sub Urok34_2()
    
    '---------------------------LATE BINDING-------------------------------
    'declare variable for outlook object(programm) & email object          '|
    Dim appOutlook As Object                                               '|
    Dim newMail As Object                                                  '|
                                                                           '|
    'link declared variable to real object instance                        '|
    'we run outlook it pc memory                                           '|
    Set appOutlook = CreateObject("Outlook.Application")                   '|
    '-----------------------------------------------------------------------
    
    'create instance of email in pc memory and link to it variable appOutlook
    Set newMail = appOutlook.CreateItem(0)  'google: OlItemType enumeration (Outlook)
    
    'output email
    newMail.display
    
    'adjust settings of email
    newMail.To = "johnny.dollar@outlook.com"
    newMail.Subject = "hello from VBA"
    newMail.Body = "message after disabling outlook warning message"
    
    'sending emai
'    newMail.Send
'
'    'close without sending and saving
'    newMail.Close (1) 'google: OlInspectorClose enumeration (Outlook)
'
'    'clearing variables
'    Set appOutlook = Nothing
'    Set newMail = Nothing

End Sub

'HW for Urok34

Sub mainSub_loop()
    
    Dim rgCellChecked As Range
    
    For Each rgCellChecked In ThisWorkbook.Worksheets("find").Range("F7:F11")
        Call helpereSub_SendAnEmail(rgCellChecked)
    Next rgCellChecked
    
End Sub


Sub helpereSub_SendAnEmail(ByVal rgCustomerProcessed As Range)
    
    '---------------------------EARLY BINDING-------------------------------
    'declare variable for outlook object(programm) & email object          '|
    Dim appOutlook As Outlook.Application                                  '|
    Dim newMail As Outlook.MailItem                                        '|
                                                                           '|
    'link declared variable to real object instance                        '|
    'we run outlook it pc memory                                           '|
    Set appOutlook = New Outlook.Application                               '|
    '-----------------------------------------------------------------------
    
    'create instance of email in pc memory and link to it variable appOutlook
    Set newMail = appOutlook.CreateItem(olMailItem)
    
    'output email
    newMail.display
    
    'adjust settings of email
    newMail.To = rgCustomerProcessed
    newMail.Subject = "Rent of " & rgCustomerProcessed.Offset(0, 4)
    
    
    newMail.Body = "Dear " & rgCustomerProcessed.Offset(0, 1) & " " & rgCustomerProcessed.Offset(0, 2) & "," _
    & vbNewLine & vbNewLine & "With this letter we confirm the succsessful rent of " & rgCustomerProcessed.Offset(0, 4) _
    & " from " & rgCustomerProcessed.Offset(0, 6) & " till " & rgCustomerProcessed.Offset(0, 7) _
    & ". Total rent cost $" & rgCustomerProcessed.Offset(0, 9) & " for " & rgCustomerProcessed.Offset(0, 8) _
    & " days." & vbNewLine & vbNewLine & "Thanks for your order!" & vbNewLine & vbNewLine & "BR" _
    & vbNewLine & "Your CarRentService"
    
        
    'sending emai
'    newMail.Send
'
'    'close without sending and saving
'    newMail.Close olDiscard
'
'    'clearing variables
'    Set appOutlook = Nothing
'    Set newMail = Nothing


End Sub


