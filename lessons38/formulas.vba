Sub Urok16_1()

    Range("C3") = 123
    MsgBox Range("C3")
    Range("C3") = None

End Sub


Sub Urok16_2()

    Range("C3") = 123
    Range("C4") = 456
    
    
    Range("C5").Formula = "=SUM(C3:C4)"

End Sub


