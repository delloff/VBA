'As alternative to if-construction
'works as with number as with strings


Sub Urok20_1()

Dim s As String

s = "a"

Select Case s
    Case "A", "a"
        Debug.Print "this is A"
    Case "B"
        Debug.Print "this is B"
    Case Else
        Debug.Print "this is smth different"
End Select

End Sub


Sub Urok20_2()

Dim s As Long

s = -10

Select Case s
    Case 1 To 6
        Debug.Print s + s
    Case Is > 7
        Debug.Print s
    Case Else
        Debug.Print "this is smth different"
End Select

End Sub

'VBA knows order of the alphabet

Sub Urok20_3()

Dim s1 As String
Dim s2 As String
Dim s3 As String

s1 = "A"
s2 = "B"
s3 = "C"

Select Case s2
    Case Is < s2
        Debug.Print "s1 goes before s2 in alphabet"
    Case Is > s2
        Debug.Print "s3 goes after s2 in alphabet"
    Case Else
        Debug.Print "s2 is s2"
End Select

End Sub
