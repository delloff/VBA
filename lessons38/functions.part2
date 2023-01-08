'sheet "MultiArr1", lesson 35
'extract numbers from string
'mid(string,pos,len from that pos to extract) = extracts part of string from given string
'mid(Test 123,7,1) -> returns 2
'Right(str,pos)/Left(str,pos) - similar to python slice


'feed information in brackets to function (byVal because s will be changed)

Function extractNums(ByVal s As String) As String
    
    'for storing nums. In python we would write res = ''
    Dim res As String
    'for len of string
    Dim sLength As Long
        sLength = Len(s)
    'for counting pos of string
    Dim i As Long
    
    
    If s <> "" Then
        'checking each pos in string
        For i = 1 To sLength
            If IsNumeric(Mid(s, i, 1)) Then
                res = res & Mid(s, i, 1)
                'if next pos is a letter add "_"
                If i + 1 <= sLength Then
                    If Not IsNumeric(Mid(s, i + 1, 1)) Then
                        res = res & "_"
                    End If
                End If
            End If
        Next i
        'remove last "_" if exists. In python we would do res[:-1]
        If Right(res, 1) = "_" Then
            res = Left(res, Len(res) - 1)
        End If
        'returning res
        extractNums = res
    Else
        res = ""
    End If
    
    
End Function

Sub testExtractNums()

    Debug.Print extractNums("Test 123")

End Sub



Function GetNumeric(CellRef As String)
    Dim StringLength As Integer
    StringLength = Len(CellRef)
    
        For i = 1 To StringLength
        If IsNumeric(Mid(CellRef, i, 1)) Then Result = Result & Mid(CellRef, i, 1)
        Next i
    GetNumeric = Result
End Function
