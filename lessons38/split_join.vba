'split
'join

'SPLIT

Sub Urok38_1()

   Dim strArr() As String
   'save values into arr. Add watch to see
    'strArr = Split("Russia,Belarus,Germany,France,Holland", ",")
    
    'we can set limit of array. It means the all words that left will be saved under one last index
    'if there are 5 words and we set limit to 10, there will be just 5 indexes created
'    strArr = Split("Russia,Belarus,Germany,France,Holland", ",", 4)
    
    'vbBinaryCompare/vbTextCompare - case sensitive/insensitive
    strArr = Split("100z200z300Z400z500", "z", -1, vbBinaryCompare)
    
End Sub


'JOIN


Sub Urok38_2()

    Dim sourceArr As Variant
    
    'put values in it. By default 2D array is created
    sourceArr = ThisWorkbook.Worksheets("MultiArr1").Range("A15:A21")
    'make 1D array by exchanging two arrays and deleting last one
    sourceArr = Application.WorksheetFunction.Transpose(sourceArr)
    
    Dim res As String
    res = Join(sourceArr, ", ")
    
    ThisWorkbook.Worksheets("MultiArr1").Range("A35") = res

End Sub


'HW
'extract just "Screenshot_1" from string

Sub Urok38_3()

    Dim path As String
    path = "C:\Users\User\Desktop\Screenshot_1.png"
    
    'get rid of xlsx extension
    Dim strArr() As String
    strArr = Split(path, ".")
    
    'now work with first part of array (untill ".")
     strArr = Split(strArr(0), "\")
    'the same as:
    'strArr = Split(strArr(0), Application.PathSeparator)
    
    Debug.Print strArr(UBound(strArr)) 'or Debug.Print strArr(4)
      

End Sub
