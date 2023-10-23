Public rinc As Integer, cinc As Integer, speed As Integer, OnMove As Boolean, GameStarted As Boolean
Dim rlim As Integer, clim As Integer, score As Integer, GameWs As Worksheet
Dim r() As Integer, c() As Integer

Sub StartGame()
  Dim i As Integer
  GameStarted = False
  Set GameWs = Sheets("SNAKE")
  
  'Format layout
  With GameWs.Range("B2:AU30")
    .Interior.Color = vbBlack
    .Borders.Color = vbBlack
    .BorderAround LineStyle:=xlContinuous, Color:=vbRed
  End With
  GameWs.Range("BC5").Select
  
  rlim = 30: clim = 47 'limits of the grid (start row and col is 2)
  rinc = 0: cinc = 0
  
  'Declare Snake array variables
  ReDim r(4)
  ReDim c(4)
  
  'Assign initial position values to array
  For i = 0 To UBound(r) - 1
    r(i) = (rlim - 3) + i
    c(i) = 24
    GameWs.Cells(r(i), c(i)).Interior.Color = vbGreen
  Next i
  GameWs.Cells(r(0), c(0)).Interior.Color = vbRed

  score = 0
  UpdateScore
  PlaceApple
  Keys.bindKeys
  speed = 80
  OnMove = False
End Sub

Sub MoveSnake()
 Dim i As Integer, tail As Integer
 

 If GameStarted = True Then
    With GameWs
      If .Cells(r(0) + rinc, c(0) + cinc).Interior.Color = vbBlack _
      Or .Cells(r(0) + rinc, c(0) + cinc).Interior.Color = vbYellow Then
        'Loop through array to assign previous value
        tail = UBound(r)
        For i = tail To 1 Step -1
          r(i) = r(i - 1)
          c(i) = c(i - 1)
        Next i
        .Cells(r(tail), c(tail)).Interior.Color = vbBlack
        .Cells(r(0), c(0)).Interior.Color = vbGreen
        
        'advance the head of the snake
        r(0) = r(0) + rinc
        c(0) = c(0) + cinc
   
        'check if catches apple
        If .Cells(r(0), c(0)).Interior.Color = vbYellow Then
          ReDim Preserve r(UBound(r) + 1)
          ReDim Preserve c(UBound(c) + 1)
          score = score + 10
          UpdateScore
          PlaceApple
        End If
        
        .Cells(r(0), c(0)).Interior.Color = vbRed 'fill head of snake
      Else
        StopGame
        MsgBox "Game Over"
      End If
    End With
    OnMove = False
  End If
End Sub

Private Sub PlaceApple()
  Dim AppleRow As Integer, AppleCol As Integer
  Randomize
  AppleRow = Int((Rnd * (rlim - 2)) + 2)
  AppleCol = Int((Rnd * (clim - 2)) + 2)
  If GameWs.Cells(AppleRow, AppleCol).Interior.Color = vbBlack Then
    With GameWs.Cells(AppleRow, AppleCol)
      .Interior.Color = vbYellow
    End With
  Else
    PlaceApple
  End If
End Sub

Private Sub UpdateScore()
  GameWs.Shapes("ScoreBox").TextFrame.Characters.Text = score
End Sub

Sub StopGame()
  GameStarted = False
  Timer.StopTimer
  Keys.freeKeys
End Sub
