Attribute VB_Name = "Snake"

Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Private Declare PtrSafe Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Long

    ' Direction
    ' 1 - up
    ' 2 - down
    ' 3 - left
    ' 4 - right
    Dim move As Boolean
    Dim direction As Integer

Sub StartSnake(ByVal difficulty As Integer)
    Call Utils.Init

    ActiveSheet.Name = "Snake"

    Range("A1:S19").Interior.ColorIndex = 2

    Range("A1:B19").ColumnWidth = 1.88
    Range("A1:B19").RowHeight = 15
    Range("A1:S2").ColumnWidth = 1.88
    Range("A1:S2").RowHeight = 15
    Range("R1:S19").ColumnWidth = 1.88
    Range("R1:S19").RowHeight = 15

    Dim board As Range: Set board = Range("C3:Q17")
    board.ColumnWidth = 3.13
    board.RowHeight = 22.5
    board.Borders.LineStyle = xlContinuous
    board.HorizontalAlignment = xlCenter
    board.VerticalAlignment = xlCenter

    Range("J10").Interior.ColorIndex = 11
    Range("J10").Select

    Call PutFood()

    ActiveSheet.Protect "password"


    Application.OnKey "{UP}", "dummy"
    Application.OnKey "{DOWN}", "dummy"
    Application.OnKey "{LEFT}", "dummy"
    Application.OnKey "{RIGHT}", "dummy"

    MsgBox "↑ ↓ ← → または、 W S A D で移動できます", , "スネークプレイ方法"

    direction = 2
    move = True

    Dim speed As Integer
    If (difficulty = 0) Then
        speed = 200
    ElseIf (difficulty = 1) Then
        speed = 150
    Else
        speed = 100
    End If

    Call moveSnake(speed)
End Sub

Private Sub moveSnake(ByVal speed As Integer)
    Dim snakelength As Integer: snakelength = 0
    Dim snake() As Range
    ReDim Preserve snake(0)
    Set snake(snakelength) = Range("J10")

    Dim head As Range
    Dim newHead As Range
    Dim tail As Range
    Dim i As Integer

    Do While move = True
        DoEvents
        Call CheckDirectionPress()
        DoEvents

        Set head = snake(0)
        Select Case direction
         Case 1
            Set newHead = head.Offset(-1, 0)
         Case 2
            Set newHead = head.Offset(1, 0)
         Case 3
            Set newHead = head.Offset(0, -1)
         Case 4
            Set newHead = head.Offset(0, 1)
        End Select

        ' You bump into yourself
        If (newHead.Interior.ColorIndex = 11) Then
            Call Finish()
            Exit Sub
        End If

        ' Too bad, you crash to the wall
        If Not Intersect(Range("C2:Q2"), newHead) Is Nothing Then
            Call Finish()
            Exit Sub
        ElseIf Not Intersect(Range("R3:R17"), newHead) Is Nothing Then
            Call Finish()
            Exit Sub
        ElseIf Not Intersect(Range("B3:B17"), newHead) Is Nothing Then
            Call Finish()
            Exit Sub
        ElseIf Not Intersect(Range("C18:Q18"), newHead) Is Nothing Then
            Call Finish()
            Exit Sub
        End If

        ReDim Preserve snake(snakelength + 1)
        For i = snakeLength To 0 Step -1
            Set snake(i + 1) = snake(i)
        Next i
        Set snake(0) = newHead

        Set tail = snake(snakeLength + 1)

        ActiveSheet.UnProtect "password"

        if (newHead.Interior.ColorIndex = 3) Then
            snakelength = snakelength + 1
            Call PutFood()

            DoEvents
            Call CheckDirectionPress()
            DoEvents
        Else
            tail.Interior.ColorIndex = 2
            Redim Preserve snake(snakeLength)
        End If

        newHead.Interior.ColorIndex = 11
        newHead.Select

        ActiveSheet.Protect "password"

        DoEvents
        Call CheckDirectionPress()
        DoEvents

        Sleep speed
    Loop
End Sub

Private Sub CheckDirectionPress()
    If GetAsyncKeyState(&H26) <> 0 Then ' ↑
        If (direction <> 2) Then
            direction = 1
        End If
    ElseIf GetAsyncKeyState(&H28) <> 0 Then ' ↓
        If direction <> 1 Then
            direction = 2
        End If
    ElseIf GetAsyncKeyState(&H25) <> 0 Then ' ←
        If (direction <> 4) Then
            direction = 3
        End If
    ElseIf GetAsyncKeyState(&H27) <> 0 Then ' →
        If (direction <> 3) Then
            direction = 4
        End If
    ElseIf GetAsyncKeyState(&H57) <> 0 Then ' w
        If (direction <> 2) Then
            direction = 1
        End If
    ElseIf GetAsyncKeyState(&H53) <> 0 Then ' s
        If direction <> 1 Then
            direction = 2
        End If
    ElseIf GetAsyncKeyState(&H41) <> 0 Then ' a
        If (direction <> 4) Then
            direction = 3
        End If
    ElseIf GetAsyncKeyState(&H44) <> 0 Then ' d
        If (direction <> 3) Then
            direction = 4
        End If
    End If
End Sub

Private Sub PutFood()
    Dim x As Integer
    Dim y As Integer
    Dim putFood As Boolean: putFood = False

    Do While putFood = False
        x = Int(Rnd * 15) + 2 + 1
        y = Int(Rnd * 15) + 2 + 1
        If (Cells(x, y).Interior.ColorIndex = 2) Then
            Cells(x, y).Interior.ColorIndex = 3
            putFood = True
        End If
    Loop
End Sub

Private Sub Finish()
    move = False
    MsgBox "ゲームオーバー", , "ゲーム終了"

    MainMenu.Show
End Sub

' For debug purpose
Private Sub reset()
    Call StartSnake(0)
End Sub

Private Sub forceStop()
    move = False
End Sub
