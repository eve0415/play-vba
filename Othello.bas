Attribute VB_Name = "Othello"

Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

    Dim vsCPU As Boolean
    Dim isBlackTurn As Boolean
    Dim isSkippedBefore As Boolean
    Dim firstPlayer As Integer ' 0: Black, 1: White

Sub StartOthello(ByVal vs As Integer, ByVal attackFirst As Boolean)
    isBlackTurn = True
    If (vs = 0) Then
        vsCPU = False
    Else
        vsCPU = True
    End If
    If (attackFirst) Then
        firstPlayer = 0
    Else
        firstPlayer = 1
    End If

    Call Utils.Init

    ActiveSheet.Name = "Othello"

    Dim board As Range: Set board = Range("A1:H8")
    board.ColumnWidth = 5.63
    board.RowHeight = 37.5
    board.Font.Size = 36
    board.Interior.ColorIndex = 10
    board.Borders.LineStyle = xlContinuous
    board.HorizontalAlignment = xlCenter
    board.VerticalAlignment = xlCenter

    Range("I1:M8").ColumnWidth = 8.38

    Range("D4:E5").Value = "●"
    Range("D4").Font.ColorIndex = 2
    Range("E5").Font.ColorIndex = 2

    Call countStone

    With Range("K3")
        .Value = "手番"
        .Font.ColorIndex = 1
        .Font.Size = 26
        .Interior.ColorIndex = 8
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Range("L3")
        .Value = "●"
        .Font.ColorIndex = 1
        .Font.Size = 36
        .Interior.ColorIndex = 10
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ActiveSheet.Protect "password"

    Range("L3").Select

    MsgBox "セルを選んで Enter キーを入力してください", , "オセロプレイ方法"

    If (vsCPU And attackFirst = False) Then
        Call RunAI
    Else
        Application.OnKey "~", "putStone"
    End If
End Sub

Private Sub putStone()
    If Intersect(Range("A1:H8"), ActiveCell) Is Nothing Then
        MsgBox "範囲外です", , "オセロシステム"
        Exit Sub
    End If

    If (ActiveCell.Value <> "") Then
        MsgBox "既に石が置かれています", , "オセロシステム"
        Exit Sub
    End If

    Dim i As Integer
    Dim placeable() As Range
    placeable = getPlaceable()

    For i = 0 To UBound(placeable)
        If (placeable(i).Row = ActiveCell.Row And placeable(i).Column = ActiveCell.Column) Then
            ActiveSheet.UnProtect "password"

            Call placeStone(ActiveCell, False)
            Call changeTurn
            ActiveSheet.Protect "password"
            Exit Sub
        End If
    Next i

    MsgBox "ここに石を置くことができません", , "オセロシステム"
End Sub

Private Function placeStone(ByVal cell As Range, Optional ByVal dry As Boolean = False) As Integer
    Dim i As Integer
    Dim j As Integer
    Dim total As Integer: total = 0
    Dim leftUp As Integer: leftUp = 0
    Dim upI As Integer: upI = 0
    Dim rightUp As Integer: rightUp = 0
    Dim rightI As Integer: rightI = 0
    Dim rightDown As Integer: rightDown = 0
    Dim downI As Integer: downI = 0
    Dim leftDown As Integer: leftDown = 0
    Dim leftI As Integer: leftI = 0

    ' Left Up
    For i = 1 To 7
        If (cell.Row - i < 1 Or cell.Column - i < 1) Then
            Exit For
        End If

        If (cell.Offset(-i, -i).Value = "") Then
            Exit For
        End If

        If (cell.Offset(-i, -i).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            leftUp = i
            total = i
            Exit For
        End If
    Next i

    ' Up
    For i = 1 To 7
        If (cell.Row - i < 1) Then
            Exit For
        End If

        If (cell.Offset(-i, 0).Value = "") Then
            Exit For
        End If

        If (cell.Offset(-i, 0).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            upI = i
            total = i + total
            Exit For
        End If
    Next i

    ' Right Up
    For i = 1 To 7
        If (cell.Row - i < 1 Or cell.Column + i > 8) Then
            Exit For
        End If

        If (cell.Offset(-i, i).Value = "") Then
            Exit For
        End If

        If (cell.Offset(-i, i).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            rightUp = i
            total = i + total
            Exit For
        End If
    Next i

    ' Right
    For i = 1 To 7
        If (cell.Column + i > 8) Then
            Exit For
        End If

        If (cell.Offset(0, i).Value = "") Then
            Exit For
        End If

        If (cell.Offset(0, i).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            rightI = i
            total = i + total
            Exit For
        End If
    Next i

    ' Right Down
    For i = 1 To 7
        If (cell.Row + i > 8 Or cell.Column + i > 8) Then
            Exit For
        End If

        If (cell.Offset(i, i).Value = "") Then
            Exit For
        End If

        If (cell.Offset(i, i).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            rightDown = i
            total = i + total
            Exit For
        End If
    Next i

    ' Down
    For i = 1 To 7
        If (cell.Row + i > 8) Then
            Exit For
        End If

        If (cell.Offset(i, 0).Value = "") Then
            Exit For
        End If

        If (cell.Offset(i, 0).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            downI = i
            total = i + total
            Exit For
        End If
    Next i

    ' Left Down
    For i = 1 To 7
        If (cell.Row + i > 8 Or cell.Column - i < 1) Then
            Exit For
        End If

        If (cell.Offset(i, -i).Value = "") Then
            Exit For
        End If

        If (cell.Offset(i, -i).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            leftDown = i
            total = i + total
            Exit For
        End If
    Next i

    ' Left
    For i = 1 To 7
        If (cell.Column - i < 1) Then
            Exit For
        End If

        If (cell.Offset(0, -i).Value = "") Then
            Exit For
        End If

        If (cell.Offset(0, -i).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            leftI = i
            total = i + total
            Exit For
        End If
    Next i

    placeStone = total

    If (Not dry) Then
        cell.Value = "●"
        cell.Font.ColorIndex = getStoneColor()
        Sleep 150

        For i = 1 To WorksheetFunction.max(leftUp, upI, rightUp, rightI, rightDown, downI, leftDown, leftI)
            If (i <= leftUp) Then
                cell.Offset(-i, -i).Font.ColorIndex = getStoneColor()
            End If
            If (i <= upI) Then
                cell.Offset(-i, 0).Font.ColorIndex = getStoneColor()
            End If
            If (i <= rightUp) Then
                cell.Offset(-i, i).Font.ColorIndex = getStoneColor()
            End If
            If (i <= rightI) Then
                cell.Offset(0, i).Font.ColorIndex = getStoneColor()
            End If
            If (i <= rightDown) Then
                cell.Offset(i, i).Font.ColorIndex = getStoneColor()
            End If
            If (i <= downI) Then
                cell.Offset(i, 0).Font.ColorIndex = getStoneColor()
            End If
            If (i <= leftDown) Then
                cell.Offset(i, -i).Font.ColorIndex = getStoneColor()
            End If
            If (i <= leftI) Then
                cell.Offset(0, -i).Font.ColorIndex = getStoneColor()
            End If

            Call countStone
            Sleep 150
        Next i
    End If
End Function

Private Function canPutStone(ByVal cell As Range) As Boolean
    Dim color As Integer
    Dim i As Integer

    ' Left Up
    For i = 1 To 7
        If (cell.Row - i < 1 Or cell.Column - i < 1) Then
            Exit For
        End If

        If (cell.Offset(-i, -i).Value = "") Then
            Exit For
        End If

        If (cell.Offset(-i, -i).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            canPutStone = True
            Exit Function
        End If
    Next i

    ' Up
    For i = 1 To 7
        If (cell.Row - i < 1) Then
            Exit For
        End If

        If (cell.Offset(-i, 0).Value = "") Then
            Exit For
        End If

        If (cell.Offset(-i, 0).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            canPutStone = True
            Exit Function
        End If
    Next i

    ' Right Up
    For i = 1 To 7
        If (cell.Row - i < 1 Or cell.Column + i > 8) Then
            Exit For
        End If

        If (cell.Offset(-i, i).Value = "") Then
            Exit For
        End If

        If (cell.Offset(-i, i).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            canPutStone = True
            Exit Function
        End If
    Next i

    ' Right
    For i = 1 To 7
        If (cell.Column + i > 8) Then
            Exit For
        End If

        If (cell.Offset(0, i).Value = "") Then
            Exit For
        End If

        If (cell.Offset(0, i).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            canPutStone = True
            Exit Function
        End If
    Next i

    ' Right Down
    For i = 1 To 7
        If (cell.Row + i > 8 Or cell.Column + i > 8) Then
            Exit For
        End If

        If (cell.Offset(i, i).Value = "") Then
            Exit For
        End If

        If (cell.Offset(i, i).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            canPutStone = True
            Exit Function
        End If
    Next i

    ' Down
    For i = 1 To 7
        If (cell.Row + i > 8) Then
            Exit For
        End If

        If (cell.Offset(i, 0).Value = "") Then
            Exit For
        End If

        If (cell.Offset(i, 0).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            canPutStone = True
            Exit Function
        End If
    Next i

    ' Left Down
    For i = 1 To 7
        If (cell.Row + i > 8 Or cell.Column - i < 1) Then
            Exit For
        End If

        If (cell.Offset(i, -i).Value = "") Then
            Exit For
        End If

        If (cell.Offset(i, -i).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            canPutStone = True
            Exit Function
        End If
    Next i

    ' Left
    For i = 1 To 7
        If (cell.Column - i < 1) Then
            Exit For
        End If

        If (cell.Offset(0, -i).Value = "") Then
            Exit For
        End If

        If (cell.Offset(0, -i).Font.ColorIndex = getStoneColor()) Then
            If (i = 1) Then
                Exit For
            End If

            canPutStone = True
            Exit Function
        End If
    Next i

    canPutStone = False
End Function

Private Sub changeTurn()
    isBlackTurn = Not isBlackTurn
    Range("L3").Font.ColorIndex = getStoneColor()

    Dim isFull As Boolean: isFull = True
    Dim cell As Range
    For Each cell In Range("A1:H8")
        If (cell.Value = "") Then
            isFull = False
        End If
    Next cell

    If (isFull) Then
        Call Finish
        Exit Sub
    End If

    Dim placeable() As Range
    placeable = getPlaceable()

    On Error Resume Next
    If (Not isnumeric(UBound(placeable))) Then
        If (isSkippedBefore) Then
            Call Finish
            Exit Sub
        End If

        MsgBox "パス", , "オセロシステム"
        isSkippedBefore = True
        Call changeTurn
    End If

    If (vsCPU And isBlackTurn And firstPlayer = 1) Then
        Call RunAI
    ElseIf (vsCPU And Not isBlackTurn And firstPlayer = 0) Then
        Call RunAI
    Else
        Application.OnKey "~", "putStone"
    End If

    isSkippedBefore = False
End Sub

Private Sub countStone()
    With Range("K5:L6")
        .Font.ColorIndex = 1
        .Font.Size = 36
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Range("K5:K6").Interior.ColorIndex = 10
    Range("K5:K6").Value = "●"
    Range("K6").Font.ColorIndex = 2

    Range("L5:L6").Value = 0

    Dim cell As Range
    For Each cell In Range("A1:H8")
        If (cell.Value <> "") Then
            If (cell.Font.ColorIndex = 1) Then
                Range("L5").Value = Range("L5").Value + 1
            Else
                Range("L6").Value = Range("L6").Value + 1
            End If
        End If
    Next cell
End Sub

Private Function getPlaceable() As Range()
    Dim arrayLength As Integer: arrayLength = 0
    Dim cellList() As Range
    Dim cell As Range

    For Each cell In Range("A1:H8")
        If (cell.Value = "") Then
            If (canPutStone(cell)) Then
                ReDim Preserve cellList(arrayLength)
                Set cellList(arrayLength) = cell
                arrayLength = arrayLength + 1
            End If
        End If
    Next cell

    getPlaceable = cellList
End Function

Private Function getStoneColor() As Integer
    If (isBlackTurn) Then
        getStoneColor = 1
    Else
        getStoneColor = 2
    End If
End Function

' Not really an AI lol
Private Sub RunAI()
    Application.OnKey "~", ""

    Dim placeable() As Range
    placeable = getPlaceable()

    Dim candidate As Integer
    Dim candidatePriority As Integer: candidatePriority = 0
    Dim base As Integer
    Dim i As Integer
    Dim calc As Integer

    For i = 0 To UBound(placeable)
        base = 0

        ' Around the corners
        ' However, when corners are already taken,
        '  around the corners are not so important anymore
        If Not (Intersect(Range("A1:B2"), Cells(placeable(i).Row, placeable(i).Column)) Is Nothing) Then
            If (Range("A1").Value = "") Then
                base = -30
            End If
        ElseIf Not (Intersect(Range("G1:H2"), Cells(placeable(i).Row, placeable(i).Column)) Is Nothing) Then
            If (Range("A1").Value = "") Then
                base = -30
            End If
        ElseIf Not (Intersect(Range("A7:B8"), Cells(placeable(i).Row, placeable(i).Column)) Is Nothing) Then
            If (Range("A1").Value = "") Then
                base = -30
            End If
        ElseIf Not (Intersect(Range("G7:H8"), Cells(placeable(i).Row, placeable(i).Column)) Is Nothing) Then
            If (Range("A1").Value = "") Then
                base = -30
            End If
        End If

        ' Corners
        If (placeable(i).Row = 1 And placeable(i).Column = 1) Then
            base = 100
        ElseIf (placeable(i).Row = 8 And placeable(i).Column = 1) Then
            base = 100
        ElseIf (placeable(i).Row = 1 And placeable(i).Column = 8) Then
            base = 100
        ElseIf (placeable(i).Row = 8 And placeable(i).Column = 8) Then
            base = 100
        End If

        calc = base + placeStone(Cells(placeable(i).Row, placeable(i).Column), True)
        If (candidatePriority < calc) Then
            candidate = i
            candidatePriority = calc
        End If
    Next i

    Sleep 2000

    Cells(placeable(candidate).Row, placeable(candidate).Column).Select
    Call putStone
End Sub

Private Sub Finish()
    Dim winner As String

    If (Range("L5").Value = Range("L6").Value) Then
        MsgBox "引き分け", , "ゲーム終了"
    Else
        If (Range("L5").Value > Range("L6").Value) Then
            winner = "黒"
        Else
            winner = "白"
        End If

        MsgBox winner + "の勝利", , "ゲーム終了"
    End If

    Application.OnKey "~", ""
    MainMenu.Show
End Sub

' For debug purpose
Private Sub reset()
    Call StartOthello(1, True)
End Sub
