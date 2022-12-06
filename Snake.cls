VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Dim isGameRunning As Boolean


Private Sub CommandButton1_Click()
    isGameRunning = False
    isGameRunning = True
    Call CreateGameLoop
End Sub

Sub CreateGameLoop()
    Call ResetGame
    Call RespawnFood
    
    t = Timer
    
    Do While isGameRunning
        If Timer - t >= 0.001 Then
            Call ProcessInput
            Call MoveSnake
            t = Timer
            DoEvents
        End If
    Loop
End Sub

Sub ResetGame()
    [headx] = WorksheetFunction.RandBetween(0, 50)
    [heady] = WorksheetFunction.RandBetween(0, 50)
    
    Range("D3:E1000").ClearContents
    
    If [headx] < 20 Then
        [Direction] = 2
    ElseIf [headx] > 40 Then
        [Direction] = 1
    End If
    
    If [heady] < 20 Then
        [Direction] = 3
    ElseIf [heady] > 40 Then
        [Direction] = 4
    End If
    ''dirÀº ÀÏµý.
End Sub

Sub MoveSnake()
    If [headx] > 50 Or [headx] < 0 Or [heady] > 50 Or [heady] < 0 Then
        isGameRunning = False
        ResetGame
    End If

    
    For i = 2 To 21
        If [headx] > Range("H" & i) - 3 And [headx] < Range("H" & i) + 3 And [heady] > Range("I" & i) - 3 And [heady] < Range("I" & i) + 3 Then
            Range("H" & i & ":I" & i) = ""
            Call GrowSnake
        End If
    Next
    
    For i = Range("D1000").End(xlUp).Row To 3 Step -1
        Range("D" & i) = Range("D" & i - 1)
        Range("E" & i) = Range("E" & i - 1)
    Next
    
    If [Direction] = 1 Or [Direction] = 2 Then
        d = IIf([Direction] = 1, -1, 1)
        [headx] = [headx] + 1 * d
    Else
        d = IIf([Direction] = 3, 1, -1)
        [heady] = [heady] + 1 * d
    End If
    
    If Range("D1000").End(xlUp).Row - 1 > [v7] Then: [v7] = Range("D1000").End(xlUp).Row - 1
    
    
    
    If Range("H1000").End(xlUp).Row = 1 Then RespawnFood
    
    For i = 3 To Range("D1000").End(xlUp).Row
        If Range("D" & i) = [headx] And Range("E" & i) = [heady] Then
            isGameRunning = False
            ResetGame
        End If
    Next
End Sub

Sub RespawnFood()
    foodcount = WorksheetFunction.Max(1, Rnd * 20)
    Range("H2:I22").ClearContents
    
    For i = 2 To foodcount + 1
        Range("H" & i) = Int(WorksheetFunction.RandBetween(0, 50))
        Range("I" & i) = Int(WorksheetFunction.RandBetween(0, 50))
    Next
End Sub


Sub GrowSnake()
    lrow = Range("D1000").End(xlUp).Row + 1
    Range("D" & lrow & ":E" & lrow) = Range("D" & lrow - 1 & ":E" & lrow - 1).Value2
End Sub

Sub ProcessInput()
    If Not [Direction] = 2 And GetAsyncKeyState(vbKeyLeft) Then
        [Direction] = 1
    ElseIf Not [Direction] = 1 And GetAsyncKeyState(vbKeyRight) Then
        [Direction] = 2
    ElseIf Not [Direction] = 4 And GetAsyncKeyState(vbKeyUp) Then
        [Direction] = 3
    ElseIf Not [Direction] = 3 And GetAsyncKeyState(vbKeyDown) Then
        [Direction] = 4
    End If
End Sub

Private Sub Worksheet_Activate()
    Application.OnKey "{LEFT}", ""
    Application.OnKey "{RIGHT}", ""
    Application.OnKey "{UP}", ""
    Application.OnKey "{DOWN}", ""
End Sub

Private Sub Worksheet_Deactivate()
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
End Sub


