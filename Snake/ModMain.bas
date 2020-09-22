Attribute VB_Name = "ModMain"
Option Explicit

'BitBlt: Used for drawing images onto the screen and the BackBuffer
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'GetTickCount: Used for Wait sub
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Plays wave files
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public sound() As String        'Stores the location of all the sounds

Public Enum snd
    apple = 0
    die = 1
    LevelComplete = 2
    GameEnd = 3
    GameStart = 4
End Enum

Public Const SND_ASYNC = &H1               'Constants used for sound
Public Const SND_NODEFAULT = &H2

Public Const DLEFT As Integer = -1  'Directions the snake uses
Public Const DRIGHT As Integer = 1
Public Const DDOWN As Integer = 2
Public Const DUP As Integer = -2

Public Const SRCCOPY = &HCC0020     'Constant used for BitBlt

Public CurLevel As Integer
Public Map(30, 30) As Byte          'Stores info about the playing field

Public hSnakes() As Snake
Public hPlayers As Integer          'Total number of human players
Public cSnakes() As Snake
Public cPlayers As Integer          'Total number of computer players
Public IsAnyComputer As Boolean     'True if there is any computer players
Public AppleCount As Integer        'Records hw many apples are on the map
Public Running As Boolean           'This is false if the game isn't running

Public GameSpeed As Integer         'The higher this is the slower the game goes
Public Paused As Boolean
Public SoundEnabled As Boolean
Public lTimer As Double             'Timer for the level


'Sub: NewGame
'Purpose: Starts a new game

Public Sub NewGame()
    Paused = False 'Make sure the game is unpaused before we start
    lTimer = 0
    LoadMap "map1.dat" 'Load the first map
    CurLevel = 1       'Set the current level to 1

    ReDim hSnakes(hPlayers - 1)
    Set hSnakes(0) = New Snake
        
    If hPlayers = 2 Then
        Set hSnakes(1) = New Snake
        hSnakes(0).init 10, 10, 28, DUP 'Initiate the snakeS
        hSnakes(1).init 10, 20, 28, DUP
        FrmMain.ScaleHeight = 300
    Else
        hSnakes(0).init 10, 15, 28, DUP
        FrmMain.ScaleHeight = 344
    End If

    If cPlayers > 0 Then
        Dim i As Integer
        IsAnyComputer = True
    
        ReDim cSnakes(cPlayers - 1)
        For i = 0 To cPlayers - 1
            Set cSnakes(i) = New Snake
            cSnakes(i).init 10, Int(Rnd * 27) + 2, 1, DDOWN
            cSnakes(i).IsComputer = True
        Next i
    Else
        IsAnyComputer = False
    End If

    AppleCount = 0          'reset the number of apples to zero.
    AddApples 1, 10         'Add a few apples
    AddApples 2, 5          '     to the map.
    FrmMain.LblApple.Caption = "Apples Left : " & AppleCount  'Set the apple label to the current number of apples.
    UpdateLives
    FrmMain.PicStartup.Visible = False
    MoveAll
    DrawAll
    Wait 1000, False
    PlaySound GameStart
    FrmMain.RenderLoop       'Finally start the loop
End Sub

Public Sub AbortGame()
    Running = False
    FrmMain.PicStartup.Visible = True
End Sub

Public Sub GameOver()
    Running = False
    PlaySound GameEnd
    With FrmMain
        .Label1.Caption = "Game Over"
        .Label1.Visible = True
        Wait 2000
        .Label1.Visible = False
        .PicStartup.Visible = True
        .MenuItemNewGame.Enabled = True
        .MenuItemAbort.Enabled = False
    End With
    If UBound(hSnakes) = 0 Then      'Only do high scores if the game is one player
        LoadScores
        If hSnakes(0).score > scores(9).s Then   'If the snakes score is greater than the lowest score then
            AddScore hSnakes(0).score       'We need a new entry in the high scores
        End If
    End If
End Sub

'Sub: NextLevel
'purpose: Sets every thing up for the next level(If there is one)
Public Sub nextlevel()
    Dim i As Integer
    MoveAll
    DrawAll
    FrmMain.Label1.Caption = "Level " & CurLevel & " Completed!"
    FrmMain.Label1.Visible = True
    Wait 2000
    FrmMain.Label1.Visible = False
    DoEvents
    DrawAll
    CurLevel = CurLevel + 1       'Make the current level one higher
    If dir(App.Path & "\map" & CurLevel & ".dat") <> "" Then 'Check to see if the map exists

        LoadMap "Map" & CurLevel & ".dat"   'If it does then load it
        If UBound(hSnakes) = 0 Then
            hSnakes(0).Place 10, 15, 28, DUP
        Else
            hSnakes(0).Place 10, 10, 28, DUP
            hSnakes(1).Place 10, 20, 28, DUP
        End If
        
        If IsAnyComputer Then
            For i = 0 To UBound(cSnakes)
                cSnakes(i).Place 10, Int(Rnd * 27) + 1, 1, DDOWN
            Next i
        End If
        AddApples 1, 10
        AddApples 2, 5

        lTimer = 0

        MoveAll
        DrawAll
        PlaySound LevelComplete
        Wait 1000, False
    Else
        'If the map doesn't exist we'll end the game

        'Let the user know they've won
        MsgBox "All levels completed!", vbExclamation, "Congratulations!"
        GameOver
    End If
End Sub

'Sub RestartLevel
'Purpose: restarts the current level

Public Sub RestartLevel()
    RemoveApples     'Remove the apples off the screen
    lTimer = 0
    AddApples 1, 10     'Put ten green apples on the screen
    AddApples 2, 5      'And five red ones
    Dim i As Integer
    
    If UBound(hSnakes) = 0 Then
        hSnakes(0).Place 10, 15, 28, DUP
    Else
        hSnakes(0).Place 10, 10, 28, DUP
        hSnakes(1).Place 10, 20, 28, DUP
    End If
        
    If IsAnyComputer Then
        For i = 0 To UBound(cSnakes)
            cSnakes(i).Place 10, Int(Rnd * 27) + 1, 1, DUP ' Restart the snakes
        Next i
    End If

    MoveAll
    DrawAll             'Draw everthing then
    Wait 1000, False          'Wait a second to give the user a little time
End Sub
'Sub: LoadMap
'Purpose: Loads a map from a file and saves it in the Map array
'Parameters: FileName;    String indicating file location must be under the apps root directory
Public Sub LoadMap(FileName As String)
    Dim tmp As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    Open App.Path & "\" & FileName For Input As #1
        For i = 0 To 29 'Goes through all te rows in the file
            Line Input #1, tmp 'Input a row from the file into the temporary string
            For j = 0 To 29 'Go through each column
                If Mid(tmp, j + 1, 1) = "X" Then 'There is a block on this square
                    Map(j, i) = 1 'set it as a block
                Else
                    Map(j, i) = 0 'Else there is nothing so set it as a blank space
                End If
            Next j
        Next i
    Close #1
End Sub


'Sub: AddApples
'Purpose: Adds different types of apples to the map
'Parameters: t;    The type of Apple, 1 for green, 2 for red
'            Num;  The Number of apples to add
Public Sub AddApples(t As Integer, Num As Integer)
    Dim added As Integer
    Dim TimeOut As Integer
    Dim x As Integer, y As Integer
    If t = 1 Or t = 2 Then
        Do While added < Num Or TimeOut > 30 ' If it can't find a place to put an apple in 30 loops then give in.
            x = Int(Rnd * 30)
            y = Int(Rnd * 30)
            TimeOut = TimeOut + 1
            If FreeSquare(x, y) Then
                TimeOut = 0 'Found a place to put an apple so set timeout to zero again.
                added = added + 1
                Map(x, y) = t + 1 ' Add the apples to the map.
                AppleCount = AppleCount + 1
            End If
        Loop
    End If
    UpdateApples
End Sub
'Sub: RemoveApples
'Purpose: To clear the map of all apples
Public Sub RemoveApples()
    Dim i As Integer, j As Integer
    For i = 0 To 29
        For j = 0 To 29
            If Map(i, j) > 1 Then Map(i, j) = 0
        Next j
    Next i
    AppleCount = 0
End Sub

Public Sub UpdateApples()
    FrmMain.LblApple.Caption = "Apples Left : " & AppleCount
    If AppleCount <= 0 Then nextlevel
End Sub

Public Sub UpdateScore()
    FrmMain.LblScore1 = "Score:" & hSnakes(0).score
    If UBound(hSnakes) = 1 Then FrmMain.LblScore2 = "Score:" & hSnakes(1).score
End Sub

Public Sub UpdateLives()
    FrmMain.LblLives1 = "Lives : " & hSnakes(0).Lives
    DoEvents
    If UBound(hSnakes) = 1 Then FrmMain.LblLives2 = "Lives : " & hSnakes(1).Lives
End Sub

Public Sub UpdateTimer()
    lTimer = lTimer + 0.4       'Increment the timer.

    If lTimer >= 100 Then       'If the timers greater than 100
        AddApples 1, 10         'Add some more apples
        AddApples 2, 5
        lTimer = 0              'And set it back to zero
    End If
End Sub
'Function: Free Square
'Purpose: Finds in a certain square on the map has anything on it
'Returns: True if the square is free
'Paramaters: x     ;Co-ordanates
'            y     ;      "
Public Function FreeSquare(x As Integer, y As Integer, Optional CheckApples As Boolean = True) As Boolean
    FreeSquare = True

    If CheckApples Then
        If Map(x, y) <> 0 Then   'If there is something on this square then
            FreeSquare = False   'This isn't a free square
            Exit Function        'and exit the function to save time
        End If
    Else
        If Map(x, y) = 1 Then   'If there is something on this square then
            FreeSquare = False   'This isn't a free square
            Exit Function        'and exit the function to save time
        End If
    End If

    Dim i As Integer
    
    For i = 0 To UBound(hSnakes)
        If hSnakes(i).UsesSquare(x, y) Then  'If a human snake uses this square then
            FreeSquare = False               'It also isn't free
            Exit Function
        End If
    Next i
    
    'Check for computer players last because it is least likely to be nessasary
    If IsAnyComputer Then
        For i = 0 To UBound(cSnakes)
            If cSnakes(i).UsesSquare(x, y) Then
                FreeSquare = False
                Exit Function
            End If
        Next i
    End If
End Function

'Sub: Wait
'Purpose: Pause the game for a small interval
'Parameters: w;      Amount of Miliseconds to Wait
Public Sub Wait(w As Integer, Optional DEvents As Boolean = True)
    Dim s As Long
    s = GetTickCount + w
    Do While GetTickCount < s
        If DEvents Then DoEvents  'Keeps every thing else running smoothy
    Loop
End Sub

'Sub: DrawAll
'Purpose: Draws everything onto the screen
Public Sub DrawAll()
    FrmMain.BackBuffer.Cls  'Clear the BackBuffer.
    FrmMain.DrawMap         'Draw the map.

    'Draw player one's snake.
    hSnakes(0).Draw FrmMain.BackBuffer.hDC, FrmMain.Picture1.hDC
    
    'If there is a player two then draw their snake.
    If UBound(hSnakes) = 1 Then hSnakes(1).Draw FrmMain.BackBuffer.hDC, FrmMain.Picture1.hDC
    
    If IsAnyComputer Then 'Draw all the computer snakes if any
        Dim i As Integer
        For i = 0 To UBound(cSnakes)
            cSnakes(i).Draw FrmMain.BackBuffer.hDC, FrmMain.Picture1.hDC
        Next i
    End If
    
    'Draw the Timer bar
    Dim y As Integer, c As Long
    y = Int(lTimer * 3)
    c = RGB(lTimer / 100 * 255, (100 - lTimer) / 100 * 255, 0)
    'Select Case lTimer   'Select the appropriate colour for the timer bar
    '    Case 0 To 50: c = RGB(0, 255, 0)    'Green
    '    Case 50 To 80: c = RGB(255, 128, 0) 'Orange
    '    Case 81 To 100: c = RGB(255, 0, 0)  'Red
    'End Select

    FrmMain.BackBuffer.Line (300, y)-(330, 300), c, BF
    
    'Copy the BackBuffer to the form
    BitBlt FrmMain.hDC, 0, 0, 330, 300, FrmMain.BackBuffer.hDC, 0, 0, SRCCOPY
End Sub

Public Sub MoveAll()
    Dim i As Integer
    For i = 0 To UBound(hSnakes)
        hSnakes(i).Move
    Next i

    If IsAnyComputer Then
        For i = 0 To UBound(cSnakes)
            cSnakes(i).Move
        Next i
    End If
End Sub

Public Sub PlaySound(s As snd)
    Dim d As String
    d = App.Path & "\sounds\"
    Select Case s
        Case 0
            d = d & "apple.wav"
        Case 1
            d = d & "die.wav"
        Case 2
            d = d & "level.wav"
        Case 3
            d = d & "gameover.wav"
        Case 4
            d = d & "newgame.wav"
    End Select
    
    If dir(d) <> "" Then
        sndPlaySound d, SND_ASYNC Or SND_NODEFAULT
    End If
End Sub
