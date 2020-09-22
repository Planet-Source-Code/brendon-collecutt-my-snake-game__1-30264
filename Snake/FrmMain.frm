VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snake Game - By Brendon Collecutt"
   ClientHeight    =   5160
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   4950
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   330
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   2160
      Picture         =   "FrmMain.frx":12FA
      ScaleHeight     =   150
      ScaleWidth      =   1920
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   2160
      Picture         =   "FrmMain.frx":223C
      ScaleHeight     =   750
      ScaleWidth      =   2130
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.PictureBox BackBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   -2880
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   2
      Top             =   -2760
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.PictureBox PicStartup 
      AutoSize        =   -1  'True
      Height          =   4560
      Left            =   0
      Picture         =   "FrmMain.frx":7616
      ScaleHeight     =   4500
      ScaleWidth      =   4500
      TabIndex        =   11
      Top             =   0
      Width           =   4560
   End
   Begin VB.Label LblPlayer2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Player Two"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label LblPlayer1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Player One"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label LblLives2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lives : 5"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label LblLives1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lives : 5"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label LblScore2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Score : 0"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Level One Completed!"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label LblScore1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Score : 0"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label LblApple 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Apples Left :"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Menu MenuGame 
      Caption         =   "Game"
      Begin VB.Menu MenuItemNewGame 
         Caption         =   "New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenuItemPause 
         Caption         =   "Pause Game"
         Shortcut        =   {F3}
      End
      Begin VB.Menu MenuItemAbort 
         Caption         =   "Abort Game"
         Shortcut        =   {F4}
      End
      Begin VB.Menu hRule1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuItemHigh 
         Caption         =   "High Scores"
      End
      Begin VB.Menu hRule3 
         Caption         =   "-"
      End
      Begin VB.Menu MenuItemExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MenuOptions 
      Caption         =   "Options"
      Begin VB.Menu MenuItemOnePlayer 
         Caption         =   "One Player"
         Checked         =   -1  'True
      End
      Begin VB.Menu MenuItemTwoPlayer 
         Caption         =   "Two Player"
      End
      Begin VB.Menu hRule2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSubComPlayers 
         Caption         =   "Computer Players"
         Begin VB.Menu mnipc 
            Caption         =   "0"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnipc 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu mnipc 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu mnipc 
            Caption         =   "3"
            Index           =   3
         End
         Begin VB.Menu mnipc 
            Caption         =   "4"
            Index           =   4
         End
      End
      Begin VB.Menu MenuSpeed 
         Caption         =   "Speed"
         Begin VB.Menu MenuItemSpeed 
            Caption         =   "Slowest"
            Index           =   0
         End
         Begin VB.Menu MenuItemSpeed 
            Caption         =   "Slow"
            Index           =   1
         End
         Begin VB.Menu MenuItemSpeed 
            Caption         =   "Medium"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MenuItemSpeed 
            Caption         =   "Fast"
            Index           =   3
         End
         Begin VB.Menu MenuItemSpeed 
            Caption         =   "Fastest"
            Index           =   4
         End
      End
      Begin VB.Menu hRule4 
         Caption         =   "-"
      End
      Begin VB.Menu MenuItemSound 
         Caption         =   "Sound"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not Paused And Running Then
        Select Case KeyCode
            Case vbKeyLeft
                hSnakes(0).ChangeDirection DLEFT
            Case vbKeyRight
                hSnakes(0).ChangeDirection DRIGHT
            Case vbKeyDown
                hSnakes(0).ChangeDirection DDOWN
            Case vbKeyUp
                hSnakes(0).ChangeDirection DUP
        
            Case vbKeyA
                If UBound(hSnakes) = 1 Then hSnakes(1).ChangeDirection DLEFT

            Case vbKeyD
                If UBound(hSnakes) = 1 Then hSnakes(1).ChangeDirection DRIGHT
            Case vbKeyS
                If UBound(hSnakes) = 1 Then hSnakes(1).ChangeDirection DDOWN
            Case vbKeyW
                If UBound(hSnakes) = 1 Then hSnakes(1).ChangeDirection DUP
        End Select
    End If
End Sub

Private Sub Form_Load()
    Randomize Timer
    Show               'Show the form
    hPlayers = 1
    cPlayers = 0
    GameSpeed = 90
    SoundEnabled = True
End Sub

'Sub:   RenderLoop
'Purpose: Main loop, handles drawing etc

Public Sub RenderLoop()
    Dim t As Long 't is the time taken to do everything

    Running = True

    Do While Running
        t = GetTickCount

        If Not Paused Then
            MoveAll           'move all the snakes
            UpdateTimer
        End If

        DrawAll                 'draw everything
        t = GetTickCount - t    'This compensates for the time taken
        If t > 200 Then t = 0   '    to move and draw everything, keeps the game
                                '         running at a steady pace
        Wait GameSpeed - t 'Wait a while just to slow the game down
        '        so it doesn't go to fast
        DoEvents 'And Do events so the form events can be handled
    Loop

    Erase hSnakes()
    Erase cSnakes()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Running = False  'Ends the Main Loop
    End
End Sub

'Sub:   DrawMap
'Purpose:  Draws the map onto the screen
Public Sub DrawMap()
    Dim i As Integer
    Dim j As Integer
    For i = 0 To 29     'Cycle through all
        For j = 0 To 29       'squares on the map
            If Map(i, j) <> 0 Then  'If there is nothing on a square don't draw it
                BitBlt BackBuffer.hDC, i * 10, j * 10, 10, 10, PicMap.hDC, (Map(i, j) - 1) * 10, 0, SRCCOPY
            End If
        Next j
    Next i
End Sub

Private Sub MenuItemAbort_Click()
    AbortGame
    MenuItemAbort.Enabled = False
    MenuItemNewGame.Enabled = True
End Sub

Private Sub MenuItemExit_Click()
    Unload Me
End Sub

Private Sub MenuItemHigh_Click()
    Load FrmHigh
End Sub

Private Sub MenuItemNewGame_Click()
    MenuItemNewGame.Enabled = False
    MenuItemAbort.Enabled = True
    NewGame
End Sub

Private Sub MenuItemOnePlayer_Click()
    If MenuItemOnePlayer.Checked = False Then
        MenuItemOnePlayer.Checked = True
        hPlayers = 1
        MenuItemTwoPlayer.Checked = False
    End If
End Sub

Private Sub MenuItemPause_Click()
    Paused = Not Paused
    Label1.Caption = "Paused"
    Label1.Visible = Paused
    MenuItemPause.Caption = IIf(Paused, "Unpause", "Pause")
End Sub

Private Sub MenuItemSound_Click()
    MenuItemSound.Checked = Not MenuItemSound.Checked
    SoundEnabled = MenuItemSound.Checked
End Sub

Private Sub MenuItemSpeed_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 4
        MenuItemSpeed(i).Checked = False
    Next i
    MenuItemSpeed(Index).Checked = True

    Select Case Index
        Case 0: GameSpeed = 200     'Slowest
        Case 1: GameSpeed = 130     'Slow
        Case 2: GameSpeed = 90      'Medium
        Case 3: GameSpeed = 30      'Fast
        Case 4: GameSpeed = 0       'Fastest(This will vary greatly on different computers
    End Select
End Sub

Private Sub MenuItemTwoPlayer_Click()
    If MenuItemTwoPlayer.Checked = False Then
        MenuItemTwoPlayer.Checked = True
        hPlayers = 2
        MenuItemOnePlayer.Checked = False
    End If
End Sub

Private Sub mnipc_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 4
        mnipc(i).Checked = False
    Next i
    mnipc(Index).Checked = True
    cPlayers = Index
End Sub
