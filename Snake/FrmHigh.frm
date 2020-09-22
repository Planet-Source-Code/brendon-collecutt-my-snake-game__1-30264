VERSION 5.00
Begin VB.Form FrmHigh 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hall of Fame!"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4575
   Icon            =   "FrmHigh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ok"
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox TxtScore 
      Height          =   285
      Index           =   9
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox TxtScore 
      Height          =   285
      Index           =   8
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox TxtScore 
      Height          =   285
      Index           =   7
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox TxtScore 
      Height          =   285
      Index           =   6
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox TxtScore 
      Height          =   285
      Index           =   5
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox TxtScore 
      Height          =   285
      Index           =   4
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox TxtScore 
      Height          =   285
      Index           =   3
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox TxtScore 
      Height          =   285
      Index           =   2
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox TxtScore 
      Height          =   285
      Index           =   1
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox TxtScore 
      Height          =   285
      Index           =   0
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox TxtName 
      Height          =   285
      Index           =   9
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3720
      Width           =   3055
   End
   Begin VB.TextBox TxtName 
      Height          =   285
      Index           =   8
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3360
      Width           =   3055
   End
   Begin VB.TextBox TxtName 
      Height          =   285
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3000
      Width           =   3055
   End
   Begin VB.TextBox TxtName 
      Height          =   285
      Index           =   6
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2640
      Width           =   3055
   End
   Begin VB.TextBox TxtName 
      Height          =   285
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Width           =   3055
   End
   Begin VB.TextBox TxtName 
      Height          =   285
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   3055
   End
   Begin VB.TextBox TxtName 
      Height          =   285
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   3055
   End
   Begin VB.TextBox TxtName 
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   3055
   End
   Begin VB.TextBox TxtName 
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   3055
   End
   Begin VB.TextBox TxtName 
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmHigh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Loaded = False Then LoadScores
    Dim i As Integer
    For i = 0 To 9
        TxtName(i).Text = scores(i).n
        TxtScore(i).Text = scores(i).s
    Next i
    Show
End Sub
