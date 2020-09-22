Attribute VB_Name = "ModHigh"
Option Explicit
Public scores(9) As score
Public Loaded As Boolean       'True if the scores have already been loaded

Public Type score
    n As String  'Name
    s As Integer 'Score
End Type

'Sub: LoadScores
'Purpose: loads scores from a file and records stores them in the scores array
Public Sub LoadScores()
    'I use random access for scores because numbers as stored as binary not Ascii
    '     it should discourage people from tampering with the file.
    Loaded = True
    Dim i As Integer
    Open App.Path & "\scores.dat" For Random As #1
        For i = 0 To 9
            Get #1, , scores(i).n
            Get #1, , scores(i).s
        Next i
    Close #1
    Exit Sub
End Sub

'Sub: SaveScores
'Purpose: Writes scores from the scores array to a file using random access
Public Sub SaveScores()
    Dim i As Integer
    Open App.Path & "\scores.dat" For Random As #1
        For i = 0 To 9
            Put #1, , scores(i).n         'Write the name
            Put #1, , scores(i).s         'Write the score
        Next i
    Close #1
End Sub

'Sub: Addscore
'Purpose: Adds a score ,in the correct place, to the scores array
'Parameters: s;   The score

Public Sub AddScore(s As Integer)
    Dim n As String
    Dim a As Integer, i As Integer
    n = InputBox("Enter your name:", "Congratulations, you have a top score!")
    If n <> "" Then
        Do While s < scores(i).s
            i = i + 1
            If i > 10 Then Exit Sub
        Loop

        For a = 8 To i Step -1
            scores(a + 1) = scores(a)   'Make room for our new score
        Next a

        scores(i).n = n
        scores(i).s = s
        SaveScores
        Load FrmHigh       'Load the high scores form so the user can se their score
    End If
End Sub
