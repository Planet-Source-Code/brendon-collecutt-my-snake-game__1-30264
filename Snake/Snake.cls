VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Snake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private seg() As segment      'All the snakes body segments

Public dir As Integer         'The current direction the snake is facing
Public Colour As Integer
Public IsComputer As Boolean  'True if this snake is controled by the computer
Public Lives As Integer       'Lives the player,If human, has left
Public score As Integer

Private Type segment
    x As Integer
    y As Integer
End Type

Public Sub init(length As Integer, x As Integer, y As Integer, direction As Integer, Optional Colour)
    Randomize Timer
    Dim i As Integer
    Lives = 5
    If VarType(Colour) = vbError Then 'Vartype(varName) will return vbError if it doesn't have a type
        Me.Colour = Int(Rnd * 5)      'Since the user hasn't given this snake a colour then give it a random one
    Else
        Me.Colour = Colour
    End If
    ReDim seg(length - 1)
    For i = 0 To length - 1
        seg(i).x = x
        seg(i).y = y
    Next i
    dir = direction
End Sub

'Sub: Place
'Purpose: Place the snake somewhere on the map
'Parameters: length;   The new length of the snake
'            x;     Co-ordante
'            y;         ''
'            direction;       the direction in which the snake moves

Public Sub Place(length As Integer, x As Integer, y As Integer, direction As Integer)
    Dim i As Integer
    
    ReDim seg(length - 1)   ' Re - dimension the array

    For i = 0 To length - 1 'Go through all the segments
        seg(i).x = x        'Putting them all in one place
        seg(i).y = y
    Next i
    dir = direction  'Set the direction to that of the one specified
End Sub

'Sub: ChangeDirection
'Purpose: Used with key events to check whether the newDirection is valid
'Parameters: NewDirection;   The direction the user wishes to move the snake
Public Sub ChangeDirection(NewDirection As Integer)

    Select Case NewDirection 'The snake can't move backwards,
        Case DLEFT           '     this code checks to make sure
            If seg(1).x <> seg(0).x - 1 Then   'that isn't happening
                dir = NewDirection
            End If
        Case DRIGHT
            If seg(1).x <> seg(0).x + 1 Then
                dir = NewDirection
            End If
        Case DDOWN
            If seg(1).y <> seg(0).y + 1 Then
                dir = NewDirection
            End If
        Case DUP
            If seg(1).y <> seg(0).y - 1 Then
                dir = NewDirection
            End If
    End Select
End Sub

'Sub: Move
'Purpose: Moves the snake forward in it current direction doing all collision checking nessesary
Public Sub Move()
    Dim tmp1 As segment
    Dim tmp2 As segment
    Dim i As Integer
    
    If IsComputer Then
        If Int(Rnd * 6) = 0 Then ChangeDirection (Int(Rnd * 5) - 2) 'A one in six chance the computer will pick a new direction
        For i = 0 To 5 ' the higher this number is, The less likely it is to crash
            If CollisionDetected Then 'If the snake is about to crash find a new direction
                ChangeDirection (Int(Rnd * 5) - 2)
            Else
                Exit For
            End If
        Next i
    End If


    If Not CollisionDetected Then
        tmp1 = seg(0)
        '<<< move the head forward
        Select Case dir
            Case DLEFT
                seg(0).x = seg(0).x - 1
            Case DRIGHT
                seg(0).x = seg(0).x + 1
            Case DDOWN
                seg(0).y = seg(0).y + 1
            Case DUP
                seg(0).y = seg(0).y - 1
        End Select
        '>>>
        
        'step up all other segments
        For i = 1 To UBound(seg)
            tmp2 = seg(i)
            seg(i) = tmp1
            tmp1 = tmp2
        Next i
    Else 'A collision was dectected so we can't move forward

        If IsComputer Or UBound(hSnakes) = 1 Then 'If the snake is computer controled or the game is two player
            If UBound(seg) > 1 Then 'If the snake is longer than two segments
                grow (-1) 'Take a piece off the tail end of the snake
            End If
        Else
            Lives = Lives - 1
            UpdateLives

            If Lives = 0 Then
                GameOver
            Else
                PlaySound die
                RestartLevel
            End If
        End If
    End If
End Sub


'Function: CollisionDetected
'Purpose: Find out if the snake will crash into anything
'         Also handles if the snake runs into an apple
'Returns: True if it will crash, False otherwise

Private Function CollisionDetected() As Boolean
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer

    x = seg(0).x + IIf(Abs(dir) = 1, Sgn(dir), 0)
    y = seg(0).y + IIf(Abs(dir) = 2, Sgn(dir), 0)

    If x = -1 Or y = -1 Or x = 30 Or y = 30 Then
        CollisionDetected = True
        Exit Function
    End If
    Select Case Map(x, y)
        Case 1
            CollisionDetected = True
            Exit Function
        Case 2
            PlaySound apple
            grow 2
            Map(x, y) = 0  'Get rid of the apple
            AppleCount = AppleCount - 1 'one less apple on the map
            UpdateApples
            score = score + 1
            UpdateScore
        Case 3
            PlaySound apple
            grow 6
            Map(x, y) = 0
            AppleCount = AppleCount - 1
            score = score + 5
            UpdateApples
            UpdateScore
        End Select
    If FreeSquare(x, y, False) = False Then
        CollisionDetected = True
    End If

End Function

'Sub: Draw
'Purpose: Draws the snake onto any given DC
'Parameters: DestDc;   The destination for drawing
'            SrcDc;    The source for the snakes pics
'            Colour;   Number indicating with coloured snake to use currently 0 to 4
Public Sub Draw(DestDc As Long, SrcDc As Long, Optional Colour)

    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
    Dim c As Integer
    Dim imgNum As Integer

    
    If VarType(Colour) = vbError Then Colour = Me.Colour 'No colour was specified so use the snakes default colour

    '<<<<<<<<< This part draws the body
    For i = 1 To UBound(seg) - 1

        x = seg(i - 1).x - seg(i + 1).x + 2
        y = seg(i - 1).y - seg(i + 1).y + 2
        c = x + y * 5
        imgNum = 0
        Select Case c
            Case 22, 2: imgNum = 4
            Case 10, 14: imgNum = 5
            Case 16, 8:
                If seg(i).x < seg(i - 1).x Or seg(i).x < seg(i + 1).x Then
                    imgNum = 6
                Else
                    imgNum = 9
                End If
            Case 6, 18:
                If seg(i).x > seg(i - 1).x Or seg(i).x > seg(i + 1).x Then
                    imgNum = 7
                Else
                    imgNum = 8
                End If
        End Select
            
        BitBlt DestDc, seg(i).x * 10, seg(i).y * 10, 10, 10, SrcDc, imgNum * 10, Colour * 10, SRCCOPY
    Next i
    '>>>>>>>>>
    
    ' <<<<<<<< This section here draws the head
    If seg(0).x - seg(1).x = -1 Then
        imgNum = 10
    ElseIf seg(0).x - seg(1).x = 1 Then
        imgNum = 12
    Else
        If seg(0).y - seg(1).y = -1 Then
            imgNum = 11
        Else
            imgNum = 13
        End If
    End If
    
    BitBlt DestDc, seg(0).x * 10, seg(0).y * 10, 10, 10, SrcDc, imgNum * 10, Colour * 10, SRCCOPY
    '>>>>>>>>>
    
    '<<<<<<<<< This draws the tail
    Dim NextSeg As Integer
    
    c = UBound(seg)
    For i = c - 1 To 0 Step -1
        If seg(c).x <> seg(i).x Or seg(c).y <> seg(i).y Then NextSeg = i: Exit For
    Next i

    If seg(c).x - seg(NextSeg).x < 0 Then
        imgNum = 0
    ElseIf seg(c).x - seg(NextSeg).x > 0 Then
        imgNum = 2
    Else
        If seg(c).y - seg(NextSeg).y = -1 Then
            imgNum = 1
        ElseIf seg(c).y - seg(NextSeg).y = 1 Then
            imgNum = 3
        End If
    End If

    BitBlt DestDc, seg(c).x * 10, seg(c).y * 10, 10, 10, SrcDc, imgNum * 10, Colour * 10, SRCCOPY
    '>>>>>>>>>
End Sub
    
'Sub:   Grow
'Purpose: Increase the snakes length by a given amount
'         Negative amounts will make the snake shrink
'Parameters: Length;    How much to make the snake grow

Public Sub grow(length As Integer)
    Dim l As Integer
    Dim i As Integer
    l = UBound(seg)
    ReDim Preserve seg(l + length)
    For i = 1 To length
        seg(l + i).x = seg(l).x 'Set all the new segments equal
        seg(l + i).y = seg(l).y '   to the tail segment.
    Next i
End Sub


'Function:   UsesSquare
'Purpose:    To find if this snake occupies the given Co-ordinate
'Parameters: x;      X Co-ordinate
'            y:      Y Co-ordinate
'Returns:    True if the snake uses this square
Public Function UsesSquare(x As Integer, y As Integer) As Boolean
    Dim i As Integer

    For i = 0 To UBound(seg) 'Cycle through all the snakes segments
        If seg(i).x = x And seg(i).y = y Then ' if there co-ordinates match then
            UsesSquare = True 'the snake does use this square
            Exit Function     'No point check any more so exit the function
        End If
    Next i
End Function
