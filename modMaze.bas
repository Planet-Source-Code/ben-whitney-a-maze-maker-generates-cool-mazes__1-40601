Attribute VB_Name = "modMaze"
Option Explicit

Public Const MIN_SIZE = 5     'Define minimum and maximum size for validation
Public Const MAX_SIZE = 100
Public Enum WallData
    NORTH = 1
    SOUTH = 2
    EAST = 4
    WEST = 8
    ALREADY_VISITED = 16
    WALL_READY = 32
    ESCAPE_PATH = 64
End Enum
    

Public Type Maze
    Initialized As Boolean          'whether the maze has been initialized
    MazeData() As WallData          'wall data can be set by boolean OR operator and checked with AND operator
    Rows As Integer                 'number of rows
    Cols As Integer                 'number of columns
    RowWidth As Integer             'width, in pixels, between rows
    ColWidth As Integer             'width, in pixels, between columns
    OutputPictureBox As PictureBox  'somewhat necessary
    NumberOfMoves As Long           'number of moves made
    SolutionMoves As Long           'solution number of moves, since the solution is calculated in advance
    EntranceX As Integer            'entrance point X location (horizontal)
    EntranceY As Integer            'entrance point Y location (vertical)
    ExitX As Integer                'exit point X location (horizontal)
    ExitY As Integer                'exit point Y location (vertical)
    Margin As Integer               'margin Size in pixels
    PlayerX As Integer              'player x position
    PlayerY As Integer              'player y position
    clrWalls As Long                'color of the walls
    clrBackground As Long           'color of the background
    clrVisited As Long              'color of the visited path
    clrCurrent As Long              'color of the current location
    clrSolution As Long             'color of the solution
End Type

Public Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public myMaze As Maze


Public Sub InitializeMaze(Rows As Integer, Cols As Integer, ByRef PicBox As PictureBox, Complexity As Double)
    'The following steps are required to create the maze
    '
    '   1)  Redimension the MazeData array so it is the correct
    '       size to hold the maze.
    '
    '   2)  Set maze base parameters
    '
    '   3)  Enable all walls (NSEW) on all sides
    '
    '   4)  Generate an escape route of adequate complexity, picking the entrance and exit points
    '
    '   5)  Remove walls until all sections of the maze are accessible
    '
    
    Dim x As Integer
    Dim y As Integer
    Dim turn As Integer
    Dim direction As WallData
    Dim lastDirection As WallData
    Dim minComplexity As Integer
    Dim bLeftChecked As Boolean
    Dim bRightChecked As Boolean
    Dim bStraightChecked As Boolean
    Dim bCheckForFinish As Boolean
    Dim bMazeFinished As Boolean
    Dim startSide As WallData
    
    Screen.MousePointer = vbHourglass
    
    '1 ---- REDIMENSION THE MAZE DATA
    ReDim myMaze.MazeData(Cols + 1, Rows + 1)
    
    
    '2 ---- SET MAZE BASE PARAMETERS
    Set myMaze.OutputPictureBox = PicBox        'sets output window data
    
    myMaze.Margin = 25                          'set parameters
    myMaze.Rows = Rows
    myMaze.Cols = Cols
    ResizeMaze True                             'initialize dimensions of maze
    
    myMaze.clrBackground = Val(GetSetting("MazeMaker", "Color", "Background", "0"))
    myMaze.clrCurrent = Val(GetSetting("MyMaze", "Color", "Current", Str(vbWhite)))
    myMaze.clrSolution = Val(GetSetting("MyMaze", "Color", "Solution", Str(RGB(0, 128, 255))))
    myMaze.clrVisited = Val(GetSetting("MyMaze", "Color", "Visited", Str(RGB(0, 0, 128))))
    myMaze.clrWalls = Val(GetSetting("MyMaze", "Color", "Walls", Str(vbWhite)))
    
    ApplyColorScheme
    
    myMaze.Initialized = True
    
    Randomize Timer
    minComplexity = (myMaze.Cols + myMaze.Rows) * Complexity
    

    'frmMaze.Visible = True
Retry:
    '3 ---- ENABLE ALL WALLS ON ALL SIDES
    For x = 1 To myMaze.Cols
        For y = 1 To myMaze.Rows
            myMaze.MazeData(x, y) = (NORTH Or SOUTH Or EAST Or WEST)
        Next y
    Next x
    
    myMaze.SolutionMoves = 0
    
    '4 ---- GENERATE AN ESCAPE ROUTE OF ADEQUATE COMPLEXITY
    'Determine the starting point
    
    If Int(Rnd * 2) = 1 Then
        'start on the left hand side
        myMaze.EntranceX = 1
        myMaze.EntranceY = Int(Rnd * myMaze.Rows) + 1
        lastDirection = EAST
        startSide = WEST
        RemoveWall myMaze.EntranceX, myMaze.EntranceY, WEST
    Else
        'start at the top
        myMaze.EntranceX = Int(Rnd * myMaze.Cols) + 1
        myMaze.EntranceY = 1
        lastDirection = SOUTH
        startSide = NORTH
        RemoveWall myMaze.EntranceX, myMaze.EntranceY, NORTH
    End If
    
    'This will be visited if the maze can not randomly generate
    'an escape route of adequate complexity
    ClearVisitedPath
    ClearEscapePath
    
    SetPlayerPosition myMaze.EntranceX, myMaze.EntranceY
    myMaze.MazeData(myMaze.EntranceX, myMaze.EntranceY) = myMaze.MazeData(myMaze.EntranceX, myMaze.EntranceY) Or ESCAPE_PATH
    
    While Not bMazeFinished
        'turn left, right or straight
        bLeftChecked = False
        bStraightChecked = False
        bRightChecked = False
        
badturn:
        bCheckForFinish = False
        If bLeftChecked And bRightChecked And bStraightChecked Then
            'there is nowhere to go, try again
            'Debug.Print "- Regenerating escape path"
            If Int(Rnd * 10) = 5 Then DoEvents    'allow user to ctrl-break
            GoTo Retry
        End If
        
        turn = Int(Rnd * 3) + 1
        Select Case turn
            Case 1
                'left turn
                If lastDirection = SOUTH Then
                    direction = EAST
                ElseIf lastDirection = EAST Then
                    direction = NORTH
                ElseIf lastDirection = NORTH Then
                    direction = WEST
                Else
                    direction = SOUTH
                End If
                bLeftChecked = True
                
            Case 2
                'straight
                direction = lastDirection
                bStraightChecked = True
    
            Case 3
                'right turn
                If lastDirection = SOUTH Then
                    direction = WEST
                ElseIf direction = WEST Then
                    direction = NORTH
                ElseIf direction = NORTH Then
                    direction = EAST
                Else
                    direction = SOUTH
                End If
                bRightChecked = True
                
        End Select
        
        'check to make sure the turn is valid. it would not be a valid turn
        'under the following circumstances:
        '
        '   1)  the turn would turn it back into the maze escape path
        '   2)  the turn would hit a wall without reaching the
        '       adequate complexity
    
        Select Case direction
            Case NORTH
                If myMaze.PlayerY > 1 Then
                    If myMaze.MazeData(myMaze.PlayerX, myMaze.PlayerY - 1) And ALREADY_VISITED Then
                        'hits the escape path
                        GoTo badturn
                    End If
                    'turn is good
                    RemoveWall myMaze.PlayerX, myMaze.PlayerY, NORTH
                    SetPlayerPosition myMaze.PlayerX, myMaze.PlayerY - 1
                
                Else 'hits the wall - since the ending side can only be south or west its a bad turn
                    GoTo badturn
                End If
    
            Case SOUTH
                If myMaze.PlayerY < myMaze.Rows Then
                    If myMaze.MazeData(myMaze.PlayerX, myMaze.PlayerY + 1) And ALREADY_VISITED Then
                        'hits the escape path
                        GoTo badturn
                    End If
                    'turn is good
                    RemoveWall myMaze.PlayerX, myMaze.PlayerY, SOUTH
                    SetPlayerPosition myMaze.PlayerX, myMaze.PlayerY + 1
                    
                Else 'hits the wall
                    If startSide = NORTH Then
                        bCheckForFinish = True
                    Else
                        GoTo badturn
                    End If
                End If
            
            Case EAST
                If myMaze.PlayerX < myMaze.Cols Then
                    If myMaze.MazeData(myMaze.PlayerX + 1, myMaze.PlayerY) And ALREADY_VISITED Then
                        'hits the escape path
                        GoTo badturn
                    End If
                    'turn is good
                    RemoveWall myMaze.PlayerX, myMaze.PlayerY, EAST
                    SetPlayerPosition myMaze.PlayerX + 1, myMaze.PlayerY
                Else
                    If startSide = WEST Then
                        bCheckForFinish = True
                    Else
                        GoTo badturn
                    End If
                End If
            
            Case WEST
                If myMaze.PlayerX > 1 Then
                    If myMaze.MazeData(myMaze.PlayerX - 1, myMaze.PlayerY) And ALREADY_VISITED Then
                        'hits the escape path
                        GoTo badturn
                    End If
                    'turn is good
                    RemoveWall myMaze.PlayerX, myMaze.PlayerY, WEST
                    SetPlayerPosition myMaze.PlayerX - 1, myMaze.PlayerY
                Else
                    'hits the wall - since the ending side can only be south or west its a bad turn
                    GoTo badturn
                End If
        End Select
        
        
        If bCheckForFinish Then
            'a side wall was hit, check to see if the escape route has
            'reached adequate complexity. If it has, declare the maze finished
            If myMaze.SolutionMoves >= minComplexity Then
                bMazeFinished = True
                RemoveWall myMaze.PlayerX, myMaze.PlayerY, direction
            Else
                GoTo badturn
            End If
        Else
            myMaze.SolutionMoves = myMaze.SolutionMoves + 1
            lastDirection = direction
        End If
        myMaze.MazeData(myMaze.PlayerX, myMaze.PlayerY) = myMaze.MazeData(myMaze.PlayerX, myMaze.PlayerY) Or ESCAPE_PATH
    Wend
            
    'Solution has finished, set the end point and exit point
    myMaze.ExitX = myMaze.PlayerX
    myMaze.ExitY = myMaze.PlayerY
    
    '5 ---- REMOVE AT LEAST ONE WALL FROM EACH REMAINING SECTION OF THE MAZE
    Dim bDontRemove As Boolean
    For x = 1 To myMaze.Cols
        For y = 1 To myMaze.Rows
            bDontRemove = False
            Select Case Int(Rnd * 4) + 1
                Case 1
                    If y > 1 Then
                        direction = NORTH
                    Else
                        bDontRemove = True
                    End If
                    
                Case 2
                    If y < myMaze.Rows Then
                        direction = SOUTH
                    Else
                        bDontRemove = True
                    End If
                    
                Case 3
                    If x < myMaze.Cols Then
                        direction = EAST
                    Else
                        bDontRemove = True
                    End If
                    
                Case 4
                    If x > 1 Then
                        direction = WEST
                    Else
                        bDontRemove = True
                    End If
            End Select
            If Not bDontRemove Then
                If GetWallCount(x, y) > 2 Then
                    RemoveWall x, y, direction
                End If
            End If
        Next y
    Next x
        
    'second pass, clean up some more
    For x = 1 To myMaze.Cols
        For y = 1 To myMaze.Rows
            bDontRemove = False
            If GetWallCount(x, y) = 3 Then
                Select Case Int(Rnd * 4) + 1
                    Case 1
                        If y > 1 Then
                            direction = NORTH
                        Else
                            bDontRemove = True
                        End If
                        
                    Case 2
                        If y < myMaze.Rows Then
                            direction = SOUTH
                        Else
                            bDontRemove = True
                        End If
                        
                    Case 3
                        If x < myMaze.Cols Then
                            direction = EAST
                        Else
                            bDontRemove = True
                        End If
                        
                    Case 4
                        If x > 1 Then
                            direction = WEST
                        Else
                            bDontRemove = True
                        End If
                End Select
                If Not bDontRemove Then
                    If GetWallCount(x, y) > 2 Then
                        RemoveWall x, y, direction
                    End If
                End If
            End If
        Next y
    Next x
                    
    ClearVisitedPath
    CloseShortcuts
    
    'open up the exit
    RemoveWall myMaze.ExitX, myMaze.ExitY, direction
    ClearVisitedPath
    
    Screen.MousePointer = vbArrow
    DrawMaze
End Sub
Public Sub ResizeMaze(Optional bNoDraw As Boolean)
    'calculate dimensions of the maze
    myMaze.RowWidth = (myMaze.OutputPictureBox.ScaleHeight - (2 * myMaze.Margin)) / myMaze.Rows
    myMaze.ColWidth = (myMaze.OutputPictureBox.ScaleWidth - (2 * myMaze.Margin)) / myMaze.Cols
    
    If Not bNoDraw Then
        'draw
        DrawMaze
    End If
    
End Sub
Private Sub SetPlayerPosition(x As Integer, y As Integer)
    'sets the new player position
    If x = 0 And y = 0 Then
        Exit Sub
    End If
    myMaze.PlayerX = x
    myMaze.PlayerY = y
    myMaze.MazeData(x, y) = myMaze.MazeData(x, y) Or ALREADY_VISITED
End Sub
Public Sub StartMaze()
    If myMaze.Initialized = True Then
        SetPlayerPosition myMaze.EntranceX, myMaze.EntranceY
        myMaze.NumberOfMoves = 0
        DrawMaze
    Else
        MsgBox "Maze not initialized!"
    End If
End Sub

Public Sub DrawMaze(Optional bDontClear As Boolean)
    'simply loop through all maze blocks and draw the blocks individually
    If bDontClear = False Then
        myMaze.OutputPictureBox.Cls
    End If
    Dim x As Integer, y As Integer
    
    For x = 1 To myMaze.Cols
        For y = 1 To myMaze.Rows
            DrawBlock x, y
        Next y
    Next x
    
End Sub

Private Sub DrawBlock(x As Integer, y As Integer, Optional bShowSolution As Boolean)
    'for the actual starting x and y coordinates on the output window
    Dim CoX As Single, CoY As Single

    Dim i As Single
    
    With myMaze
        CoX = .Margin + (x - 1) * .ColWidth     'calculate x position (top left corner)
        CoY = .Margin + (y - 1) * .RowWidth     'calculate y position
        
        If .MazeData(x, y) And ALREADY_VISITED Then
            'color in the maze spot
            .OutputPictureBox.Line (CoX + 1, CoY + 1)-(CoX + .ColWidth - 1, CoY + .RowWidth - 1), myMaze.clrVisited, BF
        End If
                
        If bShowSolution And .MazeData(x, y) And ESCAPE_PATH Then
            'color in the solution path
            .OutputPictureBox.Line (CoX, CoY)-(CoX + .ColWidth, CoY + .RowWidth), myMaze.clrSolution, BF
        End If
        
        If .PlayerX = x And .PlayerY = y Then
            'this is the player's current position
            .OutputPictureBox.Line (CoX + 2, CoY + 2)-(CoX + .ColWidth - 2, CoY + .RowWidth - 2), myMaze.clrCurrent, BF
        End If
        
        If .MazeData(x, y) And NORTH Then
            'draw northern line
            DrawLine CoX, CoY, CoX + .ColWidth, CoY, myMaze.clrWalls
        End If
        
        If .MazeData(x, y) And SOUTH Then
            'draw southern line
            DrawLine CoX, CoY + .RowWidth, CoX + .ColWidth, CoY + .RowWidth, myMaze.clrWalls
        End If

        If .MazeData(x, y) And WEST Then
            'draw western line
            DrawLine CoX, CoY, CoX, CoY + .RowWidth, myMaze.clrWalls
        End If
        
        If .MazeData(x, y) And EAST Then
            'draw eastern line
            DrawLine CoX + .ColWidth, CoY, CoX + .ColWidth, CoY + .RowWidth, myMaze.clrWalls
        End If
            
    End With
        
    
End Sub
Private Sub DrawLine(x As Single, y As Single, X1 As Single, Y1 As Single, clr As Long)
    
    myMaze.OutputPictureBox.Line (x, y)-(X1, Y1), clr
    
    Exit Sub
    
End Sub

Private Sub RemoveWall(x As Integer, y As Integer, nSide As WallData, Optional bInternalCall As Boolean)
    With myMaze
        If nSide And NORTH Then
            'remove northen wall
            If .MazeData(x, y) And NORTH Then   'ensure a wall is there
                .MazeData(x, y) = .MazeData(x, y) Xor NORTH
            End If
            'check if there is a wall above that, and remove its southern wall
            If y > 1 And Not bInternalCall Then RemoveWall x, y - 1, SOUTH, True
        End If
        If nSide And SOUTH Then
            'remove southern wall
            If .MazeData(x, y) And SOUTH Then
                .MazeData(x, y) = .MazeData(x, y) Xor SOUTH
            End If
            If y < .Rows And Not bInternalCall Then RemoveWall x, y + 1, NORTH, True
        End If
        If nSide And EAST Then
            'remove eastern wall
            If .MazeData(x, y) And EAST Then
                .MazeData(x, y) = .MazeData(x, y) Xor EAST
            End If
            If x < .Cols And Not bInternalCall Then RemoveWall x + 1, y, WEST, True
        End If
        If nSide And WEST Then
            'remove western wall
            If .MazeData(x, y) And WEST Then
                .MazeData(x, y) = .MazeData(x, y) Xor WEST
            End If
            If x > 1 And Not bInternalCall Then RemoveWall x - 1, y, EAST, True
        End If
    End With
    

End Sub
Public Sub InstallWall(x As Integer, y As Integer, nSide As WallData, Optional bInternalCall As Boolean)
    With myMaze
        If nSide And NORTH Then
            'install northen wall
            .MazeData(x, y) = .MazeData(x, y) Or NORTH
            'check if there is a wall above that, and remove its southern wall
            If y > 1 And Not bInternalCall Then InstallWall x, y - 1, SOUTH, True
        End If
        If nSide And SOUTH Then
            'install southern wall
            .MazeData(x, y) = .MazeData(x, y) Or SOUTH
            If y < .Rows And Not bInternalCall Then InstallWall x, y + 1, NORTH, True
        End If
        If nSide And EAST Then
            'install eastern wall
            .MazeData(x, y) = .MazeData(x, y) Or EAST
            If x < .Cols And Not bInternalCall Then InstallWall x + 1, y, WEST, True
        End If
        If nSide And WEST Then
            'install western wall
            .MazeData(x, y) = .MazeData(x, y) Or WEST
            If x > 1 And Not bInternalCall Then InstallWall x - 1, y, EAST, True
        End If
    End With
    
End Sub

Private Sub ClearVisitedPath()
    'clear the visitied flag from the maze
    Dim x As Integer, y As Integer
    For x = 1 To myMaze.Cols
        For y = 1 To myMaze.Rows
            If myMaze.MazeData(x, y) And ALREADY_VISITED Then
                myMaze.MazeData(x, y) = myMaze.MazeData(x, y) Xor ALREADY_VISITED
            End If
        Next y
    Next x
        
End Sub
Private Sub ClearEscapePath()
    'clear the escape path flag from the maze
    Dim x As Integer, y As Integer
    For x = 1 To myMaze.Cols
        For y = 1 To myMaze.Rows
            If myMaze.MazeData(x, y) And ESCAPE_PATH Then
                myMaze.MazeData(x, y) = myMaze.MazeData(x, y) Xor ESCAPE_PATH
            End If
        Next y
    Next x
End Sub
Public Sub ShowSolution()
    
    Dim x As Integer, y As Integer
    For x = 1 To myMaze.Cols
        For y = 1 To myMaze.Rows
            If myMaze.MazeData(x, y) And ESCAPE_PATH Then
                DrawBlock x, y, True
            End If
        Next y
    Next x
    
End Sub

Public Function MoveDirection(direction As WallData) As Boolean
    Dim bCanMove As Boolean
    Dim oldX As Integer
    Dim oldY As Integer
    
    oldX = myMaze.PlayerX
    oldY = myMaze.PlayerY
    
    Select Case direction
        Case NORTH
            If myMaze.MazeData(myMaze.PlayerX, myMaze.PlayerY) And NORTH Then
                bCanMove = False
            Else
                If myMaze.PlayerY > 1 Then
                    myMaze.PlayerY = myMaze.PlayerY - 1
                    bCanMove = True
                End If
            End If
        Case SOUTH
            If myMaze.MazeData(myMaze.PlayerX, myMaze.PlayerY) And SOUTH Then
                bCanMove = False
            Else
                myMaze.PlayerY = myMaze.PlayerY + 1
                bCanMove = True
            End If
        Case EAST
            If myMaze.MazeData(myMaze.PlayerX, myMaze.PlayerY) And EAST Then
                bCanMove = False
            Else
                myMaze.PlayerX = myMaze.PlayerX + 1
                bCanMove = True
            End If
        
        Case WEST
            If myMaze.MazeData(myMaze.PlayerX, myMaze.PlayerY) And WEST Then
                bCanMove = False
            Else
                If myMaze.PlayerX > 1 Then
                    myMaze.PlayerX = myMaze.PlayerX - 1
                    bCanMove = True
                End If
            End If
    End Select
    DrawBlock oldX, oldY
    

    If myMaze.PlayerX > myMaze.Cols Or myMaze.PlayerY > myMaze.Rows Then
        DrawMaze
        ShowSolution
        MsgBox "Congratulations, you solved the maze in " & myMaze.NumberOfMoves & " moves!", vbInformation, "MazeMaker"
        End
    End If
    
    If bCanMove Then
        myMaze.MazeData(myMaze.PlayerX, myMaze.PlayerY) = myMaze.MazeData(myMaze.PlayerX, myMaze.PlayerY) Or ALREADY_VISITED
        myMaze.NumberOfMoves = myMaze.NumberOfMoves + 1
        MoveDirection = True
        DrawMaze True
    Else
        Beep
    End If
    
        
End Function
Public Function GetWallCount(x, y) As Integer
    'returns the number of sides to a wall, used in the maze to ensure that
    'no wide open spaces appear, and that at least some sides have been taken
    'off a wall
    Dim wallval As Integer
    Dim count As Integer
    wallval = myMaze.MazeData(x, y)
    If wallval And NORTH Then
        count = 1
    End If
    If wallval And SOUTH Then
        count = count + 1
    End If
    If wallval And EAST Then
        count = count + 1
    End If
    If wallval And WEST Then
        count = count + 1
    End If
    
    GetWallCount = count
End Function

Private Sub CloseShortcuts()
    'this function makes the maze much harder. it follows the solution path,
    'and takes every exit it can find, spawning a recursive function.
    '
    'the recursive functions will also spider the maze, and look to see if they
    'ever reach the exit path again. if they do, the close off the shortcut
    '
    'sets the ALREADY_VISITED FLAG when an area is checked
    Dim x As Integer, y As Integer
    Dim nextX As Integer, nextY As Integer
    Dim originDir As WallData
    Dim nextOriginDir As WallData
    
    With myMaze
        x = .EntranceX
        y = .EntranceY
        
        Do Until (x = .ExitX And y = .ExitY)
            SetPlayerPosition x, y
            
            'Find all open exits, recurse into them, and then find the next part of the
            'escape path
            If y > 1 Then
                If (.MazeData(x, y) And NORTH) = 0 Then
                    'north direction open, see whats there
                    
                    If Not (.MazeData(x, y - 1) And ESCAPE_PATH) = ESCAPE_PATH Then
                        If Not (.MazeData(x, y - 1) And ALREADY_VISITED) = ALREADY_VISITED Then
                            CloseShortcutHelper x, y - 1, SOUTH
                        End If
                    Else
                        If originDir <> NORTH Then
                            nextX = x
                            nextY = y - 1
                            nextOriginDir = SOUTH
                        End If
                    End If
                End If
            End If
            If y < myMaze.Rows Then
                If (.MazeData(x, y) And SOUTH) = 0 Then
                    'south direction open, see whats there
                    If Not (.MazeData(x, y + 1) And ESCAPE_PATH) = ESCAPE_PATH Then
                        If Not (.MazeData(x, y + 1) And ALREADY_VISITED) = ALREADY_VISITED Then
                            CloseShortcutHelper x, y + 1, NORTH
                        End If
                    Else
                        If originDir <> SOUTH Then
                            nextX = x
                            nextY = y + 1
                            nextOriginDir = NORTH
                        End If
                    End If
                End If
            End If
            If x < myMaze.Cols Then
                If (.MazeData(x, y) And EAST) = 0 Then
                    'south direction open, see whats there
                    If Not (.MazeData(x + 1, y) And ESCAPE_PATH) = ESCAPE_PATH Then
                        If Not (.MazeData(x + 1, y) And ALREADY_VISITED) = ALREADY_VISITED Then
                            CloseShortcutHelper x + 1, y, WEST
                        End If
                    Else
                        If originDir <> EAST Then
                            nextX = x + 1
                            nextY = y
                            nextOriginDir = WEST
                        End If
                    End If
                End If
            End If
            If x > 1 Then
                If (.MazeData(x, y) And WEST) = 0 Then
                    'west direction open, see whats there
                    If Not (.MazeData(x - 1, y) And ESCAPE_PATH) = ESCAPE_PATH Then
                        If Not (.MazeData(x - 1, y) And ALREADY_VISITED) = ALREADY_VISITED Then
                            CloseShortcutHelper x - 1, y, EAST
                        End If
                    Else
                        If originDir <> WEST Then
                            nextX = x - 1
                            nextY = y
                            nextOriginDir = EAST
                        End If
                    End If
                End If
            End If
            x = nextX
            y = nextY
            'DrawMaze True
            'Sleep 100
            originDir = nextOriginDir
            
        Loop
    End With
End Sub
Private Sub CloseShortcutHelper(ByVal x As Integer, ByVal y As Integer, originDir As WallData)
    
'    DrawMaze
    Dim nextX As Integer, nextY As Integer
    Dim nextOriginDir As WallData
    Dim bExit As Boolean
    Dim nDirectionCount As Integer
    
    With myMaze
        nextX = x
        nextY = y
        Do Until bExit
            If (.MazeData(x, y) And ESCAPE_PATH) = ESCAPE_PATH Then
                Exit Sub
            End If
            
            If x = 0 Or y = 0 Then
                Exit Sub
            End If
            SetPlayerPosition x, y  'sets the alreadyvisited flag at the players position
'            DrawMaze
            nDirectionCount = 3 - GetWallCount(x, y)   'get the number of directions possible, minus the origin
            If nDirectionCount = 0 Then bExit = True 'we are at a 3 walled side, the only place to go is back, so terminate this function
            'Find all open exits, recurse into them, and then find the next part of the
            'escape path
            If y > 1 Then
                If (.MazeData(x, y) And NORTH) = 0 Then
                    'north direction open, see whats there
                    
                    If (.MazeData(x, y - 1) And ALREADY_VISITED) = 0 Then
                        
                        If originDir <> NORTH Then
                            
                            If (.MazeData(x, y - 1) And ESCAPE_PATH) = ESCAPE_PATH Then
'                                SetPlayerPosition x, y: DrawMaze True: ShowSolution
                                InstallWall x, y, NORTH
'                                bExit = True    'terminate this function
                            Else
                                If nDirectionCount > 1 Then
                                    'more than one choice remaining, spawn a new recursion
                                    CloseShortcutHelper x, y - 1, SOUTH
                                    nDirectionCount = nDirectionCount - 1
                                Else
                                    'its our only direction, take it
                                    nextX = x
                                    nextY = y - 1
                                    nextOriginDir = SOUTH
                                    nDirectionCount = nDirectionCount - 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If y < myMaze.Rows Then
                If (.MazeData(x, y) And SOUTH) = 0 Then
                    'south direction open, see whats there
                    
                    If (.MazeData(x, y + 1) And ALREADY_VISITED) = 0 Then
                        
                        If originDir <> SOUTH Then
                            If (.MazeData(x, y + 1) And ESCAPE_PATH) = ESCAPE_PATH Then
'                                SetPlayerPosition x, y: DrawMaze True: ShowSolution
                                InstallWall x, y, SOUTH
'                                bExit = True    'terminate this function
                            Else
                                If nDirectionCount > 1 Then
                                    'more than one choice remaining, spawn
                                    'a new recusion
                                    CloseShortcutHelper x, y + 1, NORTH
                                    nDirectionCount = nDirectionCount - 1
                                Else
                                    'its our only direction, take it
                                    nextX = x
                                    nextY = y + 1
                                    nextOriginDir = NORTH
                                    nDirectionCount = nDirectionCount - 1
                                End If
                            End If
                        End If
                    Else
                        nDirectionCount = nDirectionCount - 1
                    End If
                End If
            End If
            If x < myMaze.Cols Then
                If (.MazeData(x, y) And EAST) = 0 Then
                    'east direction open, see whats there
                    If (.MazeData(x + 1, y) And ALREADY_VISITED) = 0 Then
                        
                        If originDir <> EAST Then
                            If (.MazeData(x + 1, y) And ESCAPE_PATH) = ESCAPE_PATH Then
'                                SetPlayerPosition x, y: DrawMaze True: ShowSolution
                                InstallWall x, y, EAST
'                                bExit = True    'terminate this function
                            Else
                                If nDirectionCount > 1 Then
                                    'more than one choice remaining, spawn
                                    'a new recusion
                                    CloseShortcutHelper x + 1, y, WEST
                                    nDirectionCount = nDirectionCount - 1
                                Else
                                    'its our only direction, take it
                                    nextX = x + 1
                                    nextY = y
                                    nextOriginDir = WEST
                                    nDirectionCount = nDirectionCount - 1
                                End If
                            End If
                        End If
                    Else
                        'its already been visited
                        nDirectionCount = nDirectionCount - 1
                    End If
                End If
            End If
            If x > 1 Then
                If (.MazeData(x, y) And WEST) = 0 Then
                    'west direction open, see whats there
                    If (.MazeData(x - 1, y) And ALREADY_VISITED) = 0 Then   'hasn't been visited
                        
                        If (.MazeData(x - 1, y) And ESCAPE_PATH) = ESCAPE_PATH Then
'                            SetPlayerPosition x, y: DrawMaze True: ShowSolution
                            InstallWall x, y, WEST
'                            bExit = True    'terminate this function
                        Else
                            If originDir <> WEST Then
                                If nDirectionCount > 1 Then
                                    'more than one choice remaining, spawn
                                    'a new recusion
                                    CloseShortcutHelper x - 1, y, EAST
                                    nDirectionCount = nDirectionCount - 1
                                Else
                                    'its our only direction, take it
                                    nextX = x - 1
                                    nextY = y
                                    nextOriginDir = EAST
                                    nDirectionCount = nDirectionCount - 1
                                End If
                            End If
                        End If
                    Else
                        'its already been visited
                        nDirectionCount = nDirectionCount - 1
                    End If
                End If
            End If
            If (x = nextX And y = nextY) Then
                'allready examined entire area, terminate this function
                bExit = True
            Else
                'more to explore!
                If x = 0 And y = 0 Then
                    MsgBox "error"
                End If
                If nextX = 0 Or nextY = 0 Then
                    MsgBox "error"
                End If
                x = nextX
                y = nextY
            End If
            originDir = nextOriginDir
            
        Loop
    End With

End Sub
Public Sub ApplyColorScheme()
    
    'Applies the current color scheme to the window
    myMaze.OutputPictureBox.BackColor = myMaze.clrBackground
    frmMaze.BackColor = myMaze.clrBackground
    
    'Draw the maze (will clear the screen and draw appropriately)
    If myMaze.Initialized = True Then
        DrawMaze
    End If
    
End Sub
