Attribute VB_Name = "Mechanics"
Option Explicit

Public Enum Directions
    Up
    Down
    Left
    Right
End Enum

Function EmptyCellCount(gameCells() As Integer) As Integer
    Dim empties As Collection
    Set empties = GetEmptyCells(gameCells)
    
    EmptyCellCount = empties.Count
End Function

Sub RandomlyPlace2Or4(gameCells() As Integer)
    '' WARNING this will modify the gameCells argument passed in (by ref).
    If EmptyCellCount(gameCells) = 0 Then
        Call addLog("OH NO! You can't place a new tile if there are no empty spots!")
        Exit Sub
    End If
    
    Dim randomTileValue As Integer
    
    ' this will produce a 2 or a 4
    randomTileValue = (Int(Rnd * 2) + 1) * 2
    
    Dim empties As Collection
    Set empties = GetEmptyCells(gameCells)
    
    Dim randomEmptySlot As Integer
    randomEmptySlot = (Int(Rnd * empties.Count) + 1)
    
    gameCells(empties.Item(randomEmptySlot).x, empties.Item(randomEmptySlot).y) = randomTileValue
End Sub

Function GetEmptyCells(gameCells() As Integer) As Collection
    Dim empties As Collection
    Set empties = New Collection
    
    Dim t As Tile
    Dim x As Integer, y As Integer
    
    For y = 0 To UBound(gameCells, 2)
        For x = 0 To UBound(gameCells, 1)
            If gameCells(x, y) = 0 Then
                Set t = New Tile
                t.x = x
                t.y = y
                empties.Add t
            End If
        Next x
    Next y
    
    Set GetEmptyCells = empties
End Function


Function ApplyGravity(gameCells() As Integer, dx As Integer, dy As Integer, stepX As Integer, stepY As Integer, startX As Integer, startY As Integer, endX As Integer, endY As Integer) As Boolean
    
    Dim x As Integer, y As Integer
    Dim needToContinueLoop As Boolean
    Dim weDidAnythingAtAll As Boolean
    weDidAnythingAtAll = False
    
    Do
        needToContinueLoop = False
        
        For y = startY To endY Step stepY ' each row
            'start at the given cell:
            For x = startX To endX Step stepX
                If gameCells(x, y) = 0 Then
                    If x + dx >= 0 And y + dy >= 0 And x + dx <= UBound(gameCells) And y + dy <= UBound(gameCells) Then
                    ' look at our neighbour - if they're non-empty, copy over.
                        If Not gameCells(x + dx, y + dy) = 0 Then
                            'copy across and blank them out!
                            gameCells(x, y) = gameCells(x + dx, y + dy)
                            gameCells(x + dx, y + dy) = 0
                            needToContinueLoop = True
                            weDidAnythingAtAll = True
                        End If
                    End If
                End If
                ' if we're not empty we're done, skip ahead
            Next x
        Next y
    Loop While needToContinueLoop
    
    ApplyGravity = weDidAnythingAtAll
End Function

Function ApplyMerges(gameCells() As Integer, dx As Integer, dy As Integer, stepX As Integer, stepY As Integer, startX As Integer, startY As Integer, endX As Integer, endY As Integer) As Boolean
    Dim x As Integer
    Dim y As Integer
    Dim didWeMerge As Boolean
    didWeMerge = False
    For y = startY To endY Step stepY ' each row
        For x = startX To endX Step stepX
            If Not gameCells(x, y) = 0 Then
                If x + dx >= 0 And y + dy >= 0 And x + dx <= UBound(gameCells) And y + dy <= UBound(gameCells) Then
                    ' look at our neighbour - if they're the same as us, merge.
                    If gameCells(x + dx, y + dy) = gameCells(x, y) Then
                        'double ourselves and blank them!
                        gameCells(x, y) = gameCells(x, y) * 2
                        gameCells(x + dx, y + dy) = 0
                        ' apply gravity on the remainder of this row
                        Call ApplyGravity(gameCells, dx, dy, stepX, stepY, x + dx, y + dy, endX, y + dy)
                        didWeMerge = True
                    End If
                End If
            End If
            ' if the cell is empty, we're done, because that means the rest of the row is empty too (gravity)
        Next x
    Next y
    ApplyMerges = didWeMerge
End Function

Function NeighbouringTwins(gameCells() As Integer) As Boolean

    Dim areThereNeighbours As Boolean
    areThereNeighbours = False
    
    Dim x As Integer, y As Integer
    For y = 0 To UBound(gameCells)
        For x = 0 To UBound(gameCells)
            If x + 1 <= UBound(gameCells) Then
                If gameCells(x, y) = gameCells(x + 1, y) Then
                    areThereNeighbours = True
                End If
            End If
            If y + 1 <= UBound(gameCells) Then
                If gameCells(x, y) = gameCells(x, y + 1) Then
                    areThereNeighbours = True
                End If
            End If
        Next x
    Next y

    NeighbouringTwins = areThereNeighbours
End Function

Sub GameStep(gameCells() As Integer, direction As Directions)
    Dim dx As Integer, dy As Integer
    Dim stepX As Integer, stepY As Integer
    Dim startX As Integer, startY As Integer
    Dim endX As Integer, endY As Integer
    
    Select Case direction
        Case Directions.Right
            dx = -1
            dy = 0
            stepX = -1
            stepY = 1
            startX = UBound(gameCells)
            startY = 0
            endX = 0
            endY = UBound(gameCells)
        Case Directions.Left
            dx = 1
            dy = 0
            stepX = 1
            stepY = 1
            startX = 0
            startY = 0
            endX = UBound(gameCells)
            endY = UBound(gameCells)
        Case Directions.Up
            dx = 0
            dy = 1
            stepX = 1
            stepY = 1
            startX = 0
            startY = 0
            endX = UBound(gameCells)
            endY = UBound(gameCells)
        Case Directions.Down
            dx = 0
            dy = -1
            stepX = -1
            stepY = -1
            startX = UBound(gameCells)
            startY = UBound(gameCells)
            endX = 0
            endY = 0
        Case Else
            Call addLog("exiting because of unimplemented direction")
            Exit Sub
    End Select

    If EmptyCellCount(gameCells) = 0 Then
        Call addLog("GameStep(): no empty cells!  Check if there are legal moves left.")
        If NeighbouringTwins(gameCells) Then
            Call addLog("GameStep(): there are, however, valid moves left.")
        Else
            Call addLog("GameStep(): there are no valid moves left. Game over!")
            Exit Sub
        End If
    End If
    
    Dim didGravityMove As Boolean, didMergeMove As Boolean
    didGravityMove = ApplyGravity(gameCells, dx, dy, stepX, stepY, startX, startY, endX, endY)
    didMergeMove = ApplyMerges(gameCells, dx, dy, stepX, stepY, startX, startY, endX, endY)
    
    If didMergeMove Or didGravityMove Or frmMain.mnuCheatAlwaysGive.Checked Then
        Call RandomlyPlace2Or4(gameCells)
    End If
    
    Call frmMain.DrawTiles
    Call frmMain.UpdateScore
End Sub
