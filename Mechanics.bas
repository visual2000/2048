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

Function ApplyGravity(gameCells() As Integer, dx As Integer, dy As Integer, stepX As Integer, stepY As Integer, startX As Integer, startY As Integer, endX As Integer, endY As Integer) As Collection
    
    Dim x As Integer, y As Integer
    Dim needToContinueLoop As Boolean
    
    Dim a As AnimationStep
    Dim animationSteps As Collection
    Set animationSteps = New Collection
    
    Dim gravityStartX As Integer
    Dim gravityStartY As Integer
    Dim gravityEndX As Integer
    Dim gravityEndY As Integer
    
    Dim walk As Integer
    
    Dim fullCellFound As Boolean
    
    Do
        needToContinueLoop = False
        
        For y = startY To endY Step stepY ' each row
            'start at the given cell:
            For x = startX To endX Step stepX
                If gameCells(x, y) = 0 Then
                    gravityEndX = x
                    gravityEndY = y
                    fullCellFound = False
                    
                    If x + dx >= 0 And y + dy >= 0 And x + dx <= UBound(gameCells) And y + dy <= UBound(gameCells) Then
                    
                        If dx = 0 Then
                            For walk = y + dy To endY Step stepY
                                ' if we've walked to a cell that's full, record its position, jump out of the foor loop
                                If gameCells(x, walk) > 0 Then
                                    ' record it
                                    gravityStartX = x
                                    gravityStartY = walk
                                    fullCellFound = True
                                    Exit For
                                End If
                            Next walk
                        End If
                        
                        If dy = 0 Then
                            For walk = x + dx To endX Step stepX
                                ' if we've walked to a cell that's full, record its position, jump out of the foor loop
                                If gameCells(walk, y) > 0 Then
                                    ' record it
                                    gravityStartX = walk
                                    gravityStartY = y
                                    fullCellFound = True
                                    Exit For
                                End If
                            Next walk
                        End If
                        
                        If fullCellFound Then
                            gameCells(gravityEndX, gravityEndY) = gameCells(gravityStartX, gravityStartY)
                            gameCells(gravityStartX, gravityStartY) = 0
                            needToContinueLoop = True
                            ' Add to the list of proposed moves that goes back to the form
                            Set a = New AnimationStep
                            a.startX = gravityStartX
                            a.startY = gravityStartY
                            a.endX = gravityEndX
                            a.endY = gravityEndY
                            a.cellValue = gameCells(gravityEndX, gravityEndY)
                            animationSteps.Add a
                        End If
                    End If
                End If
                ' if we're not empty we're done, skip ahead
            Next x
        Next y
    Loop While needToContinueLoop
    
    Set ApplyGravity = animationSteps
End Function

Function ApplyMerges(gameCells() As Integer, dx As Integer, dy As Integer, _
                     stepX As Integer, stepY As Integer, _
                     startX As Integer, startY As Integer, _
                     endX As Integer, endY As Integer) As Collection
                     
    Dim x As Integer
    Dim y As Integer
    Dim a As AnimationStep
    Dim animationSteps As Collection
    Set animationSteps = New Collection
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
                        Set animationSteps = appendCollection(animationSteps, ApplyGravity(gameCells, dx, dy, stepX, stepY, x + dx, y + dy, endX, y + dy))
                        Set a = New AnimationStep
                        a.startX = x + dx
                        a.startY = y + dy
                        a.endX = x
                        a.endY = y
                        a.amIaMerge = True
                        a.cellValue = gameCells(x, y)
                        animationSteps.Add a
                    End If
                End If
            End If
            ' if the cell is empty, we're done, because that means the rest of the row is empty too (gravity)
        Next x
    Next y
    Set ApplyMerges = animationSteps
End Function

Function appendCollection(a As Collection, b As Collection) As Collection
    Dim newCollection As Collection
    Set newCollection = a
    Dim thing As Object
    For Each thing In b
        newCollection.Add thing
    Next thing
    
    Set appendCollection = newCollection
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

Function GameStep(gameCells() As Integer, direction As Directions) As Collection
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
            Exit Function
    End Select

    If EmptyCellCount(gameCells) = 0 Then
        Call addLog("GameStep(): no empty cells!  Check if there are legal moves left.")
        If NeighbouringTwins(gameCells) Then
            Call addLog("GameStep(): there are, however, valid moves left.")
        Else
            Call addLog("GameStep(): there are no valid moves left. Game over!")
            Set GameStep = New Collection
            Exit Function
        End If
    End If
    
    Dim animationSteps As Collection
    
    Dim didGravityMove As Boolean, didMergeMove As Boolean
    Set animationSteps = ApplyGravity(gameCells, dx, dy, stepX, stepY, startX, startY, endX, endY)
    didGravityMove = animationSteps.Count > 0
    
    Dim mergeAnimationSteps As Collection
    Set mergeAnimationSteps = ApplyMerges(gameCells, dx, dy, stepX, stepY, startX, startY, endX, endY)
    didMergeMove = mergeAnimationSteps.Count > 0
    
    ' we want to do more moves below, potentially, so let's save these:
    Set mergeAnimationSteps = appendCollection(mergeAnimationSteps, animationSteps)
    Set animationSteps = New Collection
    
    If didMergeMove Or didGravityMove Or frmMain.mnuCheatAlwaysGive.Checked Then
        Set animationSteps = ApplyGravity(gameCells, dx, dy, stepX, stepY, startX, startY, endX, endY)
        Call RandomlyPlace2Or4(gameCells)
    End If
    
    Set GameStep = appendCollection(animationSteps, mergeAnimationSteps)
End Function
