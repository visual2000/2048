Attribute VB_Name = "Sprites"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" _
 (ByVal hDestDC As Long, ByVal X As Long, _
 ByVal Y As Long, ByVal nWidth As Long, _
 ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc _
 As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
 
'creating buffers / loading sprites
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

'loading sprites
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'cleanup
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Dim sprites() As Long
'our Buffer's DC

Public myBackBuffer As Long

Public myBufferBMP As Long

' Initialisation - backbuffer bmp
Public Sub initialiseGraphics()
    'create a compatable DC for the back buffer..
    myBackBuffer = CreateCompatibleDC(GetDC(0))

    'create a compatible bitmap surface for the DC
    'that is the size of our form.
    'NOTE - the bitmap will act as the actual graphics surface inside the DC
    'because without a bitmap in the DC, the DC cannot hold graphical data..
    myBufferBMP = CreateCompatibleBitmap(GetDC(0), 64 * frmMain.cells, 64 * frmMain.cells)
    
    'final step of making the back buffer...
    'load our created blank bitmap surface into our buffer
    '(this will be used as our canvas to draw-on off screen)
    SelectObject myBackBuffer, myBufferBMP
    
    'before we can blit to the buffer, we should fill it with black
    BitBlt myBackBuffer, 0, 0, 64 * frmMain.cells, 64 * frmMain.cells, 0, 0, 0, vbWhiteness
    
    loadAllSprites
End Sub
Public Sub unloadAll()
    'this clears up the memory we used to hold
    'the graphics and the buffers we made
    
    'Delete the bitmap surface that was in the backbuffer
    DeleteObject myBufferBMP
    
    'Delete the backbuffer HDC
    DeleteDC myBackBuffer
    
    'Delete the Sprite/Graphic HDC
    Dim i As Integer
    For i = 0 To UBound(sprites)
        DeleteDC sprites(i)
    Next i
End Sub

Private Function CellSpriteId(value As Integer) As Integer
    Dim spriteId As Integer
    
    If value = 0 Then
        spriteId = 0
    Else
        spriteId = Int(Math.log(value) / Math.log(2))
    End If

    If spriteId >= 0 And spriteId <= UBound(sprites) Then
        CellSpriteId = spriteId
    Else
        Call addLog("You asked for a nonexistent sprite, id = " & CStr(value))
        CellSpriteId = -1
    End If
End Function

Public Sub DrawBoard(gameCells() As Integer)

    Dim iRow As Integer, iCol As Integer
    Dim idx As Integer
    
    BitBlt myBackBuffer, 0, 0, 64 * frmMain.cells, 64 * frmMain.cells, 0, 0, 0, vbWhiteness
    
    For iRow = 0 To frmMain.cells - 1
        For iCol = 0 To frmMain.cells - 1
            idx = iCol + frmMain.cells * iRow
            
            Debug.Print ("cell value: " & CStr(gameCells(iCol, iRow)))
            Debug.Print ("sprite id: " & CStr(CellSpriteId(gameCells(iCol, iRow))))
            
            BitBlt myBackBuffer, iCol * 64, iRow * 64, 64, 64, _
            sprites(CellSpriteId(gameCells(iCol, iRow))), 0, 0, vbSrcCopy
        Next iCol
    Next iRow

    frmMain.pbCanvas.Cls
    BitBlt frmMain.pbCanvas.hdc, 0, 0, 64 * frmMain.cells, _
           64 * frmMain.cells, myBackBuffer, 0, 0, vbSrcCopy
End Sub
    
Public Sub loadAllSprites()

    ReDim sprites(0 To 14) As Long
    
    sprites(0) = LoadGraphicDC(199) ' this is our BG tile
    sprites(1) = LoadGraphicDC(101)
    sprites(2) = LoadGraphicDC(102)
    sprites(3) = LoadGraphicDC(103)
    sprites(4) = LoadGraphicDC(104)
    sprites(5) = LoadGraphicDC(105)
    sprites(6) = LoadGraphicDC(106)
    sprites(7) = LoadGraphicDC(107)
    sprites(8) = LoadGraphicDC(108)
    sprites(9) = LoadGraphicDC(109)
    sprites(10) = LoadGraphicDC(110)
    sprites(11) = LoadGraphicDC(111)
    sprites(12) = LoadGraphicDC(112)
    sprites(13) = LoadGraphicDC(113)
    Debug.Print "loaded sprites"
    
End Sub


Public Function LoadGraphicDC(iSpriteId As Integer) As Long
    'temp variable to hold our DC address
    Dim LoadGraphicDCTEMP As Long
    
    'create the DC address compatible with
    'the DC of the screen
    LoadGraphicDCTEMP = CreateCompatibleDC(GetDC(0))
    
    'load the graphic file into the DC...
    SelectObject LoadGraphicDCTEMP, LoadResPicture(iSpriteId, vbResBitmap)
    
    'return the address of the file
    LoadGraphicDC = LoadGraphicDCTEMP
End Function
