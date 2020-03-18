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
' Load all the sprites
' Draw the board (for every step, takes the game state)
' Unload

Public Sub DrawBoard()

    Dim iRow As Integer, iCol As Integer
    Dim idx As Integer
    
    BitBlt myBackBuffer, 0, 0, 64 * frmMain.cells, 64 * frmMain.cells, 0, 0, 0, vbWhiteness
    
    Dim backgroundTile As Long
    backgroundTile = sprites(0)
    
    For iRow = 0 To frmMain.cells - 1
        For iCol = 0 To frmMain.cells - 1
            idx = iCol + frmMain.cells * iRow
            BitBlt myBackBuffer, iCol * 64, iRow * 64, 64, 64, backgroundTile, 0, 0, vbSrcCopy
        Next iCol
    Next iRow

    BitBlt frmMain.pbCanvas.hdc, 0, 0, 64 * frmMain.cells, _
    64 * frmMain.cells, myBackBuffer, 0, 0, vbSrcCopy
End Sub
    
Public Sub loadAllSprites()

    ReDim sprites(0 To 14) As Long
    
    sprites(0) = LoadGraphicDC(199) ' this is our BG tile
    ' etc
    
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
