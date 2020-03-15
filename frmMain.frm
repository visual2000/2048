VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "2048"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tAutoplay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6960
      Top             =   480
   End
   Begin MSComctlLib.StatusBar sbMainStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5220
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Game over!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3135
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image imgTile 
      Height          =   615
      Index           =   0
      Left            =   240
      Top             =   480
      Width           =   495
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameDebug 
         Caption         =   "Show debug log"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuCheats 
      Caption         =   "Cheats"
      Begin VB.Menu mnuCheatAlwaysGive 
         Caption         =   "Always give new tiles, even without move"
      End
      Begin VB.Menu mnuCheatsAutoplay 
         Caption         =   "Autoplay left/up"
         Shortcut        =   ^P
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cells = 4
Dim gameCells() As Integer
    
Dim iRowHeight As Integer
Dim iColWidth As Integer

Dim shownCongrats As Boolean

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Activate()
    Dim animationSteps As Collection
    Set animationSteps = New Collection
    
    Call addLog("frmMain Form_Activate()")
    
    Call DrawTiles(animationSteps)
    
    frmMain.SetFocus
    
End Sub

Private Sub Form_GotFocus()
    Call addLog("frmMain Form_GotFocus()")
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case 72, 37
        Call GameStep(gameCells, Directions.Left)
    Case 74, 40
        Call GameStep(gameCells, Directions.Down)
    Case 75, 38
        Call GameStep(gameCells, Directions.Up)
    Case 76, 39
        Call GameStep(gameCells, Directions.Right)
    End Select
    
End Sub

Private Sub Form_Load()
    Randomize
    Call InitWindow
    Call InitGame
    
End Sub
Private Sub InitWindow()
    ''' It appears that VB6 size units are tenth of pixels?
    
    ' they're himetrics, probably.  Views have ScaleX and ScaleY methods
    
    iRowHeight = frmMain.ScaleY(LoadResPicture(101, vbResBitmap).Height, vbHimetric, vbTwips)
    iColWidth = frmMain.ScaleX(LoadResPicture(101, vbResBitmap).Width, vbHimetric, vbTwips)
    
    Dim i As Integer
    
    For i = 1 To cells * cells - 1
        Load imgTile(i)
    Next i
    
End Sub

Private Function CellResourceId(value As Integer) As Integer
    Select Case value
        Case 0
            CellResourceId = 199
        Case Else
            CellResourceId = Int(Math.log(value) / Math.log(2)) + 100
    End Select
End Function

Private Sub InitGame()
    ReDim gameCells(cells - 1, cells - 1) As Integer
    
    ' One needs to 'Load' a control array before using it. Sigh.
    ' but only once... so ignore the 0th item...
    
    Dim cellx As Integer, celly As Integer
    shownCongrats = False
    lblGameOver.Visible = False
    
    frmMain.Width = iColWidth * cells + (frmMain.Width - frmMain.ScaleWidth)
    frmMain.Height = iRowHeight * cells + (frmMain.Height - frmMain.ScaleHeight) + sbMainStatusBar.Height
    
    For celly = 0 To cells - 1
        For cellx = 0 To cells - 1
            ' Set all cells to empty, initially
            gameCells(cellx, celly) = 0
        Next cellx
    Next celly
    
    ''' Populate the grid with tiles having potentially some info.
    Dim iRow As Integer
    Dim iCol As Integer
    Dim idx As Integer
    
    For iRow = 0 To cells - 1
        For iCol = 0 To cells - 1
            idx = iCol + cells * iRow
            imgTile(idx).Left = iCol * iColWidth
            imgTile(idx).Top = iRow * iRowHeight
            imgTile(idx).Width = iColWidth
            imgTile(idx).Height = iRowHeight
            imgTile(idx).Visible = True
            imgTile(idx).Picture = LoadResPicture(CellResourceId(gameCells(iCol, iRow)), vbResBitmap)
        Next iCol
    Next iRow

    Call RandomlyPlace2Or4(gameCells)
    
End Sub

Sub DrawTiles(animationSteps As Collection)
    ''' Populate the grid with tiles having potentially some info.
    Dim iRow As Integer
    Dim iCol As Integer
    Dim idx As Integer
    
    Dim animationStep As animationStep
    Dim animationFrame As Integer
    Dim animationFPS As Integer
    animationFPS = 5
    Dim startX As Integer
    Dim startY As Integer
    
    For animationFrame = 1 To animationFPS
        For Each animationStep In animationSteps
            If Not animationStep.amIaMerge Then
                idx = animationStep.startX + cells * animationStep.startY
                imgTile(idx).ZOrder (0)
                imgTile(idx).Left = imgTile(idx).Left + (1# / animationFPS) * iColWidth * (animationStep.endX - animationStep.startX)
                imgTile(idx).Top = imgTile(idx).Top + (1# / animationFPS) * iRowHeight * (animationStep.endY - animationStep.startY)
                Sleep 200 / animationFPS
            End If
        Next animationStep
    Next animationFrame
    
    
    
        
    For iRow = 0 To cells - 1
        For iCol = 0 To cells - 1
            idx = iCol + cells * iRow
            imgTile(idx).Left = iCol * iColWidth
            imgTile(idx).Top = iRow * iRowHeight
            imgTile(idx).Picture = LoadResPicture(CellResourceId(gameCells(iCol, iRow)), vbResBitmap)
        Next iCol
    Next iRow
End Sub

Sub UpdateScore()
    Dim score As Integer
    score = 0
    
    Dim reached2048 As Boolean
    reached2048 = False
        
    Dim x As Integer, y As Integer
    For x = 0 To cells - 1
        For y = 0 To cells - 1
            score = score + gameCells(x, y)
            If gameCells(x, y) >= 2048 Then
                reached2048 = True
            End If
        Next y
    Next x
    
    sbMainStatusBar.SimpleText = "Your score: " + CStr(score)
    
    If reached2048 And Not shownCongrats Then
        MsgBox ("Congratulations! You reached 2048!")
        shownCongrats = True
    End If
    
    Dim emptyCells As Integer
    emptyCells = EmptyCellCount(gameCells)
   
    If emptyCells = 0 Then
        If Not NeighbouringTwins(gameCells) Then
            lblGameOver.Visible = True
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmLog
    Set frmLog = Nothing
End Sub

Private Sub mnuCheatAlwaysGive_Click()
    mnuCheatAlwaysGive.Checked = Not mnuCheatAlwaysGive.Checked
End Sub

Private Sub mnuCheatsAutoplay_Click()
    mnuCheatsAutoplay.Checked = Not mnuCheatsAutoplay.Checked
    tAutoplay.Enabled = mnuCheatsAutoplay.Checked
End Sub

Private Sub mnuGameDebug_Click()
    mnuGameDebug.Checked = Not mnuGameDebug.Checked
    frmLog.Visible = mnuGameDebug.Checked
End Sub

Private Sub mnuGameExit_Click()
    Unload frmMain
    Set frmMain = Nothing
End Sub

Private Sub mnuGameNew_Click()
    Call InitGame
    Call DrawTiles(New Collection)
End Sub


Private Sub tAutoplay_Timer()
    Call GameStep(gameCells, Directions.Left)
    Call GameStep(gameCells, Directions.Up)
End Sub

