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
      TabIndex        =   1
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
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label lblTile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "1024"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   2000
      Visible         =   0   'False
      Width           =   855
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
Dim marginLeft As Integer
Dim marginTop  As Integer
Dim lineWidth As Integer
Const cells = 4
Dim gameCells() As Integer

Dim availWidth As Integer
Dim availHeight As Integer
    
Dim iRowHeight As Integer
Dim iColWidth As Integer

Dim shownCongrats As Boolean

    
Private Sub Form_Activate()
    
    ' sbMainStatusBar.SimpleText = "Window dimensions = " + CStr(frmMain.Width) + "x" + CStr(frmMain.Height)
    Call addLog("frmMain Form_Activate()")
    
    Call DrawGrid
    Call DrawTiles
    
    frmMain.SetFocus
    
End Sub

Private Sub Form_GotFocus()
    Call addLog("frmMain Form_GotFocus()")
    Call DrawGrid
    
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
    
    marginLeft = frmMain.ScaleX(20, vbPixels, vbHimetric)
    marginTop = frmMain.ScaleY(20, vbPixels, vbHimetric)
    availWidth = frmMain.Width - 100
    availHeight = frmMain.Height - sbMainStatusBar.Height - 300
    iRowHeight = (availHeight - 2 * marginTop) / cells
    iColWidth = (availWidth - 2 * marginLeft) / cells
    
    ' ugh, pixels vs these measures, again.
    lineWidth = frmMain.ScaleX(2, vbPixels, vbHimetric)
    frmMain.DrawWidth = 2
    
    Dim i As Integer
    
    For i = 1 To cells * cells
        Load lblTile(i)
    Next i
End Sub
Private Sub InitGame()
    ReDim gameCells(cells - 1, cells - 1) As Integer
    
    ' One needs to 'Load' a control array before using it. Sigh.
    ' but only once... so ignore the 0th item...
    
    Dim cellx As Integer, celly As Integer
    shownCongrats = False
    lblGameOver.Visible = False
    
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
            lblTile(idx).Left = marginLeft + iCol * iColWidth + lineWidth
            lblTile(idx).Top = marginTop + iRow * iRowHeight + lineWidth
            lblTile(idx).Width = iColWidth - 2 * lineWidth
            lblTile(idx).Height = iRowHeight - 2 * lineWidth
            lblTile(idx).Visible = True
            lblTile(idx).Caption = CStr(gameCells(iCol, iRow))
            lblTile(idx).BackColor = ColorByValue(gameCells(iCol, iRow))
            lblTile(idx).Alignment = 2 ' center
        Next iCol
    Next iRow

    Call RandomlyPlace2Or4(gameCells)
    
End Sub
Private Sub DrawGrid()
    ''' Just draw the lines of the grid here
    Dim iRow As Integer
    Dim iCol As Integer
    frmMain.Cls
    
    For iRow = 0 To cells
        frmMain.Line (marginLeft, marginTop + iRow * iRowHeight)-(availWidth - marginLeft, marginTop + iRow * iRowHeight)
        For iCol = 0 To cells
            frmMain.Line (marginLeft + iCol * iColWidth, marginTop)-(marginLeft + iCol * iColWidth, availHeight - marginTop)
        Next iCol
    Next iRow
End Sub
Sub DrawTiles()
    ''' Populate the grid with tiles having potentially some info.
    Dim iRow As Integer
    Dim iCol As Integer
    Dim idx As Integer
        
    For iRow = 0 To cells - 1
        For iCol = 0 To cells - 1
            idx = iCol + cells * iRow
            lblTile(idx).Caption = CStr(gameCells(iCol, iRow))
            lblTile(idx).BackColor = ColorByValue(gameCells(iCol, iRow))
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


Private Function ColorByValue(val As Integer) As Long

    Dim COLOR_GREEN As Long, COLOR_2 As Long, COLOR_4 As Long, COLOR_8 As Long
    Dim COLOR_16 As Long, COLOR_32 As Long, COLOR_64 As Long
    Dim COLOR_128 As Long, COLOR_256 As Long, COLOR_512 As Long
    Dim COLOR_1024 As Long, COLOR_2048 As Long, COLOR_4096 As Long
    Dim COLOR_8192 As Long
    Dim COLOR_EMPTY As Long
    
    COLOR_GREEN = &HFF00&
    COLOR_2 = &HFFFF&
    COLOR_4 = RGB(255, 128, 0)
    COLOR_8 = RGB(255, 64, 0)
    COLOR_16 = RGB(255, 0, 0)
    COLOR_32 = RGB(255, 0, 255)
    COLOR_64 = RGB(255, 0, 128)
    COLOR_128 = RGB(128, 0, 255)
    COLOR_256 = RGB(0, 255, 255)
    COLOR_512 = RGB(0, 128, 0)
    COLOR_1024 = RGB(0, 255, 128)
    COLOR_2048 = RGB(0, 0, 255)
    COLOR_4096 = RGB(192, 192, 192)
    COLOR_8192 = RGB(255, 165, 0)
    
    COLOR_EMPTY = RGB(235, 235, 235)
    
    Select Case val
        Case 0
            ColorByValue = COLOR_EMPTY
        Case 2
            ColorByValue = COLOR_2
        Case 4
            ColorByValue = COLOR_4
        Case 8
            ColorByValue = COLOR_8
        Case 16
            ColorByValue = COLOR_16
        Case 32
            ColorByValue = COLOR_32
        Case 64
            ColorByValue = COLOR_64
        Case 128
            ColorByValue = COLOR_128
        Case 256
            ColorByValue = COLOR_256
        Case 512
            ColorByValue = COLOR_512
        Case 1024
            ColorByValue = COLOR_1024
        Case 2048
            ColorByValue = COLOR_2048
        Case 4096
            ColorByValue = COLOR_4096
        Case 8192
            ColorByValue = COLOR_8192
        Case Else
            ColorByValue = COLOR_GREEN
    End Select
End Function

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
    Call DrawTiles
    Call DrawGrid
End Sub


Private Sub tAutoplay_Timer()
    Call GameStep(gameCells, Directions.Left)
    Call GameStep(gameCells, Directions.Up)
End Sub
