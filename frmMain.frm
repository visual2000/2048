VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "2048"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7875
   Icon            =   "frmMain.frx":0000
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
   Begin VB.PictureBox pbCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1200
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu d1 
         Caption         =   "-"
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
      Begin VB.Menu d2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameDebug 
         Caption         =   "Show debug log"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cells As Integer
Dim gameCells() As Integer
    
Dim cellPx As Integer

Dim shownCongrats As Boolean

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Activate()
    Dim animationSteps As Collection
    Set animationSteps = New Collection
    
    Call addLog("frmMain Form_Activate()")
       
    frmMain.SetFocus
End Sub

Private Sub Form_GotFocus()
    Call addLog("frmMain Form_GotFocus()")
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Call handleKey(KeyCode)
End Sub

Sub handleKey(KeyCode As Integer)
    Dim dummy As Collection
    Set dummy = New Collection
    
    Select Case KeyCode
    Case vbKeyH, vbKeyLeft
        Set dummy = GameStep(gameCells, Directions.Left)
    Case vbKeyJ, vbKeyDown
        Set dummy = GameStep(gameCells, Directions.Down)
    Case vbKeyK, vbKeyUp
        Set dummy = GameStep(gameCells, Directions.Up)
    Case vbKeyL, vbKeyRight
        Set dummy = GameStep(gameCells, Directions.Right)
    End Select
    
    Call Animate(gameCells, dummy)
    Call DrawBoard(gameCells, False)
    Call UpdateScore
End Sub

Private Sub Form_Load()
    Randomize
    
    cells = 4
    initialiseGraphics
    
    Call InitWindow
    Call InitGame
End Sub
Private Sub InitWindow()
    ''' It appears that VB6 size units are tenth of pixels?
    ' they're himetrics, probably.  Views have ScaleX and ScaleY methods
    
    cellPx = frmMain.ScaleY(LoadResPicture(101, vbResBitmap).Height, vbHimetric, vbPixels)
End Sub

Private Sub InitGame()
    ReDim gameCells(cells - 1, cells - 1) As Integer
    
    Dim cellx As Integer, celly As Integer
    shownCongrats = False
    
    frmMain.Width = Screen.TwipsPerPixelX * cellPx * cells + (frmMain.Width - frmMain.ScaleWidth)
    frmMain.Height = Screen.TwipsPerPixelY * cellPx * cells + (frmMain.Height - frmMain.ScaleHeight) + sbMainStatusBar.Height
    
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
    
    pbCanvas.Width = frmMain.ScaleX(cellPx * cells, vbPixels, vbTwips)
    pbCanvas.Height = frmMain.ScaleY(cellPx * cells, vbPixels, vbTwips)
    pbCanvas.Left = 0
    pbCanvas.Top = 0
    pbCanvas.Visible = True
    
    Call RandomlyPlace2Or4(gameCells)
    Call DrawBoard(gameCells, False)
    Call UpdateScore

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
        Call addLog("there are 0 empty cells!")
        If Not NeighbouringTwins(gameCells) Then
            Call DrawBoard(gameCells, True)
        End If
    End If
    
End Sub

Private Sub Form_Terminate()
    unloadAll
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
    Call DrawBoard(gameCells, False)
End Sub

Private Sub pbCanvas_KeyUp(KeyCode As Integer, Shift As Integer)
    ' if the picture box happens to have focus and gets keyevents,
    ' send them through to our other handler.
    Call handleKey(KeyCode)
End Sub

Private Sub tAutoplay_Timer()
    Dim dummy As Collection
    Set dummy = GameStep(gameCells, Directions.Left)
    Call DrawBoard(gameCells, False)
    Call UpdateScore

    Set dummy = GameStep(gameCells, Directions.Up)
    Call DrawBoard(gameCells, False)
    Call UpdateScore

End Sub

