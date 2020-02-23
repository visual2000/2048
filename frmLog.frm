VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "Debug log output"
   ClientHeight    =   3195
   ClientLeft      =   17985
   ClientTop       =   4110
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox txtLog 
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    txtLog.BackColor = vbWhite
    txtLog.ForeColor = vbBlack
    
'    txtLog.Enabled = False
    
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    ' TODO figure out how to get "inner width" of form
    txtLog.Width = frmLog.Width
    txtLog.Height = frmLog.Height
End Sub

Sub addLog(log As String)
    txtLog.Text = txtLog.Text + log + vbNewLine
    txtLog.SelStart = Len(txtLog.Text)
End Sub

