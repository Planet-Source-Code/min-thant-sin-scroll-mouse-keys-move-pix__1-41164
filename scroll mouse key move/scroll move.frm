VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Move pix with mouse, arrow keys and scrollbars"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar vsb 
      Height          =   3690
      Left            =   4875
      TabIndex        =   3
      Top             =   75
      Width           =   240
   End
   Begin VB.HScrollBar hsb 
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   3750
      Width           =   4815
   End
   Begin VB.PictureBox picCont 
      BackColor       =   &H00808080&
      Height          =   3690
      Left            =   75
      ScaleHeight     =   3630
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   75
      Width           =   4815
      Begin VB.PictureBox picMove 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   1740
         Left            =   525
         MousePointer    =   5  'Size
         Picture         =   "scroll move.frx":0000
         ScaleHeight     =   1740
         ScaleWidth      =   2415
         TabIndex        =   1
         Top             =   375
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   4875
      TabIndex        =   4
      Top             =   3750
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'For scrollbars large change value
Private Const SCROLL_PERCENT = 0.1

'For arrow keys
Private Const MOVE_SPEED = 100

Private tmpLeft As Integer
Private tmpTop As Integer

Private LastMouseX As Single  'self-explanatory
Private LastMouseY As Single  '           "

Private XDiff As Integer      'width difference
Private YDiff As Integer      'height difference
      
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
      Case vbKeyLeft
            tmpLeft = tmpLeft + MOVE_SPEED
            If tmpLeft >= 0 Then tmpLeft = 0
            picMove.Left = tmpLeft
            hsb.Value = Abs(tmpLeft)
            
      Case vbKeyRight
            tmpLeft = tmpLeft - MOVE_SPEED
            If tmpLeft <= -XDiff Then tmpLeft = -XDiff
            picMove.Left = tmpLeft
            hsb.Value = Abs(tmpLeft)
            
      Case vbKeyUp
            tmpTop = tmpTop + MOVE_SPEED
            If tmpTop >= 0 Then tmpTop = 0
            picMove.Top = tmpTop
            vsb.Value = Abs(tmpTop)
            
      Case vbKeyDown
            tmpTop = tmpTop - MOVE_SPEED
            If tmpTop <= -YDiff Then tmpTop = -YDiff
            picMove.Top = tmpTop
            vsb.Value = Abs(tmpTop)
            
      End Select
      
End Sub

Private Sub Form_Load()
      Me.KeyPreview = True
      
      picMove.AutoSize = True
      picMove.Move 0, 0
                  
      'Note: The picMove should be larger than picCont,
      'otherwise, there is no point in moving the picMove
      'because you can see the whole picture
      YDiff = Abs(picMove.Height - picCont.Height)
      XDiff = Abs(picMove.Width - picCont.Width)
      
      vsb.Max = YDiff
      vsb.LargeChange = YDiff * SCROLL_PERCENT
      
      hsb.Max = XDiff
      hsb.LargeChange = XDiff * SCROLL_PERCENT
      
End Sub

Private Sub picMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Button = vbLeftButton Then
            LastMouseX = X
            LastMouseY = Y
      End If
End Sub

Private Sub picMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Button <> vbLeftButton Then Exit Sub
       
      tmpLeft = picMove.Left + (X - LastMouseX)
      tmpTop = picMove.Top + (Y - LastMouseY)
      
      'Make sure picMove doesn't go out of boundaries
      If tmpLeft >= 0 Then tmpLeft = 0
      If tmpLeft <= -XDiff Then tmpLeft = -XDiff
      
      If tmpTop >= 0 Then tmpTop = 0
      If tmpTop <= -YDiff Then tmpTop = -YDiff
      
      picMove.Left = tmpLeft
      picMove.Top = tmpTop
      
      'Change the scrollbar values
      'note the Abs
      hsb.Value = Abs(tmpLeft)
      vsb.Value = Abs(tmpTop)
End Sub

Private Sub vsb_Change()
      picMove.Top = (-vsb.Value)
      tmpTop = picMove.Top
End Sub

Private Sub vsb_Scroll()
      vsb_Change
End Sub

Private Sub hsb_Change()
      picMove.Left = (-hsb.Value)
      tmpLeft = picMove.Left
End Sub

Private Sub hsb_Scroll()
      hsb_Change
End Sub
