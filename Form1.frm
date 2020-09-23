VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    HookForm Me.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHookForm Me.hwnd
End Sub

Public Sub Stick()
    'the first if statment prevents an error when the form is maximized _
    and and attmempt to move the form occurs.
    If Me.WindowState = vbMaximized Then Exit Sub

        If Me.Left <= 1400 Then 'move me.left to 0
            Me.Left = 0
        ElseIf Me.Left >= (Screen.Width - Me.Width) - 1400 Then
            'move the right edge of the form to the edge of the screen
            Me.Left = Screen.Width - Me.Width
        End If
    
        If Me.Top <= 1400 Then Me.Top = 0 'stick for to top
    
        If (Screen.Height - Me.Top) - 1400 <= Me.Height Then 'stick to bottom
            Me.Top = Screen.Height - Me.Height
    End If


End Sub

Public Sub scrollmove(wParam As Long)
 '***limits the upward movement of the form to the top of the screen ***
    
    'the first if statment prevents an error when the form is maximized _
    and and attmempt to move the form occurs.
    If Me.WindowState = vbMaximized Then Exit Sub
    
    If wParam < 0 Then 'the scroller was moved down
        If Me.Top <= 60 Then
            Me.Top = 0
            Exit Sub
        Else 'the scroller was moved up
            Me.Top = Me.Top - 200
        End If
        
    Else 'limit the bottom of the form to the bottom of the screen
        If (Screen.Height - Me.Top) <= Me.Height Then
            Me.Top = Screen.Height - Me.Height
        Else
            Me.Top = Me.Top + 200
        End If
    End If
  
End Sub
