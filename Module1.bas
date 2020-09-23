Attribute VB_Name = "Module1"
Option Explicit

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Const WM_DRAWCLIPBOARD = &H308
Public Const GWL_WNDPROC = (-4)
Public Const BM_SETSTATE = &HF3
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Dim PrevProc As Long

Public Sub HookForm(Button As Long)
    PrevProc = SetWindowLong(Button, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnHookForm(Button As Long)
    SetWindowLong Button, GWL_WNDPROC, PrevProc
End Sub
Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        
        
    If uMsg = WM_LBUTTONUP Or uMsg = 562 Then 'mouse button was released
        Form1.Stick
    End If
              
    If uMsg = 522 Then   'scroll.. wparam is a value for up or down scroll
        Form1.scrollmove (wParam)
    End If
        
    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
   
End Function

