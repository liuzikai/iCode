Attribute VB_Name = "WndHook"
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private Const WM_MOUSEWHEEL = &H20A
Private PrevWndProc As Long
Private PrevhWnd As Long

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    WndProc = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)
    
    If uMsg = WM_MOUSEWHEEL Then
        Call TBH.Msg_MouseWheel(wParam)
    End If
End Function

Public Function SetWndHook(ByVal hWnd As Long) As Long
    PrevhWnd = hWnd
    PrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc)
    SetWndHook = PrevWndProc
End Function

Public Sub UnSetWndHook()
    'Call SetWindowLong(PrevhWnd, GWL_WNDPROC, PrevWndProc)
End Sub
