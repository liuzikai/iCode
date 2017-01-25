Attribute VB_Name = "Module1"
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_SEPARATOR = &H800&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYCOMMAND = &H0&
Public Const GWL_WNDPROC = (-4)
Public Const WM_COMMAND = &H111
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function AppendMenu1 Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public MenuCount  As Long    '菜单数量，不包括不能触发的菜单
Public MenuText() As String  '菜单文本，ID=wParam的菜单的文本为MenuText(wParam - 1000)
Public OldWinProc As Long

Public Function OnMenu(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '{响应菜单事件}
    Select Case wMsg
        Case WM_COMMAND
            If wParam > 1000 And wParam <= 1000 + MenuCount Then
                MsgBox MenuText(wParam - 1000)
            End If
    End Select
    OnMenu = CallWindowProc(OldWinProc, hwnd, wMsg, wParam, lParam)
End Function
