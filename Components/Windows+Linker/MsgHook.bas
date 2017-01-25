Attribute VB_Name = "MsgHook"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type


Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    Message As Long
    hWnd As Long
End Type

Private Const WH_GETMESSAGE = 3
Private Const WH_CALLWNDPROC = 4
Private Const WH_CALLWNDPROCRET = 12

Private Const HC_ACTION = 0
Private Const PM_REMOVE = &H1

Private Const lGetMsg As Boolean = False
Private Const lCallWndProc As Boolean = True
Private Const lCallWndProcProtect As Boolean = False

Private lngGetMsgProc As Long, lngCallWndRetProc As Long, lngCallWndProc As Long

Private Const WM_CHILDACTIVATE = &H22


Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_NCCREATE = &H81
Public Const WM_SIZING = &H214
Public Const WM_MOVING = &H216
Public Const WM_SETTEXT = &HC

Public Const WM_PAINT = &HF
Public Const WM_SHOWWINDOW = &H18
Public Const WM_NCPAINT = &H85
Public Const WM_SETFOCUS = &H7

Public Const WM_KEYDOWN = &H100



Public Const BM_SETSTYLE = &HF4

Public Const CB_SETEXTENDEDUI = &H155
Public Const CB_ADDSTRING = &H143


Private Type CREATESTRUCT
        lpCreateParams As Long
        hInstance As Long
        hMenu As Long
        hWndParent As Long
        cy As Long
        cx As Long
        y As Long
        x As Long
        style As Long
        lpszName As Long
        lpszClass As Long
        ExStyle As Long
End Type

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public WND As clsWindowsHandler

'---Windows---
Public Const WM_CREATE = &H1
Public Const WM_ENABLE = &HA
Public Const WM_DESTROY = &H2


Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long


Public Sub iMsgProc(ByVal hWnd As Long, ByRef Msg As Long, ByRef wParam As Long, ByRef lParam As Long)
    
    'ÅÅ³ý£º&HD
    '      &H4E¡¢&H20
    
    Select Case Msg
    
    Case &H128
        WND.Msg_Windows hWnd
    Case WM_DESTROY
        WND.Msg_WM_DESTROY hWnd
    End Select
    
End Sub

Private Sub iMsgProtectProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
    
End Sub

Public Function Hook_GetMsgProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Hook_GetMsgProc = CallNextHookEx(lngGetMsgProc, nCode, wParam, lParam)
    
    If nCode = HC_ACTION And wParam = PM_REMOVE Then
        Dim P As Msg
        CopyMemory P, ByVal lParam, Len(P)
        
        Call iMsgProc(P.hWnd, P.Message, P.wParam, P.lParam)
    End If
    
End Function

Public Function Hook_CallWndProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Hook_CallWndProc = CallNextHookEx(lngCallWndProc, nCode, wParam, lParam)
    
    If nCode = HC_ACTION Then
        Dim P As CWPSTRUCT
        CopyMemory P, ByVal lParam, Len(P)
        
        Call iMsgProc(P.hWnd, P.Message, P.wParam, P.lParam)
        
    End If
    
End Function

Public Function Hook_CallWndRetProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Hook_CallWndRetProc = CallNextHookEx(lngCallWndProc, nCode, wParam, lParam)
    
    If nCode = HC_ACTION Then
        
        Dim P As CWPSTRUCT
        CopyMemory P, ByVal lParam, Len(P)
        
        iMsgProtectProc P.hWnd, P.Message, P.wParam, P.lParam
        
    End If
End Function

Public Sub SetMsgHooks()
    Dim hIns As Long, TID As Long
    
    hIns = 0
    TID = GetWindowThreadProcessId(hVBIDE, 0&)
    
    If lGetMsg Then
        lngGetMsgProc = SetWindowsHookEx(WH_GETMESSAGE, AddressOf Hook_GetMsgProc, hIns, TID)
        DBPrint "lngGetMsgProc = " & lngGetMsgProc
    End If
    
    If lCallWndProc Then
        lngCallWndProc = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf Hook_CallWndProc, hIns, TID)
        DBPrint "lngCallWndProc = " & lngCallWndProc
    End If
    
    If lCallWndProcProtect Then
        lngCallWndRetProc = SetWindowsHookEx(WH_CALLWNDPROCRET, AddressOf Hook_CallWndRetProc, hIns, TID)
        DBPrint "lngCallWndRetProc = " & lngCallWndRetProc
    End If
    
End Sub

Public Sub UnSetMsgHooks()
    If lGetMsg Then UnhookWindowsHookEx lngGetMsgProc
    If lCallWndProc Then UnhookWindowsHookEx lngCallWndProc
    If lCallWndProcProtect Then UnhookWindowsHookEx lngCallWndRetProc
End Sub

