Attribute VB_Name = "MsgHook"
Option Explicit

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

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

Private Const lGetMsg As Boolean = True
Private Const lCallWndProc As Boolean = False
Private Const lCallWndProcProtect As Boolean = False

Private lngGetMsgProc As Long, lngCallWndRetProc As Long, lngCallWndProc As Long

Public Const WM_KEYUP = &H101
Public Const WM_KEYDOWN = &H100
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_CHAR = &H102

Public Const WM_CLOSE = &H10


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

Public OBJclsAC As clsAutoComplete

Public Function iMsgProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case Msg
        
    Case WM_KEYUP
        iMsgProc = OBJclsAC.Msg_KEYUP(hWnd, wParam)
    Case WM_KEYDOWN
        iMsgProc = OBJclsAC.Msg_KEYDOWN(hWnd, wParam)
    Case WM_LBUTTONDOWN
        Call OBJclsAC.Msg_LBUTTONDOWN(hWnd)
    Case WM_CHAR
        Call OBJclsAC.Msg_Char(hWnd, wParam)
    Case WM_CLOSE
        Call OBJclsAC.Msg_CLOSE(hWnd)
    End Select
    
End Function

Private Sub iMsgProtectProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
    
End Sub

Public Function Hook_GetMsgProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim r As Long
    r = CallNextHookEx(lngGetMsgProc, nCode, wParam, lParam)
    
    If nCode = HC_ACTION And wParam = PM_REMOVE Then
        Dim P As Msg
        CopyMemory P, ByVal lParam, Len(P)
        
        If iMsgProc(P.hWnd, P.Message, P.wParam, P.lParam) = 1 Then
            P.Message = 0
            P.wParam = 0
            P.lParam = 0
            CopyMemory ByVal lParam, P, Len(P)
        End If
    End If
    
    Hook_GetMsgProc = r
End Function

Public Function Hook_CallWndProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim r As Long
    r = CallNextHookEx(lngCallWndProc, nCode, wParam, lParam)
    
    If nCode = HC_ACTION Then
        Dim P As CWPSTRUCT
        CopyMemory P, ByVal lParam, Len(P)
        
        If iMsgProc(P.hWnd, P.Message, P.wParam, P.lParam) = 1 Then
            P.Message = 0
            P.wParam = 0
            P.lParam = 0
            CopyMemory ByVal lParam, P, Len(P)
        End If
        
    End If
    
    Hook_CallWndProc = r
    
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
        
    If lGetMsg Then
        lngGetMsgProc = SetWindowsHookEx(WH_GETMESSAGE, AddressOf Hook_GetMsgProc, 0, GetCurrentThreadId)
        DBPrint "lngGetMsgProc = " & lngGetMsgProc
    End If
        
    If lCallWndProc Then
        lngCallWndProc = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf Hook_CallWndProc, 0, GetCurrentThreadId)
        DBPrint "lngCallWndProc = " & lngCallWndProc
    End If
        
    If lCallWndProcProtect Then
        lngCallWndRetProc = SetWindowsHookEx(WH_CALLWNDPROCRET, AddressOf Hook_CallWndRetProc, 0, GetCurrentThreadId)
        DBPrint "lngCallWndRetProc = " & lngCallWndRetProc
    End If
    
End Sub

Public Sub UnSetMsgHooks()
    If lGetMsg Then UnhookWindowsHookEx lngGetMsgProc
    If lCallWndProc Then UnhookWindowsHookEx lngCallWndProc
    If lCallWndProcProtect Then UnhookWindowsHookEx lngCallWndRetProc
End Sub

