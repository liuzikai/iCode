Attribute VB_Name = "MsgHook"
Option Explicit

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
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

Private Const WM_CHILDACTIVATE = &H22


Private Function iMsgProc(ByVal hWnd As Long, ByRef Msg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef ReturnValue As Long, ByVal Time As Long, ByRef pt As POINTAPI) As Boolean
    Dim sClassName As String, sCaption As String
    
    sClassName = iGetClassName(hWnd)
    sCaption = iGetCaption(hWnd)
    
    
    Select Case Msg
        
    Case &H52C
        
        Select Case sClassName
            
        Case "#32770"
            
            WindowsControler.Msg hWnd, sCaption
            
        End Select
        
    Case WM_KEYUP
        
        Select Case wParam
            
        Case vbKeyReturn
            
            iCode.CodeSort.DealMessage Msg, wParam, sCaption, sClassName
            
        Case vbKeyControl
            
            iCode.AC.DealMessage Msg, wParam, sCaption, sClassName
            
        Case vbKeyUp, vbKeyDown
            
            '这段代码因牵涉到两个模块，直接在MsgHook中处理
            If iState = isCode Then
                If iCode.AC.Visible = False Then
                    iCode.CodeSort.SortMouseKeyEvent
                Else
                    Msg = 0
                    iMsgProc = True
                End If
            End If
            
        End Select
        
    Case WM_CHAR
        
        iCode.AC.DealMessage Msg, wParam, sCaption, sClassName
        
    Case WM_KEYDOWN
        
        If sClassName = "VbaWindow" Then
            'If Not (iCode.ErrBox Is Nothing) Then If iCode.ErrBox.Visible = True Then iCode.ErrBox.Hide
        End If
        
        Select Case wParam
            
        Case vbKeyBack, vbKeyReturn, vbKeySpace, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
            
            If iCode.AC.DealMessage(Msg, wParam, sCaption, sClassName) = False Then
                Msg = 0
                iMsgProc = True
            End If
            
        Case vbKeyControl
            
            iCode.AC.DealMessage Msg, wParam, sCaption, sClassName
            
        Case vbKeyF2
            
            If iState = isCode Then WindowsControler.OV_SelText = CodeOpe.Selection
            
        Case Else
            
            iCode.AC.DealMessage Msg, wParam, sCaption, sClassName
            
        End Select
        
    Case WM_LBUTTONUP
        
        If sClassName = "VbaWindow" Then
            
            iCode.CodeSort.DealMessage Msg, wParam, sCaption, sClassName
            iCode.AC.DealMessage Msg, wParam, sCaption, sClassName
            'If Not (iCode.ErrBox Is Nothing) Then If iCode.ErrBox.Visible = True Then iCode.ErrBox.Hide
            
        End If
        
    End Select
    
End Function

Private Sub iMsgProtectProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
    
End Sub

Public Function Hook_GetMsgProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim r As Long
    r = CallNextHookEx(lngGetMsgProc, nCode, wParam, lParam)
    
    If nCode = HC_ACTION And wParam = PM_REMOVE Then
        Dim p As Msg
        CopyMemory p, ByVal lParam, Len(p)
        
        If iMsgProc(p.hWnd, p.message, p.wParam, p.lParam, r, p.Time, p.pt) = True Then CopyMemory ByVal lParam, p, Len(p)
    End If
    
    Hook_GetMsgProc = r
End Function

Public Function Hook_CallWndProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim r As Long
    r = CallNextHookEx(lngCallWndProc, nCode, wParam, lParam)
    
    If nCode = HC_ACTION Then
        Dim p As CWPSTRUCT
        CopyMemory p, ByVal lParam, Len(p)
        
        Dim EmptyPT As POINTAPI
        If iMsgProc(p.hWnd, p.message, p.wParam, p.lParam, r, 0, EmptyPT) = True Then CopyMemory ByVal lParam, p, Len(p)
        '这里可以使用API函数获取系统时间传入Time参数
        
    End If
    
    Hook_CallWndProc = r
    
End Function

Public Function Hook_CallWndRetProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Hook_CallWndRetProc = CallNextHookEx(lngCallWndProc, nCode, wParam, lParam)
    
    If nCode = HC_ACTION Then
        
        Dim p As CWPSTRUCT
        CopyMemory p, ByVal lParam, Len(p)
        
        iMsgProtectProc p.hWnd, p.message, p.wParam, p.lParam
        
    End If
End Function

Public Sub SetMsgHooks()
    Dim hIns As Long, TID As Long
    
    If CBool(App.LogMode) = True Then
        hIns = 0
        TID = GetCurrentThreadId
    Else
        hIns = App.hInstance
        TID = GetWindowThreadProcessId(hVBIDE, 0)
    End If
    
    If lGetMsg Then
        lngGetMsgProc = SetWindowsHookEx(WH_GETMESSAGE, AddressOf Hook_GetMsgProc, hIns, TID)
        'DBPrint "lngGetMsgProc = " & lngGetMsgProc
    End If
    
    If lCallWndProc Then
        lngCallWndProc = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf Hook_CallWndProc, hIns, TID)
        'DBPrint "lngCallWndProc = " & lngCallWndProc
    End If
    
    If lCallWndProcProtect Then
        lngCallWndRetProc = SetWindowsHookEx(WH_CALLWNDPROCRET, AddressOf Hook_CallWndRetProc, hIns, TID)
        'DBPrint "lngCallWndRetProc = " & lngCallWndRetProc
    End If
End Sub

Public Sub UnSetMsgHooks()
    If lGetMsg Then UnhookWindowsHookEx lngGetMsgProc
    If lCallWndProc Then UnhookWindowsHookEx lngCallWndProc
    If lCallWndProcProtect Then UnhookWindowsHookEx lngCallWndRetProc
End Sub

