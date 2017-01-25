Attribute VB_Name = "CBTHook"
Option Explicit

Public Const WH_CBT = 5

Public Const HCBT_ACTIVATE = 5
Public Const HCBT_CLICKSKIPPED = 6
Public Const HCBT_CREATEWND = 3
Public Const HCBT_DESTROYWND = 4
Public Const HCBT_KEYSKIPPED = 7
Public Const HCBT_MINMAX = 1
Public Const HCBT_MOVESIZE = 0
Public Const HCBT_QS = 2
Public Const HCBT_SETFOCUS = 9
Public Const HCBT_SYSCOMMAND = 8

Private Const WM_KEYDOWN = &H100
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_KEYUP = &H101
Private Const WM_SYSKEYUP = &H105


Public hCBTHook As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Const WM_SETTEXT = &HC
Private Const EM_SETSEL = &HB1
Private Const WM_LBUTTONUP = &H202


Public Function iHookFunc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    iHookFunc = CallNextHookEx(hCBTHook, nCode, wParam, lParam)
    
    On Error GoTo iErr
    
    Dim sClassName As String, sCaption As String
    
    sClassName = iGetClassName(wParam)
    sCaption = iGetCaption(wParam)
    
    
    Select Case nCode
                
    End Select
    
    Exit Function
    
iErr:
    
    DBErr "CBTHook - Function iHookFunc", "nCode = " & nCode, "wParam = " & wParam, "lParam = " & lParam
    
End Function


Public Sub SetCBTHook()
    Dim hIns As Long, TID As Long
    If CBool(App.LogMode) = True Then
        hIns = 0
        TID = GetCurrentThreadId
    Else
        hIns = App.hInstance
        TID = GetWindowThreadProcessId(hVBIDE, 0)
    End If
    
    hCBTHook = SetWindowsHookEx(WH_CBT, AddressOf iHookFunc, hIns, TID)
End Sub

Public Sub UnSetCBTHook()
    UnhookWindowsHookEx hCBTHook
End Sub
