VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iRemoteHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'模块名：iRemoteHook
'作者：liuzikai [liuzikai@163.com]
'日期：2015年7月17日
'版本：V1.0
'说明：执行远程Hook
'       1.注册需要拦截的消息：       .RegisterMessage(消息)
'       2.注入Hook：                .Inject(线程ID)
'       3.反注入Hook：              .UnInject
'       4.拦截消息的MiniWindow：    .MsgWindow
'       5.拦截到消息时触发：         Event GetCallWndMessage（参数均为消息内容，Result设置返回值）

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type CWPSTRUCT
        lParam As Long
        wParam As Long
        Message As Long
        hWnd As Long
End Type

Private Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As Long
End Type

Private Const WM_COPYDATA = &H4A

#If 1 Then

    Private Declare Function DLL_CALLWNDHOOK_RegisterMessage Lib "Public\iRemoteHook\iDLL_CALLWNDHOOK\Debug\iDll.dll" Alias "RegisterMessage" (ByVal Msg As Long, ByVal hReciver As Long) As Long
    Private Declare Function DLL_CALLWNDHOOK_SetHook Lib "Public\iRemoteHook\iDLL_CALLWNDHOOK\Debug\iDll.dll" Alias "SetHook" (ByVal TID As Long) As Long
    Private Declare Sub DLL_CALLWNDHOOK_UnHook Lib "Public\iRemoteHook\iDLL_CALLWNDHOOK\Debug\iDll.dll" Alias "UnHook" ()
    
    Private Declare Function DLL_GETMESSAGE_RegisterMessage Lib "Public\iRemoteHook\iDLL_GETMESSAGE\Debug\iDll.dll" Alias "RegisterMessage" (ByVal Msg As Long, ByVal hReciver As Long) As Long
    Private Declare Function DLL_GETMESSAGE_SetHook Lib "Public\iRemoteHook\iDLL_GETMESSAGE\Debug\iDll.dll" Alias "SetHook" (ByVal TID As Long) As Long
    Private Declare Sub DLL_GETMESSAGE_UnHook Lib "Public\iRemoteHook\iDLL_GETMESSAGE\Debug\iDll.dll" Alias "UnHook" ()

#Else

    Private Declare Function DLL_CALLWNDHOOK_RegisterMessage Lib "..\Public\iRemoteHook\iDLL_CALLWNDHOOK\Debug\iDll.dll" Alias "RegisterMessage" (ByVal Msg As Long, ByVal hReciver As Long) As Long
    Private Declare Function DLL_CALLWNDHOOK_SetHook Lib "..\Public\iRemoteHook\iDLL_CALLWNDHOOK\Debug\iDll.dll" Alias "SetHook" (ByVal TID As Long) As Long
    Private Declare Sub DLL_CALLWNDHOOK_UnHook Lib "..\Public\iRemoteHook\iDLL_CALLWNDHOOK\Debug\iDll.dll" Alias "UnHook" ()
    
    Private Declare Function DLL_GETMESSAGE_RegisterMessage Lib "..\Public\iRemoteHook\iDLL_GETMESSAGE\Debug\iDll.dll" Alias "RegisterMessage" (ByVal Msg As Long, ByVal hReciver As Long) As Long
    Private Declare Function DLL_GETMESSAGE_SetHook Lib "..\Public\iRemoteHook\iDLL_GETMESSAGE\Debug\iDll.dll" Alias "SetHook" (ByVal TID As Long) As Long
    Private Declare Sub DLL_GETMESSAGE_UnHook Lib "..\Public\iRemoteHook\iDLL_GETMESSAGE\Debug\iDll.dll" Alias "UnHook" ()

#End If

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Dim WithEvents MiniWindow As iMiniWindow
Attribute MiniWindow.VB_VarHelpID = -1

Event GetCallWndMessage(ByRef Result As Long, ByVal hWnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)

Public Enum iRemoteHook_Mode
    SubClass = 1
    CALLWNDPROCHOOK = 2
    GETMESSAGEHOOK = 3
End Enum

Dim m_Mode As iRemoteHook_Mode

Public Function GetThreadIDByhWnd(ByVal hWnd As Long) As Long
    GetThreadIDByhWnd = GetWindowThreadProcessId(hWnd, 0&)
End Function

Public Sub SetMode(ByVal Mode As iRemoteHook_Mode)
    m_Mode = Mode
End Sub

'为了保证运行效率，Dll中设置了消息过滤器，只有指定的消息会被传送过来。过滤器最多可以容纳10条消息（也是为了保证运行效率）。
'对于同一个消息，可以注册多个接收句柄（同一个信息上限为5个），Dll将会依次传送消息

'使用RegisterMessage注册所需消息
Public Function RegisterMessage(ByVal Msg As Long)
    If m_Mode = SubClass Then
    
    ElseIf m_Mode = CALLWNDPROCHOOK Then
        RegisterMessage = DLL_CALLWNDHOOK_RegisterMessage(Msg, MiniWindow.hWnd)
    ElseIf m_Mode = GETMESSAGEHOOK Then
        RegisterMessage = DLL_GETMESSAGE_RegisterMessage(Msg, MiniWindow.hWnd)
    Else
        DBPrint "iReomoteHook: Have not set Mode yet!"
    End If
End Function

'注入Hook到指定线程
Public Function Inject(ByVal ThreadID As Long) As Long
    If m_Mode = SubClass Then
    
    ElseIf m_Mode = CALLWNDPROCHOOK Then
        Inject = DLL_CALLWNDHOOK_SetHook(ThreadID)
    ElseIf m_Mode = GETMESSAGEHOOK Then
        Inject = DLL_GETMESSAGE_SetHook(ThreadID)
    Else
        DBPrint "iReomoteHook: Have not set Mode yet!"
    End If
End Function

'反注入Hook
Public Sub UnInject()
    If m_Mode = SubClass Then
    
    ElseIf m_Mode = CALLWNDPROCHOOK Then
        DLL_CALLWNDHOOK_UnHook
    ElseIf m_Mode = GETMESSAGEHOOK Then
        DLL_GETMESSAGE_UnHook
    End If
    
    MiniWindow.Terminate
    Set MiniWindow = Nothing
End Sub

Friend Property Get MsgWindow() As iMiniWindow
    Set MsgWindow = MiniWindow
End Property

Private Sub Class_Initialize()
    Set MiniWindow = New iMiniWindow
    MiniWindow.Initialize
End Sub

Private Sub Class_Terminate()
    If Not (MiniWindow Is Nothing) Then
        UnInject
    End If
End Sub

'Dll采用发送WM_COPYDATA消息的方式回传

Private Sub MiniWindow_GetWindowMessage(Result As Long, ByVal cHwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)
    If Message = WM_COPYDATA Then
        
        Dim cds As COPYDATASTRUCT
        CopyMemory cds, ByVal lParam, Len(cds)

        Dim cwp As CWPSTRUCT
        CopyMemory cwp, ByVal cds.lpData, Len(cwp)
        
        RaiseEvent GetCallWndMessage(Result, cwp.hWnd, cwp.Message, cwp.wParam, cwp.lParam)
        
    Else
        If Not (MiniWindow Is Nothing) Then Result = MiniWindow.CallDefaultWindowProc(cHwnd, Message, wParam, lParam)
    End If
End Sub

