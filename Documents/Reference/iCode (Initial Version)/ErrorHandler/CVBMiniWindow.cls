VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVBMiniWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***********************************************************************************************************************
'
'模 块 名: CVBMiniWindow
'
'作    者: Supergreenbean
'
'功能描述: 创建一个专门负责接受的消息不可视窗口
'
'日    期: 2005年03月03日
'
'使用说明:
'       1.创建窗口
'           例:
'               ※创建消息窗口
'               m_oMsgWindow.Create
'***********************************************************************************************************************
'
Option Explicit
                                                                 
'本模块名称
Private Const THIS_MODULE_NAME    As String = "CVBMiniWindow"
'=======================================================================================================================
'以下为类的事件
'=======================================================================================================================
'事件:收到消息
'参数意义同 WindowProc
Public Event IncomingMsg(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'=======================================================================================================================
'以下为类的事件
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为Windows API相关声明
'=======================================================================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Long, ByVal Length As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_WNDPROC = (-4&)
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
'=======================================================================================================================
'以上为Windows API相关声明
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为类内部私有变量
'=======================================================================================================================
'窗口句柄
Private m_hMsgWnd                 As Long
'默认窗口过程地址
Private m_lDefWindowProc          As Long
'API函数CallWindowProc的地址
Private m_lCallWindowProcAddress  As Long
'窗口回调过程代码
Private m_abWndProc()             As Byte
'=======================================================================================================================
'以上为类内部私有变量
'=======================================================================================================================

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为类属性
'=======================================================================================================================
'*********************************************
'*********************************************
'接收消息的窗口过程。本过程必须放在所有方法和属性前面
Public Sub WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    RaiseEvent IncomingMsg(hWnd, uMsg, wParam, lParam)
End Sub
                                                                         
'*********************************************
'*********************************************
'窗口句柄
Public Property Get hMsgWindow() As Long
    hMsgWindow = m_hMsgWnd
End Property

'=======================================================================================================================
'以上为类属性
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为类公有函数
'=======================================================================================================================
Public Function Create() As Boolean
    m_hMsgWnd = CreateWindowEx(&H80, "#32770", "MsgWindow", &H80000, 0, 0, 1, 1, 0, 0, 0, ByVal 0)
    Create = (m_hMsgWnd <> 0)
    If Create Then
        If PrepareCode Then
            Subclassing m_hMsgWnd
        End If
    End If
End Function

'=======================================================================================================================
'以上为类公共函数
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为内部函数
'=======================================================================================================================
'准备代码
Private Function PrepareCode() As Boolean
    On Error GoTo Error_Handler
    Dim lpThis As Long
    Dim lpVTable As Long
    lpThis = ObjPtr(Me)
    CopyMemory lpVTable, ByVal lpThis, 4
    Dim sCodes As String
    sCodes = ""
    'int 3
    'sCodes = "CC "
    '
    sCodes = sCodes & "55 8B EC 83 EC 10 8B FC 8D 75 08 B9 04 00 00 00 FC F3 A5 "
    'push lpThis
    sCodes = sCodes & "68 " & LongToByteCodeStr(lpThis)
    'call WndProc
    sCodes = sCodes & "FF 15 " & LongToByteCodeStr(lpVTable + &H1C)
    '
    sCodes = sCodes & "8B E5 5D "
    sCodes = sCodes & "55 8B EC 83 EC 10 8B FC 8D 75 08 B9 04 00 00 00 FC F3 A5 "
    'push m_lDefWindowProc
    sCodes = sCodes & "68 " & LongToByteCodeStr(GetWindowLong(m_hMsgWnd, GWL_WNDPROC))
    'call CallWindowProcA
    sCodes = sCodes & "FF 15 " & LongToByteCodeStr(VarPtr(m_lCallWindowProcAddress))
    '
    sCodes = sCodes & "8B E5 5D C3"
    m_abWndProc = ByteCodeStrToBin(sCodes)
    PrepareCode = True
    Exit Function
Error_Handler:
End Function

'16进制字符串变为字节数组
Private Function ByteCodeStrToBin(ByVal sByteCode As String) As Byte()
    Dim s() As String
    Dim b() As Byte
    Dim i As Long
    s = Split(Trim(sByteCode), " ")
    If UBound(s) >= 0 Then
        ReDim b(UBound(s))
    End If
    For i = 0 To UBound(s)
        b(i) = CByte("&h" & s(i))
    Next
    ByteCodeStrToBin = b
End Function

'长整型变为16进制字符串
Private Function LongToByteCodeStr(ByVal lValue As Long) As String
    Dim s As String
    s = Right$("00000000" & Hex(lValue), 8)
    LongToByteCodeStr = Mid$(s, 1, 2) & " " & LongToByteCodeStr
    LongToByteCodeStr = Mid$(s, 3, 2) & " " & LongToByteCodeStr
    LongToByteCodeStr = Mid$(s, 5, 2) & " " & LongToByteCodeStr
    LongToByteCodeStr = Mid$(s, 7, 2) & " " & LongToByteCodeStr
End Function

'子类处理窗口 hWndTarget
Private Sub Subclassing(ByVal hWndTarget As Long, Optional ByVal fUnsubclassing As Boolean = False)
    On Error Resume Next
    If hWndTarget = 0 Then Exit Sub
    If fUnsubclassing Then
        If m_lDefWindowProc <> 0 Then
            SetWindowLong hWndTarget, GWL_WNDPROC, m_lDefWindowProc
        End If
    Else
        m_lDefWindowProc = SetWindowLong(hWndTarget, GWL_WNDPROC, VarPtr(m_abWndProc(0)))
    End If
End Sub

'=======================================================================================================================
'以上为内部函数
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为类初始化与销毁过程
'=======================================================================================================================
'类初始化
Private Sub Class_Initialize()
    m_hMsgWnd = 0
    m_lCallWindowProcAddress = GetProcAddress(GetModuleHandle("user32"), "CallWindowProcA")
End Sub

'类销毁
Private Sub Class_Terminate()
    Subclassing m_hMsgWnd, True
    DestroyWindow m_hMsgWnd
End Sub
                                                                         
'=======================================================================================================================
'以上为类初始化与销毁过程
'=======================================================================================================================
