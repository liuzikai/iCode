VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVBAPIHooker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***********************************************************************************************************************
'
'模 块 名: CVBAPIHooker
'
'作    者: Supergreenbean
'
'功能描述: 拦截指定进程中的特定API
'
'日    期: 2005年03月04日
'
'使用说明:
'       1.设定目标进程
'           例:
'               假设目标进程ID为&HC80
'
'               ※设定目标进程
'               m_oAPIHooker.SetTargetTo &HC80
'
'               注:如果参数为-1,表示把目标进程设为当前进程
'
'       2.拦截目标进程中的API函数
'           例:
'               拦截目标进程中notepad.exe模块从comdlg32.dll导入的函数GetOpenFileNameW
'
'               ※设定拦截
'               m_oAPIHooker.Hook "notepad.exe", "comdlg32.dll", "GetOpenFileNameW"
'
'       3.由字符串指针获得字符串
'***********************************************************************************************************************
'
Option Explicit
                                                                 
'本模块名称
Private Const THIS_MODULE_NAME     As String = "CVBAPIHooker"
'=======================================================================================================================
'以下为类的事件
'=======================================================================================================================
'事件:调用API前
'sAPIName   - 所拦截API函数的名称
'lpArgs     - 参数堆栈的指针(地址)
'fBlock - 是否阻止所拦截API函数执行，如果事件中此参数被设为True，那么被拦截的API将不会被执行
'lArgCount  - 被拦截API的参数个数，fBlock为True时必须同时设置这个参数
Public Event BeforeAction(ByVal sAPIName As String, ByVal lpArgs As Long, fBlock As Boolean, lArgCount As Long)
'事件:调用API后
'sAPIName   - 所拦截API函数的名称
'lRetValue  - 调用API函数后的函数返回值，在事件中可修改此参数
Public Event AfterAction(ByVal sAPIName As String, ByVal fBlocked As Boolean, lRetValue As Long)
'=======================================================================================================================
'以下为类的事件
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为Windows API相关声明
'=======================================================================================================================
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Const PAGE_READWRITE       As Long = &H4
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function RegisterWindowMessage Lib "User32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Sub DebugBreak Lib "kernel32" ()
'=======================================================================================================================
'以上为Windows API相关声明
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为类内部私有变量
'=======================================================================================================================
Private Const MSGTYPE_BEFORE_CALL  As Long = &H11
Private Const MSGTYPE_AFTER_CALL   As Long = &H22
Private Const MY_MSG               As String = "jAPI_HOOK_MSG"
'lpStart是目标进程中所分配内存的起始地址
Private Type FUNCTION_TABLE
'lpStart + 0
    lpGetModuleHandle              As Long
    'lpStart + 4
    lpLoadLibrary                  As Long
    'lpStart + 8
    lpGetProcAddress               As Long
    'lpStart + C
    lpPostMessage                  As Long
    'lpStart + 10
    lpOriginal                     As Long
End Type
Private Type VARIABLES
    FunctionTable                  As FUNCTION_TABLE
    'lpStart + 14
    ESP                            As Long
    'lpStart + 18
    HookIndex                      As Long
    'lpStart + 1C
    Flag1                          As Long                                      'break flag
    'lpStart + 20
    Flag2                          As Long                                      'Block flag
    'lpStart + 24
    hWnd                           As Long
    'lpStart + 28
    Msgid                          As Long
    'lpStart + 2C
    Retval                         As Long
    'lpStart + 30
    szUser32(6)                    As Byte
    'lpStart + 37
    szPostMessageA(12)             As Byte
    'lpStart + 44
    ReturnAddress                  As Long
    'lpStart + 48
    lpStartAddress                 As Long
    'lpStart + 4C
    lpCodeStartAddress             As Long
    'lpStart + 50
    ArgCount                       As Long
    'lpStart + 54
    abCodeBuffer(255)              As Byte
    'lpStart + 153
    'HOOKSTUBs
End Type
'可安装的钩子总数
Private Const MAX_HOOK_COUNT       As Long = &H255&
'小跳转代码段的最大大小
Private Const TINY_STUB_SIZE       As Long = &H20&
'变量
Private m_tVariables               As VARIABLES
'注册的消息ID
Private m_lMsgID                   As Long
'目标进程的进程ID
Private m_lProcessID               As Long
'目标进程的进程句柄
Private m_lhProcess                As Long
'目标进程是否为自身进程
Private m_fCurrentProcess          As Boolean
'目标进程中所分配内存的起始地址
Private m_lpStartAddress           As Long
'目标进程中代码段的起始地址
Private m_lpCodeStartAddress       As Long
'已安装钩子数
Private m_lHookCount               As Long
'是否已准备好拦截API
Private m_fIsReady                 As Boolean
'所拦截API的名字
Private m_asAPIName()              As String
'函数定位对象
Private m_aoPEFnLocator()          As CVBPEFnLocator
'内存分配管理对象
Private m_oMemAllocator            As CVBMemoryAllocator
'迷你消息窗口对象
Private WithEvents m_oMsgWindow    As CVBMiniWindow
Attribute m_oMsgWindow.VB_VarHelpID = -1
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
Public Sub CallProc(ByVal CallType As Long)
    m_oMsgWindow_IncomingMsg 0, m_lMsgID, CallType, 0
End Sub

'*********************************************
'*********************************************
'=======================================================================================================================
'以上为类属性
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为类公有函数
'=======================================================================================================================
'拦截函数
Public Function Hook(ByVal vModuleHandleOrName As Variant, ByVal sLibName As String, ByVal vFunction As Variant) As Long
    On Error GoTo Error_Handler
    Dim lHookIndex As Long, lpStubAddress As Long
    Hook = -1
    If m_lHookCount < MAX_HOOK_COUNT Then
        lHookIndex = GetFreeFnLocatorIndex()
        Set m_aoPEFnLocator(lHookIndex) = New CVBPEFnLocator
        With m_aoPEFnLocator(lHookIndex)
            .SetTargetTo m_lProcessID
            If .LocateImportFunction(vModuleHandleOrName, sLibName, vFunction) Then
                If PrepareTinyStub(lHookIndex, lpStubAddress) Then
                    If .ImpReplace(lpStubAddress) Then
                        m_asAPIName(lHookIndex) = CStr(vFunction)
                        m_lHookCount = m_lHookCount + 1
                        Hook = lHookIndex
                    End If
                End If
            Else
                Set m_aoPEFnLocator(lHookIndex) = Nothing
            End If
        End With
    End If
    Exit Function
Error_Handler:
End Function

'撤消拦截
Public Function Unhook(ByVal lHookIndex) As Boolean
    If lHookIndex >= 0 And lHookIndex <= m_lHookCount Then
        Set m_aoPEFnLocator(lHookIndex) = Nothing
        Unhook = True
    End If
End Function

'以单字节方式获得堆栈内容
Public Function DumpStack(abDumpData() As Byte, Optional ByVal lDumpLength As Long = 64) As Boolean
    If lDumpLength > 0 Then
        ReDim abDumpData(lDumpLength - 1)
        DumpStack = memcpy2local(VarPtr(abDumpData(0)), m_tVariables.ESP + 4, lDumpLength)
    End If
End Function

'以四字节方式获得堆栈内容
Public Function GetArguments(alArgs() As Long, Optional ByVal lArgCount As Long = 4) As Boolean
    If lArgCount > 0 Then
        ReDim alArgs(lArgCount - 1)
        GetArguments = memcpy2local(VarPtr(alArgs(0)), m_tVariables.ESP + 4, lArgCount * 4)
    End If
End Function

'设定目标进程
Public Function SetTargetTo(Optional ByVal pid As Long = -1) As Boolean
    If pid <= 0 And pid <> -1 Then Exit Function
    If m_lProcessID <> 0 Then
        Call EndWork
    End If
    If pid = -1 Then
        m_lProcessID = GetCurrentProcessId
        m_fCurrentProcess = True
    Else
        m_lProcessID = pid
        m_fCurrentProcess = False
    End If
    m_lhProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, m_lProcessID)
    SetTargetTo = StartWork()
    m_fIsReady = SetTargetTo
End Function

'=======================================================================================================================
'以上为类公共函数
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为内部函数
'=======================================================================================================================
'获得未使用的函数定位对象
Private Function GetFreeFnLocatorIndex() As Long
    Dim i As Long
    For i = 0 To MAX_HOOK_COUNT - 1
        If m_aoPEFnLocator(i) Is Nothing Then
            GetFreeFnLocatorIndex = i
            Exit Function
        End If
    Next
End Function

'拦截前的准备工作
Private Function StartWork() As Boolean
    Set m_oMemAllocator = New CVBMemoryAllocator
    m_oMemAllocator.Allocate Len(m_tVariables) + MAX_HOOK_COUNT * TINY_STUB_SIZE, m_lpStartAddress, m_lProcessID
    If m_lpStartAddress <> 0 Then
        m_lpCodeStartAddress = m_lpStartAddress + &H54
        m_tVariables.lpStartAddress = m_lpStartAddress
        m_tVariables.lpCodeStartAddress = m_lpCodeStartAddress
        StartWork = PrepareStub
    End If
End Function

'撤消拦截要做的工作
Private Function EndWork()
    m_fIsReady = False
    'clean hooks
    If m_lHookCount >= 0 Then
        Dim i As Long
        For i = 0 To MAX_HOOK_COUNT - 1
            Set m_aoPEFnLocator(i) = Nothing
        Next
    End If
    Set m_oMemAllocator = Nothing
    '关闭打开的进程句柄
    CloseHandle m_lhProcess
End Function

'准备总处理代码段
Private Function PrepareStub() As Boolean
    On Error GoTo Error_Handler
    Dim sCodes As String
    sCodes = ""
    'int 3 测试用
    'sCodes = "CC "
    Dim lpThis As Long
    Dim lpVTable As Long
    lpThis = ObjPtr(Me)
    CopyMemory lpVTable, ByVal lpThis, 4
    sCodes = sCodes & "58 55 "
    sCodes = sCodes & "BD " & LongToByteCodeStr(m_lpStartAddress)
    sCodes = sCodes & "89 45 44 "
    sCodes = sCodes & "83 7D 0C 00 75 1E 8D 7D 30 57 FF 55 00 83 F8 00 75 07 8D 7D 30 57 FF 55 04 "
    sCodes = sCodes & "8D 7D 37 57 50 FF 55 08 89 45 0C 89 65 14 "
    If m_fCurrentProcess Then
        sCodes = sCodes & "6A 11 68 " & LongToByteCodeStr(lpThis) & "FF 15 " & LongToByteCodeStr(lpVTable + &H1C)
    Else
        sCodes = sCodes & "6A 00 6A 11 FF 75 28 FF 75 24 FF 55 0C C6 45 1C 01 80 7D 1C 01 74 FA "
    End If
    sCodes = sCodes & "33 C0 80 7D 20 01 74 11 5D "
    sCodes = sCodes & "BF " & LongToByteCodeStr(m_lpStartAddress)
    sCodes = sCodes & "FF 57 10 55 "
    sCodes = sCodes & "BD " & LongToByteCodeStr(m_lpStartAddress)
    sCodes = sCodes & "EB 0D 8B 7D 50 8B 04 24 8D 64 BC 04 50 33 C0 "
    sCodes = sCodes & "8D 7D 2C 89 07 "
    If m_fCurrentProcess Then
        sCodes = sCodes & "6A 22 68 " & LongToByteCodeStr(lpThis) & "FF 15 " & LongToByteCodeStr(lpVTable + &H1C)
    Else
        sCodes = sCodes & "6A 00 6A 22 FF 75 28 FF 75 24 FF 55 0C C6 45 1C 01 80 7D 1C 01 74 FA "
    End If
    sCodes = sCodes & "8B 45 2C 5D "
    sCodes = sCodes & "BF " & LongToByteCodeStr(m_lpStartAddress)
    sCodes = sCodes & "FF 77 44 C3"
    Dim abCodes() As Byte
    abCodes = ByteCodeStrToBin(sCodes)
    If UBound(abCodes) <= UBound(m_tVariables.abCodeBuffer) Then
        CopyMemory m_tVariables.abCodeBuffer(0), abCodes(0), UBound(abCodes) + 1
        PrepareStub = memcpy2remote(m_lpStartAddress, VarPtr(m_tVariables), Len(m_tVariables))
    Else
        PrepareStub = False
    End If
    Exit Function
Error_Handler:
End Function

'准备小跳转代码段
Private Function PrepareTinyStub(ByVal lHookIndex As Long, lpStubAddress As Long) As Boolean
    Dim sCodes As String
    sCodes = ""
    'sCodes = "CC "
    sCodes = sCodes & "B8 " & LongToByteCodeStr(m_lpStartAddress)
    sCodes = sCodes & "C7 40 18 " & LongToByteCodeStr(lHookIndex)
    sCodes = sCodes & "FF 25 " & LongToByteCodeStr(m_lpStartAddress + &H4C)
    lpStubAddress = m_lpStartAddress + Len(m_tVariables) + lHookIndex * TINY_STUB_SIZE
    Dim abCodes() As Byte
    abCodes = ByteCodeStrToBin(sCodes)
    If UBound(abCodes) + 1 <= TINY_STUB_SIZE Then
        PrepareTinyStub = memcpy2remote(lpStubAddress, VarPtr(abCodes(0)), UBound(abCodes) + 1)
    End If
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

'把其它进程内存内容复制到到本地内存
Private Function memcpy2local(ByVal lpDestAddr As Long, ByVal lpSrcAddr As Long, ByVal lLen As Long) As Boolean
    On Error GoTo Error_Handler
    Dim lOldProtect As Long
    Dim lBytesReadWrite As Long
    If m_lhProcess Then
        If VirtualProtectEx(m_lhProcess, ByVal lpSrcAddr, lLen, PAGE_READWRITE, lOldProtect) <> 0 Then
            ReadProcessMemory m_lhProcess, ByVal lpSrcAddr, ByVal lpDestAddr, lLen, lBytesReadWrite
            Call VirtualProtectEx(m_lhProcess, ByVal lpSrcAddr, lLen, lOldProtect, lOldProtect)
            memcpy2local = (lBytesReadWrite <> 0)
        End If
    End If
    Exit Function
Error_Handler:
End Function

'把本地内存内容复制到到其它进程内存
Private Function memcpy2remote(ByVal lpDestAddr As Long, ByVal lpSrcAddr As Long, ByVal lLen As Long) As Boolean
    On Error GoTo Error_Handler
    Dim lOldProtect As Long
    Dim lBytesReadWrite As Long
    If m_lhProcess Then
        If VirtualProtectEx(m_lhProcess, ByVal lpDestAddr, lLen, PAGE_READWRITE, lOldProtect) <> 0 Then
            WriteProcessMemory m_lhProcess, ByVal lpDestAddr, ByVal lpSrcAddr, lLen, lBytesReadWrite
            Call VirtualProtectEx(m_lhProcess, ByVal lpDestAddr, lLen, lOldProtect, lOldProtect)
            memcpy2remote = (lBytesReadWrite <> 0)
        End If
    End If
    Exit Function
Error_Handler:
End Function

'收到消息
Private Sub m_oMsgWindow_IncomingMsg(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    If uMsg = m_lMsgID Then
        Dim lHookIndex As Long
        If memcpy2local(VarPtr(lHookIndex), m_lpStartAddress + &H18, 4) Then
            Select Case wParam
            Case MSGTYPE_BEFORE_CALL
                Dim lpOriFn As Long
                Dim ESP As Long
                Dim fBlock As Boolean
                Dim lArgCount As Long
                lpOriFn = m_aoPEFnLocator(lHookIndex).ImpJumpAddress
                memcpy2remote m_lpStartAddress + &H10, VarPtr(lpOriFn), 4
                Call memcpy2local(VarPtr(ESP), m_lpStartAddress + &H14, 4)
                m_tVariables.ESP = ESP
                RaiseEvent BeforeAction(m_asAPIName(lHookIndex), ESP + 4, fBlock, lArgCount)
                SetBlockFlag fBlock
                If fBlock Then
                    SetArgCount lArgCount
                End If
                Call ClearBreakFlag
            Case MSGTYPE_AFTER_CALL
                Dim lRetValue As Long
                memcpy2local VarPtr(lRetValue), m_lpStartAddress + &H2C, 4
                RaiseEvent AfterAction(m_asAPIName(lHookIndex), GetBlockFlag, lRetValue)
                SetReturnValue lRetValue
                Call ClearBreakFlag
            End Select
        End If
    End If
End Sub

'写入返回值
Private Sub SetReturnValue(ByVal lRetValue As Long)
    memcpy2remote m_lpStartAddress + &H2C, VarPtr(lRetValue), 4
End Sub

'设定参数个数
Private Sub SetArgCount(ByVal lArgCount As Long)
    memcpy2remote m_lpStartAddress + &H50, VarPtr(lArgCount), 4
End Sub

'清除中断标志
Private Sub ClearBreakFlag()
    Dim breakflag As Long
    If Not m_fCurrentProcess Then
        Do
            memcpy2local VarPtr(breakflag), m_lpStartAddress + &H1C, 1
            If breakflag = 1 Then Exit Do
            DoEvents
        Loop
        memcpy2remote m_lpStartAddress + &H1C, VarPtr(0), 1
    End If
End Sub

'设定阻止标志
Private Sub SetBlockFlag(ByVal fBlock As Boolean)
    Dim Blockflag As Long
    Blockflag = IIf(fBlock, 1, 0)
    memcpy2remote m_lpStartAddress + &H20, VarPtr(Blockflag), 4
End Sub

'获得阻止标志
Private Function GetBlockFlag() As Boolean
    Dim Blockflag As Long
    memcpy2local VarPtr(Blockflag), m_lpStartAddress + &H20, 4
    GetBlockFlag = (Blockflag <> 0)
End Function

'=======================================================================================================================
'以上为内部函数
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为类初始化与销毁过程
'=======================================================================================================================
'类初始化
Private Sub Class_Initialize()
    m_lMsgID = RegisterWindowMessage(MY_MSG)
    Set m_oMsgWindow = New CVBMiniWindow
    If m_oMsgWindow.Create() Then
        ReDim m_aoPEFnLocator(MAX_HOOK_COUNT - 1)
        ReDim m_asAPIName(MAX_HOOK_COUNT - 1)
        m_lHookCount = -1
        m_lProcessID = 0
        With m_tVariables
            .FunctionTable.lpGetModuleHandle = GetProcAddress(GetModuleHandle("kernel32"), "GetModuleHandleA")
            .FunctionTable.lpLoadLibrary = GetProcAddress(GetModuleHandle("kernel32"), "LoadLibraryA")
            .FunctionTable.lpGetProcAddress = GetProcAddress(GetModuleHandle("kernel32"), "GetProcAddress")
            .Msgid = m_lMsgID
            .hWnd = m_oMsgWindow.hMsgWindow
            Dim tmp() As Byte
            tmp = StrConv("User32", vbFromUnicode)
            CopyMemory .szUser32(0), tmp(0), 6
            tmp = StrConv("PostMessageA", vbFromUnicode)
            CopyMemory .szPostMessageA(0), tmp(0), 12
        End With
    End If
End Sub

'类销毁
Private Sub Class_Terminate()
    Call EndWork
    Set m_oMsgWindow = Nothing
End Sub

'=======================================================================================================================
'以上为类初始化与销毁过程
'=======================================================================================================================

'由字符串指针获得字符串
Public Function iGetStrFromPtr(ByVal lpStr As Long) As String
    Dim lHookIndex As Long
    memcpy2local VarPtr(lHookIndex), m_lpStartAddress + &H18, 4
    iGetStrFromPtr = m_aoPEFnLocator(lHookIndex).RemoteStrFromPtr(lpStr)
End Function

