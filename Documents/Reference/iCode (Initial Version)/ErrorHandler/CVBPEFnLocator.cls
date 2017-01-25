VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVBPEFnLocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***********************************************************************************************************************
'
'模 块 名: CVBPEFnLocator
'
'作    者: Supergreenbean
'
'功能描述: 获取、替换指定进程中特定模块的输入/输出函数地址
'
'日    期: 2005年03月02日
'
'使用说明:
'         1. 设定目标进程
'               例:
'                   ※设定目标进程为进程ID等于&H1234的进程
'                   oPEFnLocator.SetTargetTo &H1234
'
'                   注:如果不执行这步,将默认把目标进程设为当前进程
'
'         2. 定位特定模块中的特定输入/输出函数
'
'           a)根据模块句柄定位输入/输出函数
'               例:
'                   假定Shell32.dll的模块句柄为&H773A0000,
'
'                   ※如何定位shell32.dll中,从user32.dll里所导入的ShowCaret函数的地址
'                       oPEFnLocator.LocateImportFunction &H773A0000,"user32.dll","ShowCaret"
'
'                   ※如何定位shell32.dll中导出函数Control_RunDLL的地址
'                       oPEFnLocator.LocateExportFunction &H773A0000,"Control_RunDLL"
'
'           b)根据模块名称定位输入/输出函数
'               例:
'                   ※如何定位shell32.dll中,从user32.dll里所导入的ShowCaret函数的地址
'                       oPEFnLocator.LocateImportFunction "shell32.dll","user32.dll","ShowCaret"
'
'                   ※如何定位shell32.dll中导出函数Control_RunDLL的地址
'                       oPEFnLocator.LocateExportFunction "shell32.dll","Control_RunDLL"
'
'         3. 替换跳转地址
'               例:
'                   ※如何替换已定位的函数跳转地址为&H123456
'                       oPEFnLocator.ImpReplace &H123456
'         4. 恢复跳转地址
'               例:
'                   ※如何恢复已定位的函数原跳转地址
'                       oPEFnLocator.ImpRestore
'***********************************************************************************************************************
'
Option Explicit
                                                                 
'本模块名称
Private Const THIS_MODULE_NAME  As String = "CVBPEFnLocator"
'=======================================================================================================================
'以下为Windows API相关声明
'=======================================================================================================================
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Const PAGE_READWRITE    As Long = &H4
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Const TH32CS_SNAPMODULE = &H8
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, lpme As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, lpme As MODULEENTRY32) As Long
Private Type MODULEENTRY32
    dwSize                      As Long
    th32ModuleID                As Long
    th32ProcessID               As Long
    GlblcntUsage                As Long
    ProccntUsage                As Long
    modBaseAddr                 As Long
    modBaseSize                 As Long
    hModule                     As Long
    szModule                    As String * 256
    szExePath                   As String * 260
End Type
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, lpFilePart As Long) As Long
'=======================================================================================================================
'以上为Windows API相关声明
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为类内部私有变量
'=======================================================================================================================
'目标函数在函数输入表中的入口地址
Private m_lImpEntryAddress      As Long
'目标函数在函数输入表中的跳转地址
Private m_lImpJumpAddress       As Long
'目标模块文件名（输入）
Private m_sImpModuleName        As String
'目标函数在函数输出表中的入口地址
Private m_lExpEntryAddress      As Long
'目标函数在函数输出表中的跳转地址
Private m_lExpJumpAddress       As Long
'目标模块文件名（输出）
Private m_sExpModuleName        As String
'函数输入表跳转地址是否被替换过
Private m_fImpReplaced          As Boolean
'函数输出表跳转地址是否被替换过
Private m_fExpReplaced          As Boolean
'是否在类被销毁时自动恢复被替换地址
Private m_fAutoRestore          As Boolean
'目标进程的进程ID
Private m_lProcessID            As Long
'目标进程的进程句柄
Private m_lhProcess             As Long
'=======================================================================================================================
'以上为类内部私有变量
'=======================================================================================================================

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为类属性
'=======================================================================================================================
'目标模块文件名
Public Property Let AutoRestore(fAutoRestore As Boolean)
    m_fAutoRestore = fAutoRestore
End Property

Public Property Get AutoRestore() As Boolean
    AutoRestore = m_fAutoRestore
End Property

'目标模块文件名
Public Property Let ImpModuleName(sImpModuleName As String)
    m_sImpModuleName = sImpModuleName
End Property

Public Property Get ImpModuleName() As String
    ImpModuleName = m_sImpModuleName
End Property

'目标函数在函数输入表中的跳转地址
Public Property Let ImpJumpAddress(lImpJumpAddress As Long)
    m_lImpJumpAddress = lImpJumpAddress
End Property

Public Property Get ImpJumpAddress() As Long
    ImpJumpAddress = m_lImpJumpAddress
End Property

'目标函数在函数输入表中的入口地址
Public Property Let ImpEntryAddress(lImpEntryAddress As Long)
    m_lImpEntryAddress = lImpEntryAddress
End Property

Public Property Get ImpEntryAddress() As Long
    ImpEntryAddress = m_lImpEntryAddress
End Property

'目标模块文件名
Public Property Let ExpModuleName(sExpModuleName As String)
    m_sExpModuleName = sExpModuleName
End Property

Public Property Get ExpModuleName() As String
    ExpModuleName = m_sExpModuleName
End Property

'目标函数在函数输入表中的跳转地址
Public Property Let ExpJumpAddress(lExpJumpAddress As Long)
    m_lExpJumpAddress = lExpJumpAddress
End Property

Public Property Get ExpJumpAddress() As Long
    ExpJumpAddress = m_lExpJumpAddress
End Property

'目标函数在函数输入表中的入口地址
Public Property Let ExpEntryAddress(lExpEntryAddress As Long)
    m_lExpEntryAddress = lExpEntryAddress
End Property

Public Property Get ExpEntryAddress() As Long
    ExpEntryAddress = m_lExpEntryAddress
End Property

'=======================================================================================================================
'以上为类属性
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为类公有函数
'=======================================================================================================================
'替换函数跳转地址
Public Function ImpReplace(ByVal lAddress As Long, Optional ByVal lpTargetAddress As Long = 0) As Boolean
    On Error GoTo Error_Handler
    Dim lOldProtect As Long
    lpTargetAddress = IIf(lpTargetAddress = 0, m_lImpEntryAddress, lpTargetAddress)
    If lpTargetAddress = 0 Then Exit Function
    ImpReplace = memcpy2remote(lpTargetAddress, VarPtr(lAddress), 4)
    m_fImpReplaced = ImpReplace
    Exit Function
Error_Handler:
End Function

'恢复原函数跳转地址
Public Function ImpRestore() As Boolean
    On Error GoTo Error_Handler
    Dim lOldProtect As Long
    If Not m_fImpReplaced Then Exit Function
    If m_lImpJumpAddress = 0 Then Exit Function
    ImpRestore = memcpy2remote(m_lImpEntryAddress, VarPtr(m_lImpJumpAddress), 4)
    m_fImpReplaced = False
    Exit Function
Error_Handler:
End Function

' 定位输入函数
Public Function LocateImportFunction(ByVal vModuleHandleOrName As Variant, ByVal sLibName As String, ByVal vFunction As Variant) As Boolean
    On Error GoTo Error_Handler
    Dim hModule As Long
    hModule = GetModuleHandle(vModuleHandleOrName, m_lProcessID)
    If hModule = 0 Then Exit Function
    If TypeName(vModuleHandleOrName) = "String" Then
        m_sImpModuleName = vModuleHandleOrName
    Else
        m_sImpModuleName = GetModuleFileName(hModule, m_lProcessID)
    End If
    If m_sImpModuleName = "" Then Exit Function
    m_lImpEntryAddress = 0
    m_lImpJumpAddress = 0
    Dim lpIMAGE_IMPORT_DESCRIPTOR As Long
    Dim lpLibName As Long
    Dim lpIMAGE_IMPORT_BY_NAME As Long
    Dim lpFunction As Long, lpTemp As Long, lpCounter As Long
    lpIMAGE_IMPORT_DESCRIPTOR = hModule + MemValue(hModule + MemValue(hModule + &H3C) + &H80)
    Do
        If MemValue(lpIMAGE_IMPORT_DESCRIPTOR) = 0 Then Exit Do
        lpLibName = hModule + MemValue(lpIMAGE_IMPORT_DESCRIPTOR + &HC)
        If UCase$(sLibName) = UCase$(RemoteStrFromPtr(lpLibName)) Then
            lpIMAGE_IMPORT_BY_NAME = hModule + MemValue(lpIMAGE_IMPORT_DESCRIPTOR)
            lpCounter = 0
            Do
                lpFunction = MemValue(lpIMAGE_IMPORT_BY_NAME + lpCounter * 4)
                If lpFunction = 0 Then Exit Do
                lpTemp = lpFunction + 2
                If TypeName(vFunction) <> "String" Then
                    If lpTemp And &H80000000 Then                               '数字入口
                        If (lpTemp And &H7FFFFFFF) = CLng(vFunction) Then Exit Do
                    End If
                Else
                    If (lpTemp And &H80000000) = 0 Then                         '字符入口
                        If UCase$(RemoteStrFromPtr(hModule + lpTemp)) = UCase$(vFunction) Then Exit Do
                    End If
                End If
                lpCounter = lpCounter + 1
            Loop
            If lpFunction <> 0 Then
                lpIMAGE_IMPORT_BY_NAME = hModule + MemValue(lpIMAGE_IMPORT_DESCRIPTOR + &H10)
                m_lImpEntryAddress = lpIMAGE_IMPORT_BY_NAME + lpCounter * 4
                m_lImpJumpAddress = MemValue(m_lImpEntryAddress)
            End If
        End If
        lpIMAGE_IMPORT_DESCRIPTOR = lpIMAGE_IMPORT_DESCRIPTOR + &H14
    Loop
    LocateImportFunction = m_lImpEntryAddress <> 0
    Exit Function
Error_Handler:
End Function

'定位输出函数
Public Function LocateExportFunction(ByVal vModuleHandleOrName As Variant, ByVal vFunction As Variant) As Boolean
    On Error GoTo Error_Handler
    Dim hModule As Long
    hModule = GetModuleHandle(vModuleHandleOrName, m_lProcessID)
    If hModule = 0 Then Exit Function
    If TypeName(vModuleHandleOrName) = "String" Then
        m_sExpModuleName = vModuleHandleOrName
    Else
        m_sExpModuleName = GetModuleFileName(hModule, m_lProcessID)
    End If
    If m_sExpModuleName = "" Then Exit Function
    m_lExpEntryAddress = 0
    m_lExpJumpAddress = 0
    Dim lpEXPORT_TABLE As Long
    Dim lNumberOfNames As Long, lNumberOfFunctions As Long, lBase As Long
    Dim lpNamesTable As Long, lpFunctionsTable As Long, lpOrdinalsTable As Long
    Dim lpIMAGE_ExpORT_BY_NAME As Long
    Dim lpFunction As Long, lNameOrdinal As Long
    lpEXPORT_TABLE = hModule + MemValue(hModule + MemValue(hModule + &H3C) + &H78)
    lBase = MemValue(lpEXPORT_TABLE + &H10)
    lNumberOfFunctions = MemValue(lpEXPORT_TABLE + &H14)
    lNumberOfNames = MemValue(lpEXPORT_TABLE + &H18)
    lpFunctionsTable = hModule + MemValue(lpEXPORT_TABLE + &H1C)
    lpNamesTable = hModule + MemValue(lpEXPORT_TABLE + &H20)
    lpOrdinalsTable = hModule + MemValue(lpEXPORT_TABLE + &H24)
    If TypeName(vFunction) = "String" Then
        Do While lNumberOfNames > 0
            lNumberOfNames = lNumberOfNames - 1
            lpFunction = hModule + MemValue(lpNamesTable + lNumberOfNames * 4)
            lNameOrdinal = MemValue((lpOrdinalsTable + lNumberOfNames * 2), True)
            If UCase$(RemoteStrFromPtr(lpFunction)) = UCase$(vFunction) Then
                If lNameOrdinal + lBase >= lNumberOfFunctions Then Exit Do
                m_lExpEntryAddress = lpFunctionsTable + 4 * lNameOrdinal
                m_lExpJumpAddress = hModule + MemValue(m_lExpEntryAddress)
                Exit Do
            End If
        Loop
    Else
        If vFunction - lBase < lNumberOfFunctions Then
            m_lExpEntryAddress = lpFunctionsTable + 4 * (vFunction - lBase)
            m_lExpJumpAddress = hModule + MemValue(m_lExpEntryAddress)
        End If
    End If
    LocateExportFunction = m_lExpEntryAddress <> 0
    Exit Function
Error_Handler:
End Function

'初始化函数
Public Function SetTargetTo(Optional ByVal pid As Long = -1) As Boolean
    m_lProcessID = IIf(pid = -1, GetCurrentProcessId, pid)
    m_lhProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, m_lProcessID)
    SetTargetTo = (m_lhProcess <> 0)
End Function

'从指针获取字符串
Public Function RemoteStrFromPtr(ByVal lpString As Long) As String
    On Error GoTo Error_Handler
    Dim b(254) As Byte, b2() As Byte, lStrLen As Long, lpBuffer As Long
    Dim lPosOfZero As Long
    Dim lOldProtect As Long
    Dim lBytesReadWrite As Long
    Dim lAddress As Long
    lStrLen = 0
    lpBuffer = VarPtr(b(0))
    If m_lhProcess Then
        ReDim Preserve b2(254)
        Do
            lAddress = lpString + lStrLen * 255
            If VirtualProtectEx(m_lhProcess, ByVal lAddress, 255, PAGE_READWRITE, lOldProtect) <> 0 Then
                ReadProcessMemory m_lhProcess, ByVal lAddress, ByVal lpBuffer, 255, lBytesReadWrite
                Call VirtualProtectEx(m_lhProcess, ByVal lAddress, 255, lOldProtect, lOldProtect)
                CopyMemory b2(lStrLen * 255), ByVal lpBuffer, 255
            Else
                Exit Do
            End If
            lPosOfZero = InStrB(b, ChrB(0))
            If lPosOfZero <> 0 Then Exit Do
            lStrLen = lStrLen + 1
            ReDim Preserve b2((lStrLen + 1) * 255 - 1)
        Loop
        lStrLen = 255 * lStrLen + lPosOfZero - 1
        If lStrLen >= 1 Then
            ReDim Preserve b2(lStrLen - 1)
            RemoteStrFromPtr = StrConv(b2, vbUnicode)
            Debug.Print "RemoteStrFromPtr= " & RemoteStrFromPtr
        End If
    End If
    Exit Function
Error_Handler:
End Function


'=======================================================================================================================
'以上为类公共函数
'=======================================================================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=======================================================================================================================
'以下为内部函数
'=======================================================================================================================

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

'读取指定内存4字节或2字节内容
Private Function MemValue(ByVal lAddress As Long, Optional ByVal fWORD As Boolean = False) As Long
    memcpy2local VarPtr(MemValue), lAddress, IIf(fWORD, 2, 4)
End Function

'根据模块句柄获取模块名称
Private Function GetModuleFileName(ByVal hModule As Long, ByVal pid As Long) As String
    On Error GoTo Error_Handler
    Dim hSnapshot As Long, t As MODULEENTRY32
    Dim lStrLen As Long
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, pid)
    t.dwSize = LenB(t)
    If hSnapshot <> -1 Then
        If Module32First(hSnapshot, t) Then
            Do
                If t.hModule = hModule Then Exit Do
            Loop While Module32Next(hSnapshot, t)
            GetModuleFileName = t.szExePath
            lStrLen = InStr(GetModuleFileName, Chr(0)) - 1
            If lStrLen > 0 Then
                GetModuleFileName = GetFileName(Left$(GetModuleFileName, lStrLen))
            End If
        End If
        CloseHandle hSnapshot
    End If
    Exit Function
Error_Handler:
End Function

'根据模块名称获取模块句柄
Private Function GetModuleHandle(ByVal vModuleHandleOrName As Variant, ByVal pid As Long) As Long
    On Error GoTo Error_Handler
    Dim hSnapshot As Long, t As MODULEENTRY32
    Dim lStrLen As Long, s As String, sModuleName As String
    If TypeName(vModuleHandleOrName) <> "String" Then
        GetModuleHandle = vModuleHandleOrName
        Exit Function
    Else
        sModuleName = vModuleHandleOrName
    End If
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, pid)
    t.dwSize = LenB(t)
    If hSnapshot <> -1 Then
        If Module32First(hSnapshot, t) Then
            Do
                s = t.szExePath
                lStrLen = InStr(s, Chr(0)) - 1
                If lStrLen > 0 Then
                    s = Left$(s, lStrLen)
                End If
                If UCase$(GetFileName(s)) = UCase$(sModuleName) Then
                    GetModuleHandle = t.hModule
                    Exit Do
                End If
            Loop While Module32Next(hSnapshot, t)
        End If
        CloseHandle hSnapshot
    End If
    Exit Function
Error_Handler:
End Function

'从文件路径字符串中分离出文件名
Private Function GetFileName(ByVal sFilePath As String) As String
    On Error GoTo Error_Handler
    Dim lPosOfBackSlash As Long
    Dim lStrLen As Long
    lPosOfBackSlash = InStrRev(sFilePath, "\")
    lStrLen = Len(sFilePath)
    If lStrLen > lPosOfBackSlash Then
        GetFileName = Right$(sFilePath, lStrLen - lPosOfBackSlash)
    End If
    Exit Function
Error_Handler:
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
    m_lImpEntryAddress = 0
    m_lImpJumpAddress = 0
    m_sImpModuleName = ""
    m_lExpEntryAddress = 0
    m_lExpJumpAddress = 0
    m_sExpModuleName = ""
    m_fImpReplaced = False
    m_fExpReplaced = False
    m_fAutoRestore = True
    '默认设定目标进程为当前进程
    Call SetTargetTo(-1)
End Sub

'类销毁
Private Sub Class_Terminate()
    If m_fAutoRestore Then
        '恢复已替换地址
        Call ImpRestore
    End If
    '关闭打开的进程句柄
    CloseHandle m_lhProcess
End Sub
                                                                         
'=======================================================================================================================
'以上为类初始化与销毁过程
'=======================================================================================================================
