VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 功  能：读、写注册表
Option Explicit
' 扩充环境字符串。具体操作过程与命令行处理的所为差不多。
' 也就是说，将由百分号封闭起来的环境变量名转换成那个变量的内容。比如，“%path%”会扩充成完整路径。
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" _
   (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
'用于操作注册表的 API 申明    --开始--
' API 申明   ( 参数若未说明,则全是 ByRef )
' RegCreateKey 在指定的项下创建一个新项。如指定的项已经存在，那么函数会打开现有的项
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
' RegCreateKeyEx 在指定项下创建新项的更复杂的方式。在Win32环境中建议使用这个函数。如指定的项已经存在，则函数会打开现有的项
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
    ByVal lpClass As String, ByVal dwOptions As Long, _
    ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, _
    phkResult As Long, lpdwDisposition As Long) As Long
' RegOpenKey 打开一个现有的注册表项
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
' RegOpenKeyEx
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
' RegCloseKey 关闭系统注册表中的一个项（或键）
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
' RegSetValue 设置指定子键的默认数据
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, _
    ByVal lpData As String, ByVal cbData As Long) As Long
' RegSetValueEx 设置指定子键的指定数据
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
    ByVal dwType As Long, ByVal lpData As Any, ByVal cbData As Long) As Long
' RegSetValueExLong
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
   (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, _
   lpValue As Long, ByVal cbData As Long) As Long
' RegSetValueExByte
Private Declare Function RegSetValueExByte Lib "advapi32.dll" Alias "RegSetValueExA" _
   (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, _
   lpValue As Byte, ByVal cbData As Long) As Long
' RegQueryValueEx 读取数据: 前三个ByVal  后三个 按地址
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Long) As Long
' RegEnumKeyEx 枚举指定项下方的子项
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, _
    ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
' RegEnumValue 枚举指定项的值
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
' RegDeleteValue 删除注册表数据
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal hKey As Long, ByVal lpValueName As String) As Long
' RegDeleteKey   删除注册表主键以外的子键
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'用于操作注册表的 API 申明    --结束--
' 常数申明
'//注册表访问权
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_READ = &H20019

'//打开/建立选项
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1

'//Key 创建/打开
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2

'//预定义存取类型
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF

'//严格代码定义
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_ACCESS_DENIED = 5
Private Const ERROR_NO_MORE_ITEMS = 259
Private Const ERROR_MORE_DATA = 234 '//  错误
' 自定义注册文件错误
Private Const REGAGENT_NOKEY = -1002
Private Const REGAGENT_NOSUBKEY = -1003

'//注册表结构
Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Boolean
End Type

'' 根键结构，数据和字符串形式！
'Private Type RootKeyHexAndString
'    rootKeyHex As ROOTKEY_Const
'    rootKeyString As String
'End Type

' 枚举申明
Public Enum ROOTKEY_Const ' 根键 RootKey，这里申明为枚举
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
    
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_DYN_DATA = &H80000006
End Enum
Public Enum VALUE_Const ' 值的数据类型枚举
    REG_BINARY = 3                     ' Free form binary 二进制数据
    REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number 高位排在底位的双字
    REG_DWORD = 4                      ' 32-bit number 32位双字节数据
    
    REG_EXPAND_SZ = 2                  ' Unicode nul terminated string 扩展字符串,可以加入变量如%PATH%
    REG_MULTI_SZ = 7                   ' Multiple Unicode strings  多字符串
    REG_SZ = 1                         ' Unicode nul terminated string 以0结尾的字符串
    
'    REG_RESOURCE_LIST = 8              ' Resource list in the resource map 设备驱动程序资源列表
'    REG_RESOURCE_REQUIREMENTS_LIST = 10
'    REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
'    REG_NONE = 0                       ' No value type
'    REG_LINK = 6                       ' Symbolic Link (unicode) UNICODE 符号链接
End Enum


' ***************************************************************************************************
' 写入 注册表数据 -- 开始 --' 它们用法基本一样！（若不存在则创建，否则替换原值为新值。）
' ***************************************************************************************************
' 写入 字符串值
Public Function setString_写入字符串值(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
    sValueName_值名称 As String, sValueValue_值值 As String, Optional lValueType_值类型 As VALUE_Const = REG_SZ)
    Dim handle As Long, strT As String: strT = StripNulls(sValueValue_值值)
    Dim RetVal As Long, lngL As Long ' 字符串长度，考虑中英文！
    lngL = LenB(StrConv(strT, vbFromUnicode))
    ' 创建键
    RegCreateKey lRootkey_根键, sKeyName_项名称, handle
    RetVal = RegSetValueEx(handle, sValueName_值名称, ByVal 0&, lValueType_值类型, ByVal strT, lngL)
    '// 返回关闭结果
    RegCloseKey handle
    '// 返回写入成功结果
    setString_写入字符串值 = (RetVal = 0)
End Function

' 写入 双字节值（DWORD）！！！写入的是十进制数据值 ！！！
Public Function setDword_写入双字节值(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
    sValueName_值名称 As String, ByVal lValueValue_值值 As Long) As Boolean
    Dim handle As Long, RetVal As Long
    ' --------------------------------------------------------------------------
    Dim SecAttr As SECURITY_ATTRIBUTES '//键的安全设置
    '//设置新键值的名称和默认安全设置
    SecAttr.nLength = Len(SecAttr) '//结构大小
    SecAttr.lpSecurityDescriptor = 0 '//默认安全权限
    SecAttr.bInheritHandle = True '//设置的默认值
    '// 打开或创建键
    'If RegOpenKeyEx(lRootkey_根键, sKeyName_项名称, ByVal 0&, KEY_ALL_ACCESS, handle) Then Exit Function
    RetVal = RegCreateKeyEx(lRootkey_根键, sKeyName_项名称, ByVal 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SecAttr, handle, RetVal)
    If RetVal Then Exit Function
    ' --------------------------------------------------------------------------
    RetVal = RegSetValueExLong(handle, sValueName_值名称, ByVal 0&, REG_DWORD, lValueValue_值值, Len(lValueValue_值值))
    '// 返回关闭结果
    RegCloseKey handle
    '// 返回写入成功结果
    setDword_写入双字节值 = (RetVal = 0)
End Function

' 写入 二进制值
Public Function setBinary_写入二进制值(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
    sValueName_值名称 As String, ByVal lValueValue_值值 As Variant) As Boolean
    Dim handle As Long, RetVal As Long
    ' --------------------------------------------------------------------------
    Dim SecAttr As SECURITY_ATTRIBUTES '//键的安全设置
    '//设置新键值的名称和默认安全设置
    SecAttr.nLength = Len(SecAttr) '//结构大小
    SecAttr.lpSecurityDescriptor = 0 '//默认安全权限
    SecAttr.bInheritHandle = True '//设置的默认值
    '// 打开或创建键
    'If RegOpenKeyEx(lRootkey_根键, sKeyName_项名称, ByVal 0&, KEY_ALL_ACCESS, handle) Then Exit Function
    RetVal = RegCreateKeyEx(lRootkey_根键, sKeyName_项名称, ByVal 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SecAttr, handle, RetVal)
    If RetVal Then Exit Function
    ' --------------------------------------------------------------------------
    ' 1、lValueValue_值值 ＝ 字节数组，如 B()
    If VarType(lValueValue_值值) = vbArray + vbByte Then
        Dim binValue() As Byte, Length As Long
        binValue = lValueValue_值值
        Length = UBound(binValue) - LBound(binValue) + 1
        RetVal = RegSetValueExByte(handle, sValueName_值名称, ByVal 0&, REG_BINARY, binValue(0), Length)
    ' 2、lValueValue_值值 ＝ 整型或长整型，如 520
    ElseIf VarType(lValueValue_值值) = vbLong Or VarType(lValueValue_值值) = vbInteger Then
        setBinary_写入二进制值 = setDword_写入双字节值(lRootkey_根键, sKeyName_项名称, sValueName_值名称, lValueValue_值值)
        Exit Function
    ' 3、lValueValue_值值 ＝字符串，如 "BE 3E FF AB"
    ElseIf VarType(lValueValue_值值) = vbString Then
        ' 转化数据
        Dim ByteArray() As Byte
        Dim tmpArray() As String '//转换ASCII字符到16进制字节
        Dim i As Long, lb As Long, ub As Long
        Dim sTemp As String: sTemp = lValueValue_值值
        ' 以空格分割字符串
        tmpArray = Split(sTemp, " ")
        lb = LBound(tmpArray): ub = UBound(tmpArray)
        
        ' 为动态数组分配空间
        ReDim ByteArray(lb To ub)
        ' 循环转换
        For i = lb To ub - 1
            ByteArray(i) = CByte(Val("&H" & Right$(tmpArray(i), 2)))
        Next i
        ' 注意：最后一个不知道字符串后面多了2个什么，要用 Left$(tmpArray(ub), 2)
        ByteArray(ub) = CByte(Val("&H" & Left$(tmpArray(ub), 2)))
        
        ' 将数据写入到注册表，注意：最后是 ub - lb + 1
        RetVal = RegSetValueExByte(handle, sValueName_值名称, ByVal 0&, REG_BINARY, ByteArray(0), ub - lb + 1)
    End If
    '// 返回关闭结果
    RegCloseKey handle
    '// 返回写入成功结果
    setBinary_写入二进制值 = (RetVal = 0)
End Function
' ***************************************************************************************************
' 写入 注册表数据 -- 结束 --' 它们用法基本一样！（若不存在则创建，否则替换原值为新值。）
' ###################################################################################################
' 读取 注册表中数据 -- 开始 --
' ###################################################################################################
' 读取 字符串值，可以读取的字符串值类型有： REG_SZ , REG_MULTI_SZ , REG_EXPAND_SZ
Public Function getString_读取字符串值(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
    sValueName_值名称 As String, Optional ByVal sz_Type As VALUE_Const = REG_SZ, _
    Optional blnForceSingleString As Boolean = False) As Variant
    ' 若设置 blnForceSingleString = True 则函数返回单一字符串，且不经任何处理，只去掉字符串尾！
On Error GoTo fG ' 没有值出错
    Dim handle As Long, strR As String
    RegOpenKey lRootkey_根键, sKeyName_项名称, handle
    strR = RegQueryStringValue(handle, sValueName_值名称)
    If sz_Type = REG_SZ Then ' 一般字符串，去掉结尾的符号。
        strR = StripNulls(strR) ' Left$(strR, InStr(1, strR, Chr$(0)) - 1)
    ElseIf sz_Type = REG_EXPAND_SZ Then ' 扩充环境字符串，查询环境变量和返回定义值
        If Not blnForceSingleString Then strR = StripNulls(ExpandEnvStr(StripNulls(strR)))
    ElseIf sz_Type = REG_MULTI_SZ Then ' 多行字符串
        If Not blnForceSingleString Then
            Dim S() As String
            If Len(strR) <> 0 Then ' 读到的是非空字符串，可以分割。
                S = Split(Left$(strR, Len(strR) - 1), Chr$(0))
            Else ' 若是空字符串，要定义S(0) ，否则出错！
                ReDim S(0) As String
            End If
            ' 函数返回值，返回一个字符串数组？！
            getString_读取字符串值 = S()
            RegCloseKey handle
            Exit Function
        End If
    End If
    ' 函数返回值' 去掉结尾的符号。
    getString_读取字符串值 = StripNulls(strR)
    RegCloseKey handle
    Exit Function
fG:
    getString_读取字符串值 = ""
End Function
' // 读取字符串， strBuf 可能是多行的！
Private Function RegQueryStringValue(ByVal hKey As Long, strValueName As String) As String
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    RegQueryValueEx hKey, strValueName, ByVal 0&, lValueType, ByVal 0&, lDataBufSize
    strBuf = String(lDataBufSize, Chr$(0))
    RegQueryValueEx hKey, strValueName, ByVal 0&, ByVal 0&, ByVal strBuf, lDataBufSize
    RegQueryStringValue = strBuf ' Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
End Function
'// 查询环境变量和返回定义值
Private Function ExpandEnvStr(sData As String) As String
    '// 如： %PATH% 则返回 "c:\;c:\windows;"
    Dim C As Long, S As String, sT As String
    sT = sData & "  " ' 不知为什么要加两个字符，否则返回值会少最后两个字符！
    S = "" '// 不支持Windows 95
    '// get the length
    C = ExpandEnvironmentStrings(sT, S, C)
    '// 展开字符串
    S = String$(C - 1, Chr$(0))
    C = ExpandEnvironmentStrings(sT, S, LenB(S))

    '// 返回环境变量
    ExpandEnvStr = StripNulls(S)
End Function

' 读取 双字节值（DWORD）！！！返回的是十进制数据值 ！！！
Public Function getDword_读取双字节值(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
    sValueName_值名称 As String) As Long
    Dim glStatus        As Long         ' 操作成功与否的状态
    Dim llKeyID         As Long         '打开键的ID
    Dim llKeyValue      As Long         '存放读取的双字节值
    getDword_读取双字节值 = Empty       '预先置为空
    glStatus = ERROR_SUCCESS            '假设成功
    '确定参数有效
    If Len(sKeyName_项名称) = 0 Then     '主键未设置（子键未设置则读默认值）
        glStatus = REGAGENT_NOKEY
        Exit Function
    End If
    '首先打开主键
    glStatus = RegOpenKey(lRootkey_根键, sKeyName_项名称, llKeyID)
    
    If glStatus = ERROR_SUCCESS Then    '成功则取双字节值
        glStatus = RegQueryValueEx(llKeyID, sValueName_值名称, ByVal 0&, REG_DWORD, llKeyValue, Len(llKeyValue))
        If glStatus = ERROR_SUCCESS Then
            getDword_读取双字节值 = llKeyValue
        End If
        glStatus = RegCloseKey(llKeyID)
    End If
End Function

' 读取 二进制值，似乎也可以读取双字节值（DWORD）！！！返回的是十进制数据值 ！！！
Public Function getBinary_读取二进制值(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
    sValueName_值名称 As String, Optional blnGetAsLong As Boolean = False, _
    Optional blnGetAsString As Boolean = False) As Variant
    ' blnGetAsLong = True ，则只取得长整型形式的数据，方法不同！！！
    If blnGetAsLong Then
        Dim glStatus        As Long         ' 操作成功与否的状态
        Dim llKeyID         As Long         ' 打开键的ID
        Dim llKeyValue      As Long         ' 存放读取的二进制值
        getBinary_读取二进制值 = Empty      ' 预先置为空
        glStatus = ERROR_SUCCESS            ' 假设成功
        
        '确定参数有效
        If Len(sKeyName_项名称) = 0 Then     '主键未设置（子键未设置则读默认值）
            glStatus = REGAGENT_NOKEY
            Exit Function
        End If
        '首先打开主键
        glStatus = RegOpenKey(lRootkey_根键, sKeyName_项名称, llKeyID)
        If glStatus = ERROR_SUCCESS Then    '成功则取二进制值
            glStatus = RegQueryValueEx(llKeyID, sValueName_值名称, ByVal 0&, REG_BINARY, llKeyValue, Len(llKeyValue))
            If glStatus = ERROR_SUCCESS Then
                getBinary_读取二进制值 = llKeyValue
            End If
            glStatus = RegCloseKey(llKeyID)
        End If
        Exit Function
    End If
    
    Dim handle As Long
    Dim resBinary() As Byte
    Dim Length As Long
    Dim RetVal As Long
    Dim valueType As Long
        
    '// 打开键, 不存在则退出
    If RegOpenKeyEx(lRootkey_根键, sKeyName_项名称, ByVal 0&, KEY_READ, handle) Then
        If blnGetAsString Then
            getBinary_读取二进制值 = "00"
        Else
            ReDim resBinary(0)
            getBinary_读取二进制值 = resBinary()
        End If
        Exit Function
    End If
    '// 准备 1K  resBinary 用于接收
    Length = 1024
    ReDim resBinary(0 To Length - 1) As Byte
    '// 读注册表值
    RetVal = RegQueryValueEx(handle, sValueName_值名称, ByVal 0&, valueType, resBinary(0), Length)
    '// 若resBinary 太小则重读
    If RetVal = ERROR_MORE_DATA Then
       '// resBinary放大,且重新读取
       ReDim resBinary(0 To Length - 1) As Byte
       RetVal = RegQueryValueEx(handle, sValueName_值名称, ByVal 0&, valueType, resBinary(0), Length)
    End If
    '// resBinary 调整
    If Length <> UBound(resBinary) + 1 Then
       ReDim Preserve resBinary(0 To Length - 1) As Byte
    End If
    ' 返回字符串，注意：要将字节数组进行转化！
    If blnGetAsString Then
        Dim i As Long
        Dim ret As String, tStr As String
        '循环数据，把字节转换为16进制字符串
        For i = LBound(resBinary) To UBound(resBinary)
           tStr = CStr(Hex(resBinary(i)))
           If (Len(tStr) = 1) Then tStr = "0" & tStr
           ret = ret & " " & tStr
        Next i
        getBinary_读取二进制值 = Trim$(ret)
    Else
        getBinary_读取二进制值 = resBinary()
    End If
End Function
' ###################################################################################################
' 读取 注册表中数据 -- 结束 --
' ###################################################################################################
' 删除 注册表数据（值、项） -- 开始 --
' ***************************************************************************************************
' 删除注册表键值,如果成功返回 True
Public Function DeleteValue_删除值(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
    sValueName_值名称 As String) As Boolean
    Dim lRetval As Long      '//打开和输出注册表键的返回值
    Dim lReghWnd As Long     '//打开注册表键的句柄
    Dim sREGSZData As String '//把获取值放入缓冲区
    Dim lSLength As Long     '//缓冲区大小.  改变缓冲区大小要在调用之后
    '//打开键
    lRetval = RegOpenKeyEx(lRootkey_根键, sKeyName_项名称, ByVal 0&, KEY_ALL_ACCESS, lReghWnd)
    '//成功打开
    If lRetval = ERROR_SUCCESS Then
       '//删除指定值
       lRetval = RegDeleteValue(lReghWnd, sValueName_值名称)  '//如果已存在则先删除
       '//如出现错误则删除值并返回False
       If lRetval <> ERROR_SUCCESS Then Exit Function
       '//注意: 如果成功打开仅关闭注册表键
       lRetval = RegCloseKey(lReghWnd)
       '//如成功关闭则返回 True 或者其它错误
       If lRetval = ERROR_SUCCESS Then DeleteValue_删除值 = True
   End If
End Function

' 删除 项 注意：会删除该项下的所有值；若该项下还有项，则不能删除该项，但会删除该项下的所有的值数据！
' 设置 sKeyName_项名称 = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Hewanglan"，则删除 Hewanglan 这个项。
' sDelKeyName_删除项名称 完全可以不必设置！
Public Function DeleteSubKey_删除项(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
   Optional sDelKeyName_删除项名称 As String = "") As Boolean
    Dim llKeyID     As Long         '键ID
    Dim lRetval     As Long
    ' 确定参数有效' sKeyName_项名称未设置（为空）
    If Len(sKeyName_项名称) = 0 Then DeleteSubKey_删除项 = False: Exit Function

    ' 尝试打开项
    If (RegOpenKeyEx(lRootkey_根键, sKeyName_项名称, ByVal 0&, KEY_ALL_ACCESS, llKeyID) = ERROR_SUCCESS) Then
        lRetval = RegDeleteKey(llKeyID, ByVal sDelKeyName_删除项名称) ' 删除项
        RegCloseKey llKeyID ' 返回关闭结果
    Else
        lRetval = -1 ' 该项不存在，打开失败，更不用删除了。
    End If
    DeleteSubKey_删除项 = (lRetval = 0)
End Function
' ***************************************************************************************************
' 删除 注册表数据（值、项） -- 结束 --
' ###################################################################################################
' ### 创建 项 ### （注意：不写任何数据。其实用上面的 写入 *值 的函数也可以。）如：
' ###################################################################################################
' setBinary_写入二进制值 HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run" & "\Liu", "", ""
' 上面调用后，将在 Run 项下创建名为 Liu 的项，与调用下面的函数效果一样。
' CreateKey_创建项 HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run" & "\Liu"
' 创建 SubKey_项名称
Public Function CreateKey_创建项(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String) As Long
    Dim glStatus        As Long         ' 操作成功与否的状态
    Dim llKeyID         As Long         '键ID
    glStatus = ERROR_SUCCESS            '假设成功
    ' 确定参数有效
    If Len(sKeyName_项名称) = 0 Then ' sKeyName_项名称 未设置
        glStatus = REGAGENT_NOKEY
        Exit Function
    End If
    ' 创建 sKeyName_项名称
    glStatus = RegCreateKey(lRootkey_根键, sKeyName_项名称, llKeyID)
    If glStatus = ERROR_SUCCESS Then
        CreateKey_创建项 = llKeyID
    End If
End Function
' ###################################################################################################
' ### 创建 项 ###
' ###################################################################################################

' ###################################################################################################
' ### 存在性的判断（项、子项、值） ###
' ###################################################################################################
' 1、判断 项是否存在？ ###
Public Function IsKeyExists_项存在吗(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String) As Boolean
    Dim lReghWnd As Long     ' 打开注册表键的句柄
    ' 尝试打开项
    If (RegOpenKeyEx(lRootkey_根键, sKeyName_项名称, ByVal 0&, KEY_ALL_ACCESS, lReghWnd) = ERROR_SUCCESS) Then
        ' 注意: 如果成功打开关闭注册表键，设置函数返回值。
        RegCloseKey lReghWnd
        IsKeyExists_项存在吗 = True
    Else
        IsKeyExists_项存在吗 = False ' 不成功，说明键不存在。
    End If
End Function

' 2、判断 项下面是否有子项？ ###
Public Function HasSubKey_有子项吗(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String) As Boolean
    Dim lReghWnd As Long     ' 打开注册表键的句柄
    Dim sName As String, RetVal As Long, dwIndex As Long
    ' 尝试打开项
    If (RegOpenKey(lRootkey_根键, sKeyName_项名称, lReghWnd) = ERROR_SUCCESS) Then
        dwIndex = 0
        sName = Space(255)
        RetVal = Len(sName)
        HasSubKey_有子项吗 = False
        ' 枚举项下面的子项
        Do While RegEnumKeyEx(lReghWnd, dwIndex, sName, RetVal, ByVal 0&, _
                           vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            ' 注意: 当第一次找到一个子项后，退出循环，设置函数返回值。
            HasSubKey_有子项吗 = True
            Exit Do
        Loop
        RegCloseKey lReghWnd ' 关闭注册表键
    Else
        HasSubKey_有子项吗 = False ' 它没有子项
    End If
End Function

' 3、判断 值是否存在？ 不能判断 sValueName_值名称 ="" 的！ ###
Public Function IsValueExists_值存在吗(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
    sValueName_值名称 As String) As Boolean
    ' 打开注册表键的句柄
    Dim lReghWnd As Long
    Dim sData As String
    ' 尝试打开项
    If (RegOpenKeyEx(lRootkey_根键, sKeyName_项名称, ByVal 0&, KEY_READ, lReghWnd) = ERROR_SUCCESS) Then
        ' 查询数据
        If (RegQueryValueEx(lReghWnd, sValueName_值名称, ByVal 0&, 0, ByVal sData, 1) = ERROR_SUCCESS) Then
            RegCloseKey lReghWnd ' 关闭注册表键
            IsValueExists_值存在吗 = True ' 值存在
        Else
            IsValueExists_值存在吗 = False ' 值不存在
        End If
    Else
        IsValueExists_值存在吗 = False ' 项不存在，打开项失败了。
    End If
End Function
' ###################################################################################################
' ### 存在性的判断（项、子项、值） ###
' ###################################################################################################

' ###################################################################################################
' ###　枚举注册表中的数据（子项名、值）　###
' ###################################################################################################
' 1、枚举子项名： sKeyName_项名称 下面的所有子项名
' 若 sKeyName_项名称 = "" 则列出根键（如：HKEY_LOCAL_MACHINE）下的所有子项。
Public Function EnumSubKeys_枚举子项名(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
    sSubKey_子项名() As String) As Long
    Dim sName As String, RetVal As Long
    Dim lCount As Long      ' 子项的个数
    Dim lReghWnd As Long    ' 打开注册表键的句柄
    
    Erase sSubKey_子项名 ' 先清除原数组数据
    
    ' 尝试打开项
    If (RegOpenKey(lRootkey_根键, sKeyName_项名称, lReghWnd) = ERROR_SUCCESS) Then
        lCount = 0 ' 初始化项的个数为0,即数组下标从0开始。
        sName = Space(255)
        RetVal = Len(sName)
        
        ' 当没有遇到 ERROR_NO_MORE_ITEMS 时，一直循环
        While RegEnumKeyEx(lReghWnd, lCount, sName, RetVal, ByVal 0&, _
                           vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            
            ReDim Preserve sSubKey_子项名(lCount) ' 重新定义动态数组 (+1)
            
            sSubKey_子项名(lCount) = Left$(sName, RetVal) ' 添加子项名称到数组
                        
            ' 准备读取下一个数据
            lCount = lCount + 1 ' 子项的个数累加 (+1)
            sName = Space(255)
            RetVal = Len(sName)
            
        Wend
    
        RegCloseKey lReghWnd ' 关闭注册表键
    Else
        lCount = -1 ' 出错 （项不存在，不可能枚举！）
    End If
    
    ' 函数返回值（子项的总个数 -1 ，因为下标是从 0 开始的）。
    EnumSubKeys_枚举子项名 = lCount
End Function

' 2、枚举值： sKeyName_项名称 下面所有的值
Public Function EnumValues_枚举值(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
    sValueName_值名称() As String, sValueValue_值值() As Variant, Optional ByVal OnlyType_值的类型 As VALUE_Const = -1) As Long
    
    Dim mainKey As Long, rName As String, Cnt As Long
    Dim rData As String, rType As Long, RetData As Long, RetVal As Long
    Dim lCount As Long
    ' 尝试打开项
    If RegOpenKey(lRootkey_根键, sKeyName_项名称, mainKey) = ERROR_SUCCESS Then
        ' 初始化
        lCount = 0
        Cnt = 0
        rName = Space(255)
        rData = Space(255)
        RetVal = 255
        RetData = 255
        Erase sValueName_值名称
        Erase sValueValue_值值
        
        ' 循环取得该项下面的所有值的数据（值的名称、具体值），若名称为(默认)的值无数据，不会取得它的数据。
        While RegEnumValue(mainKey, Cnt, rName, RetVal, 0, _
                           rType, ByVal rData, RetData) <> ERROR_NO_MORE_ITEMS
                
            If (OnlyType_值的类型 = -1) Or (OnlyType_值的类型 = rType) Then ' 判断值的类型，只列给定类型。
                
                ReDim Preserve sValueName_值名称(lCount) As String ' 重新定义动态数组 (+1)
                ReDim Preserve sValueValue_值值(lCount) As Variant
                
                ' 值的名称 ！！！注意：去掉字符串尾！！！
                sValueName_值名称(lCount) = StripNulls(Trim$(Left$(rName, RetVal)))
                
                ' 取得值的具体值
                If (rType = REG_BINARY) Then
                    If RetData > 0 Then ' 注意：>0 ，有待改进！
                        sValueValue_值值(lCount) = getBinary_读取二进制值(lRootkey_根键, sKeyName_项名称, sValueName_值名称(lCount), , True)
                    Else
                        sValueValue_值值(lCount) = ""
                    End If
                ElseIf (rType = REG_DWORD) Then
                    sValueValue_值值(lCount) = getDword_读取双字节值(lRootkey_根键, sKeyName_项名称, sValueName_值名称(lCount))
                Else 'If (rType = REG_SZ) Then' 其他类型，暂时只考虑字符串！
                    sValueValue_值值(lCount) = getString_读取字符串值(lRootkey_根键, sKeyName_项名称, sValueName_值名称(lCount), rType, True)
                End If
                    
                lCount = lCount + 1 ' 值的个数累加 (+1)
            End If
                
            ' 准备读取下一个数据
            Cnt = Cnt + 1
            rName = Space(255)
            rData = Space(255)
            RetVal = 255
            RetData = 255
        Wend ' 循环
        RegCloseKey lRootkey_根键 ' 完成后，关闭注册表键
    Else
        lCount = -1 ' 出错，返回 －1
    End If
    ' 函数返回值（若名称为(默认)的值有数据，则是值的总个数；否则，是总个数 -1 ！！！）。
    EnumValues_枚举值 = lCount
End Function
' ###################################################################################################
' ###　枚举注册表中的数据（项名、值）　###
' ###　导出、导入注册表中的数据　###
' ###################################################################################################
' 1、ExportFormReg_导出注册表项
Public Function ExportFormReg_导出注册表项(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
    ByVal sRegFile_导出目标文件路径 As String, Optional RegVersion4_导出版本4 As Boolean = True, _
    Optional IncludeSubkeys_包含子项 As Boolean = True, Optional ByVal Output_状态输出文本框 As Variant) As Boolean
    ' 若文件已经存在，直接退出。
    If (Dir$(sRegFile_导出目标文件路径, vbReadOnly Or vbHidden Or vbArchive Or vbSystem) <> "") Then
        ExportFormReg_导出注册表项 = False: Exit Function
    End If
    On Error GoTo errh
    Dim opened      As Boolean  ' 打开标记：是否已经打开
    Dim fn          As Integer  ' 可用的文件号
    Dim strRegVer   As String   ' 注册表版本，4 和 5.0 导出的文件第一行文字不同。
    strRegVer = IIf(RegVersion4_导出版本4, "REGEDIT4", "Windows Registry Editor Version 5.00")
    ' 写文件：导出的文件第一行文字
    fn = FreeFile                       ' 取得可用的文件号
    Open sRegFile_导出目标文件路径 For Output As #fn     ' 打开文件准备写入
        opened = True                   ' 设置 打开标记
        Print #fn, strRegVer & vbCrLf   ' 导出的文件第一行文字。
    Close #fn                           ' 关闭文件
    opened = False                      ' 设置 打开标记
    ' 调用函数 1.1 生成注册表项文件（.reg 文件）
    If (GenerateRegFile(lRootkey_根键, sKeyName_项名称, sRegFile_导出目标文件路径, RegVersion4_导出版本4, IncludeSubkeys_包含子项, Output_状态输出文本框) = False) Then GoTo errh
    ExportFormReg_导出注册表项 = True   ' 成功导出了注册表文件（.reg 文件）函数返回值 True
    Exit Function
errh:
    On Error Resume Next
    If opened Then Close #fn            ' 若前面已经打开文件，关闭！
    ExportFormReg_导出注册表项 = False  ' 出错，函数返回值 False
End Function
' 1.1 生成注册表项文件（.reg 文件）
Private Function GenerateRegFile(ByVal lRootkey_根键 As ROOTKEY_Const, sKeyName_项名称 As String, _
    sRegFile_导出目标文件路径 As String, Optional RegVersion4_导出版本4 As Boolean = True, _
    Optional IncludeSubkeys_包含子项 As Boolean = True, Optional Output_状态输出文本框 As Variant) As Boolean
    On Error GoTo errh
    Dim SubKeyName()    As String               ' 子项名称
    Dim aValueName()    As String               ' 值的名称
    Dim aValueValue()   As Variant              ' 值的具体数据值
    Dim opened          As Boolean              ' 打开标记：是否已经打开
    Dim fn              As Integer              ' 可用的文件号
    Dim u As Long, l    As Long, tmp As String  ' 循环变量及临时变量
    Dim hasOutput       As Boolean              ' 是否传入了参数：Output_状态输出文本框，注意：Variant 申明！
    Dim sPath           As String               ' 要导出的注册表项位置
    sPath = RootKeyToString(lRootkey_根键) & "\" & sKeyName_项名称
    hasOutput = Not IsMissing(Output_状态输出文本框)
    If hasOutput Then
        DoEvents
        Output_状态输出文本框.Text = sPath ' 在状态文本框显示当前注册表项位置
    End If
    ' 写文件：导出的文件第一行以后的所有文字
    fn = FreeFile                                       ' 取得可用的文件号
    Open sRegFile_导出目标文件路径 For Append As #fn    ' 以追加方式写文件，将数据追加到文件未
        Print #fn, "[" & sPath & "]"                    ' 要导出的注册表项名称
        '如果 (默认) 不为空，则需要处理
        If (getString_读取字符串值(lRootkey_根键, sKeyName_项名称, "") <> vbNullChar) Then
            Print #fn, "@=" & Chr$(34) & getString_读取字符串值(lRootkey_根键, sKeyName_项名称, "") & Chr$(34)
        End If
        ' 取得所有字符串数据。类型：REG_SZ 其他呢？有待改进！！！
        u = EnumValues_枚举值(lRootkey_根键, sKeyName_项名称, aValueName, aValueValue, REG_SZ) - 1
        For l = 0 To u
            If (Len(aValueName(l)) > 0) Then
                Print #fn, Chr$(34) & aValueName(l) & Chr$(34) & "=" & _
                           Chr$(34) & aValueValue(l) & Chr$(34)
            End If
        Next
        ' 取得所有二进制数据。 类型：REG_BINARY
        u = EnumValues_枚举值(lRootkey_根键, sKeyName_项名称, aValueName, aValueValue, REG_BINARY) - 1
        For l = 0 To u
            Print #fn, Chr$(34) & aValueName(l) & Chr$(34) & "=hex:" & _
                  Replace(Trim$(aValueValue(l)), " ", ",")
        Next
        ' 取得所有 Dword 数据。 类型：REG_DWORD
        u = EnumValues_枚举值(lRootkey_根键, sKeyName_项名称, aValueName, aValueValue, REG_DWORD) - 1
        For l = 0 To u
            tmp = IIf(RegVersion4_导出版本4, "0x", "") & Right$("00000000" & Hex$(aValueValue(l)), 8)
            Print #fn, Chr$(34) & aValueName(l) & Chr$(34) & "=dword:" & tmp
        Next
        ' 写入一空行
        Print #fn, ""
        On Error Resume Next ' 关闭文件，设置打开标记
        Close #fn
        opened = False
        ' 如果需要导出其子项，则历遍所有子项。
        If IncludeSubkeys_包含子项 Then
            u = EnumSubKeys_枚举子项名(lRootkey_根键, sKeyName_项名称, SubKeyName) - 1
            For l = 0 To u
                If (GenerateRegFile(lRootkey_根键, SubKeyName(l), sRegFile_导出目标文件路径, RegVersion4_导出版本4, _
                    IncludeSubkeys_包含子项, Output_状态输出文本框) = False) Then GoTo errh
            Next
        End If
    Close #fn ' 关闭文件，设置打开标记
    opened = False
    GenerateRegFile = True      ' 文件写入成功，函数返回值 True
    Exit Function
errh:
    On Error Resume Next
    If opened Then Close #fn    ' 若前面已经打开文件，关闭！
    GenerateRegFile = False     ' 出错，函数返回值 False
End Function
' 2、将 .reg 文件导入到注册表中
Public Function ImportToReg_导入到注册表(ByVal sRegFile_欲导入的文件路径 As String) As Boolean
    On Error GoTo noExists
    
    Dim Lines() As String, i As Long, sTemp As String, FileNum As Integer
    Dim Args() As String, k As Long, sLine As String, l As Long, tmp As String
    Dim lRootKey As ROOTKEY_Const, strSubKey As String
    ' 若文件已经存在，直接退出。
    If Len(Dir(sRegFile_欲导入的文件路径, vbReadOnly Or vbHidden Or vbArchive Or vbSystem)) = 0 Then
noExists:
        ImportToReg_导入到注册表 = False
        Exit Function
    End If
    FileNum = FreeFile ' 取得可用的文件号
    Open sRegFile_欲导入的文件路径 For Binary As #FileNum   ' 打开文件
        sTemp = Input(LOF(FileNum), #FileNum)               ' 取得其中的所有数据内容
    Close #FileNum
    ' 按行将数据内容分割成数组
    Lines = Split(sTemp, vbCrLf)
    ' 判断文件内容是否正确，注意：仅判断第一行的标记。
    If UCase$(StripNulls(Lines(0))) <> "REGEDIT4" Then
        If UCase$(StripNulls(Lines(0))) <> "WINDOWS REGISTRY EDITOR VERSION 5.00" Then
            ImportToReg_导入到注册表 = False
            Exit Function
        End If
    End If
    ' 循环每一行
    For i = 1 To UBound(Lines)
        sLine = Replace(Trim$(Lines(i)), Chr$(34), vbNullString)
        If (Left$(sLine, 1) = "[") Then ' "[" 标记
            sLine = Mid$(sLine, 2, Len(sLine) - 2) ' 删除 "[" 和 "]"
            If (Left$(sLine, 1) = "-") Then ' 我们需要删除此项
                sTemp = Mid$(sLine, 2, Len(sLine) - 1) ' 删除 "-"
                lRootKey = StringToRootKey(sTemp, strSubKey) ' 字符串转化为根键值和项名称
                DeleteSubKey_删除项 lRootKey, strSubKey    ' 调用函数：删除项
            Else
                For k = i + 1 To UBound(Lines)
                    sTemp = Trim$(Replace(Lines(k), Chr$(34), "")) ' 删除引号
                    If (Left$(sTemp, 1) = "[") Then ' 遇到新项，返回
                        i = k - 1
                        Exit For
                    End If
                    If (sTemp = "") Or (InStr(1, sTemp, "=") < 1) Or _
                       (Left$(sTemp, 1) = ";") Then GoTo jump ' 跳过此行
                    Args = Split(sTemp, "=") 'get arguments
                    If (Trim$(Args(1)) = "-") Then ' 删除值
                        lRootKey = StringToRootKey(sLine, strSubKey) ' 字符串转化为根键值和项名称
                        DeleteValue_删除值 lRootKey, strSubKey, Args(0)
                    Else ' 添加数据值
                        ' 二进制值
                        If (LCase$(Left$(Args(1), 4)) = "hex:") Then 'binary
                            tmp = Replace(Mid$(Args(1), 5, Len(Args(1)) - 4), _
                                                                         ",", " ")
                            lRootKey = StringToRootKey(sLine, strSubKey) ' 字符串转化为根键值和项名称
                            setBinary_写入二进制值 lRootKey, strSubKey, Args(0), tmp
                        ' DWORD 值
                        ElseIf (LCase$(Left$(Args(1), 6)) = "dword:") Then
                            lRootKey = StringToRootKey(sLine, strSubKey) ' 字符串转化为根键值和项名称
                            setDword_写入双字节值 lRootKey, strSubKey, Args(0), _
                                CLng(Val("&H" & Mid$(Args(1), 7, Len(Args(1)) - 6)))
                        ' 字符串值，此处类型仅 REG_SZ
                        Else
                            lRootKey = StringToRootKey(sLine, strSubKey) ' 字符串转化为根键值和项名称
                            setString_写入字符串值 lRootKey, strSubKey, IIf(Args(0) = "@", "", Args(0)), Args(1)
                            If (Args(0) = "@") And (Args(1) = "") Then _
                            DeleteValue_删除值 lRootKey, strSubKey, "" '(数值未设置)
                        End If
                    End If
jump:
                Next
            End If
        End If
    Next
    ImportToReg_导入到注册表 = True ' 成功导入，函数返回值
End Function
' ###################################################################################################
' ###　导出、导入注册表中的数据　###
' ###################################################################################################
' 1、去掉字符串尾
Public Function StripNulls(OriginalStr As String) As String
   ' This removes the extra Nulls so String comparisons will work
   If (InStr(OriginalStr, Chr$(0)) > 0) Then
      OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr$(0)) - 1)
   End If
   StripNulls = OriginalStr
End Function
' 2、根键值转化为字符串
Public Function RootKeyToString(ByVal lRootkey_根键 As ROOTKEY_Const) As String
    ' 根键值转化为字符串
    Select Case lRootkey_根键
        Case HKEY_CLASSES_ROOT:     RootKeyToString = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_USER:     RootKeyToString = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE:    RootKeyToString = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS:            RootKeyToString = "HKEY_USERS"
        Case HKEY_CURRENT_CONFIG:   RootKeyToString = "HKEY_CURRENT_CONFIG"
        Case HKEY_PERFORMANCE_DATA: RootKeyToString = "HKEY_PERFORMANCE_DATA"
        Case HKEY_DYN_DATA:         RootKeyToString = "HKEY_DYN_DATA"
    End Select
End Function
' 3、字符串转化为根键值和项名称，主要在导入文件到注册表中时有用。
' 参数： strPath 为字符串形式的注册表路径，如 "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\HeWanglan"
'        sKeyName_项名称 返回项名称，如 "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\HeWanglan"
Public Function StringToRootKey(ByVal strPath As String, sKeyName_项名称 As String) As ROOTKEY_Const
    Dim pos As Long ' 取得第一个反斜杠位置
    Dim rk As String ' 字符串形式的根键值
    ' 将长的根键值替换成短的
    strPath = Replace$(strPath, "HKEY_CURRENT_USER", "HKCU", , , 1)
    strPath = Replace$(strPath, "HKEY_LOCAL_MACHINE", "HKLM", , , 1)
    strPath = Replace$(strPath, "HKEY_CLASSES_ROOT", "HKCR", , , 1)
    strPath = Replace$(strPath, "HKEY_USERS", "HKUS", , , 1)
    strPath = Replace$(strPath, "HKEY_PERFORMANCE_DATA", "HKPD", , , 1)
    strPath = Replace$(strPath, "HKEY_DYN_DATA", "HKDD", , , 1)
    strPath = Replace$(strPath, "HKEY_CURRENT_CONFIG", "HKCC", , , 1)
    pos = InStr(1, strPath, "\") ' 取得第一个反斜杠位置
    If (pos = 0) Then ' 只有根键
        rk = UCase$(strPath)
        sKeyName_项名称 = ""
    Else
        rk = UCase$(Left$(strPath, 4)) ' 取得 根键值
        sKeyName_项名称 = Right$(strPath, Len(strPath) - pos) ' 取得 项名称
    End If
    ' 字符串转化为根键值，函数返回值。
    Select Case rk
        Case "HKCU": StringToRootKey = HKEY_CURRENT_USER
        Case "HKLM": StringToRootKey = HKEY_LOCAL_MACHINE
        Case "HKCR": StringToRootKey = HKEY_CLASSES_ROOT
        Case "HKUS": StringToRootKey = HKEY_USERS
        Case "HKPD": StringToRootKey = HKEY_PERFORMANCE_DATA
        Case "HKDD": StringToRootKey = HKEY_DYN_DATA
        Case "HKCC": StringToRootKey = HKEY_CURRENT_CONFIG
    End Select
End Function