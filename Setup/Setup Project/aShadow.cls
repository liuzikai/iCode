VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "aShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* cShadow - cSelfSub based sample that creates a form shadow.
'*
'* Note: because this sample will likely be used in the wild as-is, I've commented out the unused
'*  subclass procedures, sc_DelMsg, sc_CallOrigWndProc and the sc_lParamUser property
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
'* v1.1 Shadow color property added...................................................... 20060406
'* v1.2 把阴影优化成四周阴影 咱干的是没啥技术含量的活，向牛叉的原作者致敬！ akebe  136775565
'*************************************************************************************************

Option Explicit

'-Selfsub declarations----------------------------------------------------------------------------
Private Enum eMsgWhen                                                       'When to callback什么时候回调
  MSG_BEFORE = 1                                                            'Callback before the original WndProc回调之前最初的指向
  MSG_AFTER = 2                                                             'Callback after the original WndProc回调后原来的指向
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc回调之前和之后原来的指向
End Enum

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback所有消息回调
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries味精表条目的数量
Private Const WNDPROC_OFF   As Long = &H38                                  'Thunk offset to the WndProc execution address 铛抵消指向执行地址
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index SetWindowsLong指向索引
Private Const IDX_SHUTDOWN  As Long = 1                                     'Thunk data index of the shutdown flag 铛数据索引的关闭标记
Private Const IDX_HWND      As Long = 2                                     'Thunk data index of the subclassed hWnd 铛hWnd子类化指数的数据
Private Const IDX_WNDPROC   As Long = 9                                     'Thunk data index of the original WndProc 铛数据指数最初的指向
Private Const IDX_BTABLE    As Long = 11                                    'Thunk data index of the Before table铛数据在表的索引
Private Const IDX_ATABLE    As Long = 12                                    'Thunk data index of the After table 铛后的数据索引表
Private Const IDX_PARM_USER As Long = 13                                    'Thunk data index of the User-defined callback parameter data index铛数据索引定义的回调参数数据的索引

Private z_ScMem             As Long                                         'Thunk base address 铛基地址
Private z_Sc(64)            As Long                                         'Thunk machine-code initialised here 铛机器码这里初始化的
Private z_Funk              As Collection                                   'hWnd/thunk-address collection hWnd / thunk-address集合

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Const DEF_COLOR             As Long = 0
Private Const DEF_DEPTH             As Long = 7
Private Const DEF_TRANSPARENCY      As Long = 140

Private Const AC_SRC_OVER           As Long = &H0
Private Const AC_SRC_ALPHA          As Long = &H1
Private Const BITSPIXEL             As Long = 12
Private Const SIZE_RESTORED         As Long = 0
Private Const SW_HIDE               As Long = 0
Private Const SW_SHOWNOACTIVATE     As Long = 4
Private Const SWP_HIDEWINDOW        As Long = &H80
Private Const SWP_SHOWWINDOW        As Long = &H40
Private Const ULW_ALPHA             As Long = &H2

Private Const WM_MDIACTIVATE        As Long = &H222
Private Const WM_DISPLAYCHANGE      As Long = &H7E
Private Const WM_THEMECHANGED       As Long = &H31A
Private Const WM_WINDOWPOSCHANGED   As Long = &H47
Private Const WM_SIZE               As Long = &H5
Private Const WS_EX_LAYERED         As Long = &H80000
Private Const WS_EX_NOPARENTNOTIFY  As Long = &H4
Private Const WS_EX_TRANSPARENT     As Long = &H20
Private Const WS_POPUP              As Long = &H80000000

Private Type tBGRA
    Blue                            As Byte
    Green                           As Byte
    Red                             As Byte
    Alpha                           As Byte
End Type

Private Type tBITMAPINFOHEADER
  biSize                            As Long
  biWidth                           As Long
  biHeight                          As Long
  biPlanes                          As Integer
  biBitCount                        As Integer
  biCompression                     As Long
  biSizeImage                       As Long
  biXPelsPerMeter                   As Long
  biYPelsPerMeter                   As Long
  biClrUsed                         As Long
  biClrImportant                    As Long
End Type

Private Type tBLENDFUNCTION
  BlendOp                           As Byte
  BlendFlags                        As Byte
  SourceConstantAlpha               As Byte
  AlphaFormat                       As Byte
End Type

Private Type tOSVERSIONINFO
  dwOSVersionInfoSize               As Long
  dwMajorVersion                    As Long
  dwMinorVersion                    As Long
  dwBuildNumber                     As Long
  dwPlatformId                      As Long
  szCSDVersion                      As String * 128
End Type

Private Type tPOINT
  X                                 As Long
  Y                                 As Long
End Type

Private Type tRECT
  Left                              As Long
  Top                               As Long
  Right                             As Long
  Bottom                            As Long
End Type

Private Type tSAFEARRAYBOUND
  cElements                         As Long
  lLbound                           As Long
End Type

Private Type tSAFEARRAY2D
  cDims                             As Integer
  fFeatures                         As Integer
  cbElements                        As Long
  cLocks                            As Long
  pvData                            As Long
  Bounds(0 To 1)                    As tSAFEARRAYBOUND
End Type



Private Type tSIZE
  cx                                As Long
  cy                                As Long
End Type

Private Type tWINDOWPOS
  hWnd                              As Long
  hWndInsertAfter                   As Long
  X                                 As Long
  Y                                 As Long
  cx                                As Long
  cy                                As Long
  flags                             As Long
End Type

Private m_Color                     As Long                                 'Private shadow color property value私有阴影颜色属性值
Private m_Depth                     As Long                                 'Private shadow depth property value私有阴影深度属性值
Private m_Transparency              As Long                                 'Private shadow transparency property value私有阴影透明属性值

Private bIsLayered                  As Boolean                              'Layered windows supported布尔的分层窗口支持
Private bIsLuna                     As Boolean                              'Luna theme? 布尔
Private bIsXP                       As Boolean                              'Windows XP?布尔的Windows XP吗
Private bLastShow                   As Boolean                              'The previous show state布尔之前的显示状态
Private cx                          As Long                                 'Width 宽度
Private cy                          As Long                                 'Height 高度
Private hWndBt                      As Long                                 'Bottom shadow window handle 底影子窗口句柄
Private hWndRt                      As Long                                 'Right shadow window handle 正确的影子窗口句柄
Private hWndTt                      As Long                                 '上面的阴影（akebe）
Private hWndLt                      As Long                                 '坐标的阴影（akebe）
Private hWndForm                    As Long                                 'Parent window handle 父窗口句柄
Private wp                          As tWINDOWPOS                           'Parent window position 父窗口的位置

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As tBITMAPINFOHEADER, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal cchMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetThemeDocumentationProperty Lib "uxtheme.dll" (ByVal pszThemeName As Long, ByVal pszPropertyName As Long, ByVal pszValueBuff As Long, ByVal cchMaxValChars As Long) As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As tOSVERSIONINFO) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDest As Long, ptDst As Any, pSize As Any, ByVal hdcSrc As Long, ptSrc As Any, ByVal crKey As Long, pBlend As Any, ByVal dwFlags As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Private Sub Class_Initialize()                                              'Class initialization 类初始化
  CheckEnvironment
  m_Depth = DEF_DEPTH
  m_Transparency = DEF_TRANSPARENCY
End Sub

Private Sub Class_Terminate()
  sc_Terminate                                                              'Terminate all subclassing 终止所有子类化
  
  If hWndRt <> 0 Then
    DestroyWindow hWndRt
    DestroyWindow hWndBt
    DestroyWindow hWndLt '(akebe)
    DestroyWindow hWndTt '(akebe)
  End If
End Sub

'Shadow depth property
Public Property Get Color() As Long
  Color = m_Color
End Property
'
Public Property Let Color(ByVal NewValue As Long)
  Const COLOR_SYS_MASK As Long = &H80000000
  
  If (NewValue And COLOR_SYS_MASK) Then                                       'If the system color bit is set 如果系统已经设置了颜色位
    NewValue = GetSysColor(NewValue Xor COLOR_SYS_MASK)                       'Get the translated system color 得到颜色的翻译系统
  End If
  
  If NewValue <> m_Color Then
    m_Color = NewValue
    DisplayShadows
  End If
End Property

'Shadow depth property 阴影深度属性
Public Property Get Depth() As Long
  Depth = m_Depth
End Property
'
Public Property Let Depth(ByVal NewValue As Long)
  If NewValue <> m_Depth Then
    m_Depth = NewValue
    DisplayShadows
  End If
End Property

'Shadow transparency property 影子透明属性
Public Property Get Transparency() As Byte
  Transparency = CByte(m_Transparency)
End Property
'
Public Property Let Transparency(ByVal NewValue As Byte)
  If NewValue <> m_Transparency Then
    m_Transparency = CLng(NewValue)
    DisplayShadows
  End If
End Property

'Set the form to be shadowed设置表单阴影
Public Function Shadow(frm As Form) As Boolean
  If hWndForm <> 0 Then
    zError "Shadow", "Only a single form per cShadow instance is allowed"
    Exit Function
  End If

  If bIsLayered Then
    hWndForm = frm.hWnd
  
    If sc_Subclass(hWndForm) Then
      sc_AddMsg hWndForm, WM_WINDOWPOSCHANGED, MSG_BEFORE '这个要是不执行就不行了
      sc_AddMsg hWndForm, WM_SIZE, MSG_BEFORE
      sc_AddMsg hWndForm, WM_DISPLAYCHANGE, MSG_BEFORE
      
      If bIsXP Then
        sc_AddMsg hWndForm, WM_THEMECHANGED, MSG_BEFORE
      End If
      
      CreateWindows
      Shadow = True
    End If
  End If
End Function

'Validate the OS and color depth 验证操作系统和颜色深度
Private Sub CheckEnvironment()
  Dim OSV As tOSVERSIONINFO

  bIsLayered = False

  With OSV
    .dwOSVersionInfoSize = Len(OSV)                                         'Set the length element 设置长度元素
    GetVersionExA OSV                                                       'Fill the type with OS version info 填满类型操作系统的版本信息

    
    If .dwMajorVersion >= 5 Then                                       'If the major version is 5 or greater then the OS supports transparency 如果主要版本5或更高然后操作系统支持透明度
      bIsLayered = True
      
      If .dwMinorVersion > 0 Then
        bIsXP = True                                                  'Assume luna window shape, if people like this control enough i'll access the theme api's to get the actual window shape.
        bIsLuna = IsLuna                                                '假设luna窗口形状,如果人们喜欢这种控制足够我将访问api的主题来获得实际的窗口的形状
      End If
      
      bIsLayered = (GetDeviceCaps(GetDC(0), BITSPIXEL) >= 16)           'Ensure we have enough screen colors确保我们有足够的屏幕颜色
    End If
  End With
End Sub




'Create the shadow windows 创建影子窗口
Private Sub CreateWindows()
  Const EX_STYLE As Long = WS_EX_LAYERED Or WS_EX_TRANSPARENT Or WS_EX_NOPARENTNOTIFY

  hWndRt = CreateWindowExA(EX_STYLE, "#32770", vbNullString, WS_POPUP, 0, 0, 0, 0, hWndForm, 0, App.hInstance, 0)
  hWndBt = CreateWindowExA(EX_STYLE, "#32770", vbNullString, WS_POPUP, 0, 0, 0, 0, hWndForm, 0, App.hInstance, 0)
  
  hWndTt = CreateWindowExA(EX_STYLE, "#32770", vbNullString, WS_POPUP, 0, 0, 0, 0, hWndForm, 0, App.hInstance, 0) '(akebe)
  hWndLt = CreateWindowExA(EX_STYLE, "#32770", vbNullString, WS_POPUP, 0, 0, 0, 0, hWndForm, 0, App.hInstance, 0) '(akebe)
End Sub

'Display the right and bottom shadows 显示右边和底部的阴影
Private Sub DisplayShadows()
  If bIsLayered Then
    If IsWindowVisible(hWndForm) <> 0 Then
      With wp
        DisplayShadowSub .X + .cx, .Y, m_Depth, .cy + m_Depth + m_Depth, 1   '创建窗体大小的阴影(1右)akebe调整
        DisplayShadowSub .X + .cx, .Y, m_Depth, .cy + m_Depth + m_Depth, 3   '创建窗体大小的阴影(3左)akebe添加
        DisplayShadowSub .X, .Y + .cy, .cx + m_Depth, m_Depth, 2   '创建窗体大小的阴影(2下)akebe调整
        DisplayShadowSub .X, .Y + .cy, .cx + m_Depth, m_Depth, 4   '创建窗体大小的阴影(4上)akebe添加
      End With
    End If
  End If
End Sub

'Display the content of the specified shadow window 显示指定的影子窗口的内容（把只有2个值的bRight改成Integer可以支持多个值 akebe）
Private Sub DisplayShadowSub(ByVal X As Long, ByVal Y As Long, cx As Long, cy As Long, ByVal bRight As Integer)
  Dim dc        As Long
  Dim iX        As Long
  Dim iY        As Long
  Dim hDib      As Long
  Dim hWin      As Long
  Dim nAlpha    As Long
  Dim aPixels() As Long
  Dim pBmpBits  As Long
  Dim pt0       As tPOINT
  Dim pt        As tPOINT
  Dim sz        As tSIZE
  Dim bs        As tBLENDFUNCTION
  Dim bmpHeader As tBITMAPINFOHEADER
  Dim SafeArray As tSAFEARRAY2D

  dc = CreateCompatibleDC(0)                                                'Get a screen compatible memory dc 得到一个屏幕兼容的内存

  With bmpHeader                                                            'Initialize a bitmap header 初始化一个位图标题
    .biSize = Len(bmpHeader)                                                'Bitmap header size位图头大小
    .biWidth = cx                                                           'Bitmap/window pixel width 位图像素宽度/窗口
    .biHeight = cy                                                          'Bitmap/window pixel height 位图像素高度/窗口
    .biPlanes = 1                                                           'Graphics planes 图形的飞机
    .biBitCount = 32                                                        '32bits per pixel BGRA (Blue, Green, Red, Alpha) 32比特每像素BGRA(蓝、绿、红、α)
    .biSizeImage = cx * cy * 4                                              'Memory size, width * height * 32bit 内存大小、宽*高* 32位
  End With

  hDib = CreateDIBSection(dc, bmpHeader, 0, pBmpBits, 0, 0)                 'Create a device independant bitmap as per the header, compatible with the dc (compatible with the screen)
                                                                            '创建一个设备独立位图按头,兼容直流(兼容的屏幕)
  With SafeArray                                                            'Construct a VB safearray header that matches the specs of the bitmap 构造一个VB safearray头规格相匹配的位图
    .cbElements = 4                                                         '4 bytes per element - 32bits per pixel 4个字节每个元素——32位/像素
    .cDims = 2                                                              'We'll treat the pixels as a two dimensional (x, y) array 我们将把像素作为一个二维数组(x,y)
    .pvData = pBmpBits                                                      'The data pointer points to the bitmap data (pixels) 指针指向的数据位图数据(像素)
    .Bounds(0).lLbound = 0                                                  'Lowest bound will be 0 最低为0
    .Bounds(0).cElements = cy                                               'The number of elements元素的数量
    .Bounds(1).lLbound = 0                                                  'Lowest bound will be 0 最低为0
    .Bounds(1).cElements = cx                                               'The number of elements 元素的数量
  End With

  CopyMemory ByVal VarPtrArray(aPixels()), VarPtr(SafeArray), 4             'Copy the address of our safearray over the address of aPixels() safearray 复制我们的safearray aPixels的地址的地址()safearray

  If bRight = 1 Then hWin = hWndRt
  If bRight = 2 Then hWin = hWndBt
  If bRight = 3 Then hWin = hWndTt
  If bRight = 4 Then hWin = hWndLt

  If bRight = 1 Then 'akebe调整
    For iY = 0 To cy - 1
      If iY < cx Then
        nAlpha = (255 * iY) \ cx
      ElseIf iY >= (cy - cx) Then
        nAlpha = ((cy - iY) * 255) \ cx
      Else
        nAlpha = 255
      End If

      For iX = 0 To cx - 1
        aPixels(iX, iY) = MakeBGRA((nAlpha * (cx - iX)) \ cx)
      Next iX
    Next iY
End If

If bRight = 3 Then 'akebe添加
    For iY = 0 To cy - 1  '在阴影高度里面循环

      nAlpha = 255
      
      For iX = 0 To cx - 1 '在阴影宽度里面循环
        aPixels(iX, iY) = MakeBGRA((nAlpha * iX) \ cx)
      Next iX
    Next iY
End If

If bRight = 2 Then 'akebe调整
    For iX = 0 To cx - 1
      If iX < cy Then
        nAlpha = (255 * iX) \ cy
      Else
        nAlpha = 255
      End If

      For iY = 0 To cy - 1
        aPixels(iX, iY) = MakeBGRA((nAlpha * iY) \ cy)
      Next iY
    Next iX
End If

If bRight = 4 Then 'akebe添加
    For iX = 0 To cx - 1
      If iX < cy Then
        nAlpha = (255 * iX) \ cy
      Else
        nAlpha = 255
      End If

      For iY = 0 To cy - 1
        aPixels(iX, iY) = MakeBGRA((nAlpha * (cy - iY)) \ cy)
      Next iY
    Next iX
End If

  
  CopyMemory ByVal VarPtrArray(aPixels()), 0&, 4
  
  With bs                                                                   'Setup the blend function 设置混合函数
    .AlphaFormat = AC_SRC_ALPHA                                             'Use the alpha channel for individual pixel transparency使用alpha通道单个像素的透明度
    .BlendFlags = 0
    .BlendOp = AC_SRC_OVER                                                  'Alpha overlay  阿尔法覆盖
    .SourceConstantAlpha = m_Transparency                                   'Alpha transparency for overall transparencyAlpha透明度整体透明度
  End With

  pt.X = X                                                                  'Setup the window position and size data设置窗口位置和大小的数据
  pt.Y = Y
  sz.cx = cx
  sz.cy = cy

  hDib = SelectObject(dc, hDib)                                             'Select the bitmap into the memory display context选择位图到内存中显示上下文
  UpdateLayeredWindow hWin, dc, pt, sz, dc, pt0, 0, bs, ULW_ALPHA           'Do the layered update分层更新
  SelectObject dc, hDib                                                     'Trash the bitmap 垃圾的位图
  DeleteDC dc                                                               'Delete the memory display context 删除内存显示上下文
End Sub

'Return whether the Luna theme is active 返回Luna主题是否活跃
Private Function IsLuna() As Boolean
  Dim hLib   As Long
  Dim nPos   As Long
  Dim sTheme As String
  Dim sName  As String

  hLib = LoadLibraryA("uxtheme.dll")

  If hLib <> 0 Then
    sTheme = String$(255, 0)

    GetCurrentThemeName StrPtr(sTheme), Len(sTheme), 0, 0, 0, 0
    nPos = InStr(1, sTheme, vbNullChar)

    If nPos > 0 Then
      sTheme = Left$(sTheme, nPos - 1)
      sName = String$(255, 0)

      GetThemeDocumentationProperty StrPtr(sTheme), StrPtr("ThemeName"), StrPtr(sName), Len(sName)
      nPos = InStr(1, sName, vbNullChar)

      If nPos > 0 Then
        sName = Left$(sName, nPos - 1)
        bIsLuna = (StrComp(sName, "Luna", vbTextCompare) = 0)
      End If
    End If

    FreeLibrary hLib
  End If
End Function

'Pre-multiply the shadow color with the passed alpha value. This is needed to get nice looking colors according to MSDN.
'　“Pre-multiply阴影颜色传递的alpha值。这是需要根据MSDN好看的颜色
Private Function MakeBGRA(ByVal Alpha As Byte) As Long
  Dim fFactor As Double
  Dim BGRA    As tBGRA
    
  fFactor = CDbl(Alpha) / 255#                                              'Calculate the factor计算系数
'的注意,nColor RGB格式,这个过程的一部分是BGRA格式转换
'Note that nColor is in RGB format, part of this process is to convert to BGRA format
  With BGRA                                                                 'Blue, Green, Red, Alpha'蓝色,绿色,红色,阿尔法
    .Blue = ((m_Color And &HFF0000) \ &H10000) * fFactor                     'Factor the blue component
    .Green = ((m_Color And &HFF00&) \ &H100&) * fFactor                      'Factor the green component
    .Red = (m_Color And &HFF) * fFactor                                      'Factor the red component
    .Alpha = Alpha                                                          'Store the alpha value
  End With
  
'Copy the BGRA type to long复制BGRA类型长
  CopyMemory MakeBGRA, BGRA, 4
End Function

'Show/hide the shadow windows显示/隐藏影子窗口
Private Sub Show(ByVal bShow As Boolean, Optional ByVal bForce As Boolean = False)
  If Not bForce Then
    If bLastShow = bShow Then
      Exit Sub
    End If
  End If

  bLastShow = bShow

  If bShow Then
    ShowWindow hWndRt, SW_SHOWNOACTIVATE
    ShowWindow hWndBt, SW_SHOWNOACTIVATE
    ShowWindow hWndTt, SW_SHOWNOACTIVATE '(akebe)
    ShowWindow hWndLt, SW_SHOWNOACTIVATE '(akebe)
  Else
    ShowWindow hWndRt, SW_HIDE
    ShowWindow hWndBt, SW_HIDE
    ShowWindow hWndLt, SW_HIDE '(akebe)
    ShowWindow hWndTt, SW_HIDE '(akebe)
  End If
End Sub

'Size/position the shadows 阴影大小/位置
Private Sub SizePos()
  With wp
    If .flags And SWP_HIDEWINDOW Then                                       'If the parent form is being hidden 如果父窗体被隐藏
      Show False                                                            'Hide the shadow windows also 隐藏的影子windows

    Else
      If .cx <> cx Then                                                     'If the parent's width has changed 如果父母的宽度已经改变了
        cx = .cx                                                            'Store the new width “存储新的宽度
        DisplayShadowSub .X, .Y + .cy, .cx + m_Depth, m_Depth, 2   '创建窗体大小的阴影(2下)
        DisplayShadowSub .X, .Y + .cy, .cx + m_Depth, m_Depth, 4   '创建窗体大小的阴影(4上)akebe
      End If

      If .cy <> cy Then                                                     'If the parent's height has changed 如果父母的身高已经改变了
        cy = .cy                                                            'Store the new height 存储新的高度
        DisplayShadowSub .X + .cx, .Y, m_Depth, .cy + m_Depth + m_Depth, 1   '创建窗体大小的阴影(1右)
        DisplayShadowSub .X + .cx, .Y, m_Depth, .cy + m_Depth + m_Depth, 3   '创建窗体大小的阴影(3左)akebe
      End If
      
      
     MoveWindow hWndRt, .X + .cx, .Y - m_Depth, m_Depth, .cy + m_Depth + m_Depth, False         '拷贝阴影窗体到指定位置 右部
     MoveWindow hWndBt, .X - m_Depth, .Y + .cy, .cx + m_Depth, m_Depth, False '拷贝阴影窗体到指定位置 下
     
     MoveWindow hWndTt, .X - m_Depth, .Y, m_Depth, .cy, False            '拷贝阴影窗体到指定位置（akebe）左边
     MoveWindow hWndLt, .X - m_Depth, .Y - m_Depth, .cx + m_Depth, m_Depth, False '拷贝阴影窗体到指定位置（akebe）上边
     
      If (.flags And SWP_SHOWWINDOW) Then
        Show True
      End If
    End If
  End With
End Sub

'-SelfSub codeSelfSub代码------------------------------------------------------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True) As Boolean 'Subclass the specified window handle 子类指定的窗口句柄
'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'* lng_hWnd——处理窗口的子类
'* lParamUser -可选,用户定义的回调参数
'* nOrdinal -可选,序数指数的回调过程。1 =去年私有方法,2 =第二最后一个私有方法,等等。
'* oCallback -可选的,将得到的对象回调。如果未定义,回调被发送到该对象的实例
'* bIdeSafety -可选,启用/禁用IDE的安全措施。注:您真的应该只有禁用IDE安全在一个用户控件设计时

'*************************************************************************************************
Const CODE_LEN      As Long = 260                                           'Thunk length in bytes铛在字节长度
Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg tables 每铛字节分配,数据+代码+味精表
Const PAGE_RWX      As Long = &H40&                                         'Allocate executable memory分配可执行的内存
Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated memory提交分配的内存
Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated memory flag释放分配的内存国旗
Const IDX_EBMODE    As Long = 3                                             'Thunk data index of the EbMode function address铛数据索引EbMode函数的地址
Const IDX_CWP       As Long = 4                                             'Thunk data index of the CallWindowProc function address 铛数据索引CallWindowProc函数的地址
Const IDX_SWL       As Long = 5                                             'Thunk data index of the SetWindowsLong function address铛数据索引SetWindowsLong函数的地址
Const IDX_FREE      As Long = 6                                             'Thunk data index of the VirtualFree function address 铛数据索引VirtualFree函数的地址
Const IDX_BADPTR    As Long = 7                                             'Thunk data index of the IsBadCodePtr function address 铛数据索引IsBadCodePtr函数的地址
Const IDX_OWNER     As Long = 8                                             'Thunk data index of the Owner object's vTable address 铛所有者对象的数据索引的vTable地址
Const IDX_CALLBACK  As Long = 10                                            'Thunk data index of the callback method address铛回调方法的数据索引地址
Const IDX_EBX       As Long = 16                                            'Thunk code patch index of the thunk data铛铛的代码补丁索引数据
Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name 这个例程的名称
  Dim nAddr         As Long
  Dim nID           As Long
  Dim nMyID         As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid 确保窗口句柄是有效的
    zError SUB_NAME, "Invalid window handle"
    Exit Function
  End If
 
  nMyID = GetCurrentProcessId                                               'Get this process's ID 得到这个进程的ID
  GetWindowThreadProcessId lng_hWnd, nID                                    'Get the process ID associated with the window handle 得到窗口句柄关联的进程ID
  If nID <> nMyID Then                                                      'Ensure that the window handle doesn't belong to another process确保窗口句柄不属于另一个进程
    zError SUB_NAME, "Window handle belongs to another process"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner如果用户没有指定回调的所有者
    Set oCallback = Me                                                      'Then it is me然后这是我
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of the specified ordinal method获得指定的地址顺序的方法
  If nAddr = 0 Then                                                         'Ensure that we've found the ordinal method确保我们发现顺序的方法
    zError SUB_NAME, "Callback method not found"
    Exit Function
  End If
    
  If z_Funk Is Nothing Then                                                 'If this is the first time through, do the one-time initialization如果这是第一次通过,一次性初始化
    Set z_Funk = New Collection                                             'Create the hWnd/thunk-address collection创建hWnd / thunk-address集合
    z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
    z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&

    z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk data存储在铛CallWindowProc函数地址数据
    z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data存储在铛SetWindowLong函数地址数据
    z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data存储在铛VirtualFree函数地址数据
    z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the IsBadCodePtr function address in the thunk data存储在铛IsBadCodePtr函数地址数据
  End If
  
  z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable memory分配可执行的内存

  If z_ScMem <> 0 Then                                                      'Ensure the allocation succeeded确保配置成功了
    On Error GoTo CatchDoubleSub                                            'Catch double subclassing 捉双子类化
      z_Funk.Add z_ScMem, "h" & lng_hWnd                                    'Add the hWnd/thunk-address to the collection hWnd / thunk-address添加到集合中
    On Error GoTo 0
  
    If bIdeSafety Then                                                      'If the user wants IDE protection 如果用户希望IDE的保护
      z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data 存储在铛EbMode函数地址数据
    End If
    
    z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address 补丁铛数据地址
    z_Sc(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data 商店的窗口句柄铛数据
    z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data 铛之前存储的地址表的数据
    z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data  商店的地址表后铛数据
    z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data 商店老板回调对象地址的铛数据
    z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data 存储铛的回调地址数据
    z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data 存储铛lParamUser回调参数数据
    
    nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc 设置新的指向,返回原始指向的地址
    If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly 确保新的指向正确设置
      zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
      GoTo ReleaseMemory
    End If
        
    z_Sc(IDX_WNDPROC) = nAddr                                               'Store the original WndProc address in the thunk data 铛数据存储原始指向地址
    RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory 复制铛代码/数据分配的内存
    sc_Subclass = True                                                      'Indicate success 显示成功
  Else
    zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
  End If
  
  Exit Function                                                             'Exit sc_Subclass 退出sc_Subclass
 
CatchDoubleSub:
  zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
  VirtualFree z_ScMem, 0, MEM_RELEASE                                       'sc_Subclass has failed after memory allocation, so release the memory sc_Subclass没有内存分配后,释放内存
End Function

'Terminate all subclassing 终止所有子类化
Private Sub sc_Terminate()
  Dim i As Long

  If Not (z_Funk Is Nothing) Then                                           'Ensure that subclassing has been started 确保子类化已经启动
    With z_Funk
      For i = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order 遍历集合用倒序的窗口句柄
        z_ScMem = .Item(i)                                                  'Get the thunk address 铛的地址
        If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory 确保铛没有已经发布了记忆
          sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
        End If
      Next i                                                                'Next member of the collection 下一个集合的成员
    End With
    Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection 摧毁hWnd / thunk-address集合
  End If
End Sub

'UnSubclass the specified window handle UnSubclass指定的窗口句柄
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started 确保子类化已经启动
    zError "sc_UnSubclass", "Window handle isn't subclassed"
  Else
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the thunk hasn't already released its memory 确保铛没有已经发布了记忆
      zData(IDX_SHUTDOWN) = -1                                              'Set the shutdown indicator 设置关闭指示器
      zDelMsg ALL_MESSAGES, IDX_BTABLE                                      'Delete all before messages 之前删除所有消息
      zDelMsg ALL_MESSAGES, IDX_ATABLE                                      'Delete all after messages 删除所有消息后
    End If
    z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection 从集合中移除指定的窗口句柄
  End If
End Sub

'Add the message value to the window handle's specified callback table 添加消息的窗口句柄的值指定回调表
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory 确保铛没有已经发布了记忆
    If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table... 如果消息之前添加到原始指向表……
      zAddMsg uMsg, IDX_BTABLE                                              'Add the message to the before table 之前将信息添加到表中
    End If
    If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table... 如果消息被添加到原始指向表后……
      zAddMsg uMsg, IDX_ATABLE                                              'Add the message to the after table ”后将消息添加到表中
    End If
  End If
End Sub

'Delete the message value from the window handle's specified callback table 删除消息窗口句柄的值指定回调表
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory 确保铛没有已经发布了记忆
    If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table... 如果消息是要从原始指向前表中删除…
      zDelMsg uMsg, IDX_BTABLE                                              'Delete the message from the before table 删除表之前的消息
    End If
    If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table... 如果要删除的消息后从最初指向表……
      zDelMsg uMsg, IDX_ATABLE                                              'Delete the message from the after table 删除后表的信息
    End If
  End If
End Sub

'Call the original WndProc 调用原始的指向
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory 确保铛没有已经发布了记忆
    sc_CallOrigWndProc = _
        CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter 调用的原始指向窗口句柄参数传递
  End If
End Function

'Get the subclasser lParamUser callback parameter 得到subclasser lParamUser回调参数
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory 确保铛没有已经发布了记忆
    sc_lParamUser = zData(IDX_PARM_USER)                                    'Get the lParamUser callback parameter lParamUser回调参数
  End If
End Property

'Let the subclasser lParamUser callback parameter 让subclasser lParamUser回调参数
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory 确保铛没有已经发布了记忆
    zData(IDX_PARM_USER) = NewValue                                         'Set the lParamUser callback parameter 设置lParamUser回调参数
  End If
End Property

'-The following routines are exclusively for the sc_ subclass routines----------------------------专门为sc_子类以下例程程序

'Add the message to the specified table of the window handle 添加消息到指定表的窗口句柄
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count 表条目数
  Dim nBase  As Long                                                        'Remember z_ScMem 记得z_ScMem
  Dim i      As Long                                                        'Loop index 循环索引

  nBase = z_ScMem                                                            'Remember z_ScMem so that we can restore its value on exit 记得z_ScMem退出,这样我们就能恢复其价值
  z_ScMem = zData(nTable)                                                    'Map zData() to the specified table 地图zData()指定的表

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table... 如果ALL_MESSAGES被添加到表……
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES 设置表条目数到ALL_MESSAGES
  Else
    nCount = zData(0)                                                       'Get the current table entry count 获取当前的表条目数
    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow 检查信息表溢出
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
      GoTo Bail
    End If

    For i = 1 To nCount                                                     'Loop through the table entries 遍历表条目
      If zData(i) = 0 Then                                                  'If the element is free... 如果元素是免费的…
        zData(i) = uMsg                                                     'Use this element 使用这个元素
        GoTo Bail                                                           'Bail 保释
      ElseIf zData(i) = uMsg Then                                           'If the message is already in the table... 如果消息已经在桌子上……
        GoTo Bail                                                           'Bail  保释
      End If
    Next i                                                                  'Next message table entry 下一个信息表条目

    nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count 在下降:我= nCount + 1,新表条目数
    zData(nCount) = uMsg                                                    'Store the message in the appended table entry 存储消息附加表中的条目
  End If

  zData(0) = nCount                                                         'Store the new table entry count 存储新的表条目数
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem 恢复z_ScMem的价值
End Sub

'Delete the message from the specified table of the window handle从指定表删除消息的窗口句柄
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count表条目数
  Dim nBase  As Long                                                        'Remember z_ScMem记得z_ScMem
  Dim i      As Long                                                        'Loop index循环索引

  nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit 记得z_ScMem退出,这样我们就能恢复其价值
  z_ScMem = zData(nTable)                                                   'Map zData() to the specified table 地图zData()指定的表

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...如果ALL_MESSAGES正在从表中删除…
    zData(0) = 0                                                            'Zero the table entry count零表条目数
  Else
    nCount = zData(0)                                                       'Get the table entry count表条目数
    
    For i = 1 To nCount                                                     'Loop through the table entries遍历表条目
      If zData(i) = uMsg Then                                               'If the message is found... 如果找到的消息……
        zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use 空味精价值——也使重用的元素
        GoTo Bail                                                           'Bail保释
      End If
    Next i                                                                  'Next message table entry下一个信息表条目
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
  End If
  
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem恢复z_ScMem的价值
End Sub

'Error handler错误处理程序
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure返回指定DLL的地址/过程
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address 获取指定程序的地址
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located 在IDE中,验证程序的地址
End Function

'Map zData() to the thunk address for the specified window handle 地图zData()的铛地址指定的窗口句柄
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started确保子类化已经启动
    zError "zMap_hWnd", "Subclassing hasn't been started"
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles抓住unsubclassed窗口句柄
    z_ScMem = z_Funk("h" & lng_hWnd)                                        'Get the thunk address“铛地址
    zMap_hWnd = z_ScMem
  End If
  
  Exit Function                                                             'Exit returning the thunk address退出返回铛的地址

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
'返回指定的地址顺序方法oCallback对象,1 =去年私有方法,2 =第二最后一个私有方法,等等
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry我们希望找到价值指向的vTable方法条目
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable地址的vTable
  Dim i     As Long                                                         'Loop index循环索引
  Dim j     As Long                                                         'Loop limit 循环的限制
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance获得地址回调对象的实例
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method类方法的调查
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method探究一种方法
      If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method调查用户控件的方法
        Exit Function                                                       'Bail... 保释……
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry 撞到下一个条目
  j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries设置一个合理的限制,扫描256 vTable条目
  Do While i < j
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry 获得地址存储在这个虚表条目
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address? 输入一个无效的代码地址吗?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address 返回指定的vTable入口地址
      Exit Do                                                               'Bad method signature, quit loop 坏方法签名,退出循环
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry 获得的字节指着vTable条目
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value... 如果字节不匹配预期的值…
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address 返回指定的vTable入口地址
      Exit Do                                                               'Bad method signature, quit loop 坏方法签名,退出循环
    End If
    
    i = i + 4                                                             '下一个虚表条目

  Loop
End Function

'Probe at the specified start address for a method signature指定起始地址探测方法签名
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address 起始地址
  nLimit = nAddr + 32                                                       'Probe eight entries调查八个条目
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth虽然我们没有达到探测深度
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry 得到vTable条目
    
    If nEntry <> 0 Then                                                     'If not an implemented interface如果没有一个实现接口
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry 获得的价值指向vTable条目
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature检查本机或pcode方法签名
        nMethod = nAddr                                                     'Store the vTable entry存储vTable条目
        bSub = bVal                                                         'Store the found method signature 存储方法签名
        zProbe = True                                                       'Indicate success 显示成功
        Exit Function                                                       'Return 返回
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry 下一个虚表条目
  Loop
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'-Subclass callback, usually ordinal #1, the last method in this source file----------------------子类的回调,通常序数# 1,最后一个方法在这个源文件-
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter

'* bBefore:表示之前或之后的回调是否最初的指向。通常
'*你就会知道,除非uMsg值被指定为回调
'* MSG_BEFORE_AFTER(之前和之后都原始的指向)。
'* bHandled——在原始指向回调之前,将bHandled设置为True会阻止
'*消息被传递到原始指向和(如果设置)之后
'*原始指向回调。
'* lReturn——指向返回值。设置为每MSDN文档信息的价值,
'*和/或在原指向回调后,返回值的设置
'*由原来的指向。
'* lng_hWnd -窗口句柄。
'* uMsg——消息值。
'*按钮——消息相关的数据。
'* lParam——消息相关的数据。
'* lParamUser——用户定义的回调参数


'*************************************************************************************************
  Dim tLayered As Boolean

  If uMsg = WM_WINDOWPOSCHANGED Then                                        'Select the message number 选择消息数量
    CopyMemory wp, ByVal lParam, Len(wp)                                    'Copy the WINDOWPOS data复制WINDOWPOS数据
    SizePos                                                                 'Position shadows位置的阴影

  ElseIf uMsg = WM_SIZE Then                                                'Parent form has been minimized/restored/maximized 父窗体最小化/恢复/最大化
    If wParam = SIZE_RESTORED Then                                          'If the parent has been restored 如果父母已经恢复
      If IsWindowVisible(hWndForm) = 1 Then                                 'If the parent is visible 如果父母是可见的
        Show True                                                           'Show shadows 显示阴影
      End If
    End If

  ElseIf uMsg = WM_DISPLAYCHANGE Then                                       'The display settings have been changed 显示设置已经改变了
    tLayered = bIsLayered                                                   'Save the the existing layered status 保存现有的分层状态
    CheckEnvironment                                                        'Check OS, color depth etc. 检查操作系统、颜色深度等。
    If tLayered Then                                                        'If we were layer enabled如果我们层启用
      If Not bIsLayered Then                                                'But not anymore但现在已经不是了
        DestroyWindow hWndBt                                                'Get rid of the shadow 摆脱阴影
        DestroyWindow hWndRt                                                'Get rid of the shadow摆脱阴影
        DestroyWindow hWndTt                                                'Get rid of the shadow摆脱阴影(AKEBE)
        DestroyWindow hWndLt                                                'Get rid of the shadow摆脱阴影(AKEBE)
      End If
     Else
      If bIsLayered Then                                                    'If we're now layered but previously weren't如果我们现在分层,但以前没有
        CreateWindows                                                       'Creates the shadow windows 创建阴影窗户
        DisplayShadows                                                      'Display the shadows 显示阴影
        If IsWindowVisible(hWndForm) Then                                   'If the parent is visible 如果父母是可见的
          Show True, True                                                   'Display the shadows 显示阴影
        End If
      End If
    End If

  ElseIf uMsg = WM_THEMECHANGED Then                                        'Windows XP theme has changed Windows XP主题已经改变了
    CheckEnvironment                                                        'Check OS, color depth etc.检查操作系统、颜色深度等。
    DisplayShadows                                                          'Display the shadows 显示阴影


  End If
  

End Sub


