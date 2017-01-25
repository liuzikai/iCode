VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindowsControler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const WM_SETTEXT = &HC
Private Const EM_SETSEL = &HB1

Private WithEvents OV_Timer As XTimer
Attribute OV_Timer.VB_VarHelpID = -1

Dim m_OV_SelText As String
Dim OV_Hwnd As Long

Public Sub Msg(ByVal hWnd As Long, ByVal Caption As String)
    DBPrint Caption
    Select Case Caption
    Case "生成工程"
        DBPrint Caption & " - Enter"
        SetBuildWindow hWnd
    Case "打开工程"
        SetOpenWindow hWnd
    Case "文件另存为", "工程另存为"
        SetSaveWindow hWnd
    Case "添加文件"
        SetAddFileWindow hWnd
    Case "添加引用", "添加 ActiveX 控件"
        SetAddPartWindow hWnd
    Case "对象浏览器"
        If m_OV_SelText <> "" Then
            OV_Hwnd = hWnd
            OV_Timer.Enabled = True
        End If
    End Select
    
End Sub

Public Property Let OV_SelText(ByVal Value As String)
    m_OV_SelText = Value
End Property

Private Function SetObjectViewer(ByVal hObjectViewer As Long) As Boolean
    
    Dim h As Long, h0 As Long, h1 As Long, h2 As Long
    
    SetObjectViewer = False
    
    h0 = FindWindowEx(hObjectViewer, 0, "VbaWindow", "对象浏览器")
    h1 = FindWindowEx(h0, 0, "ToolBarWindow32", "")
    h2 = FindWindowEx(h1, 0, "ComboBox", "")
    h = FindWindowEx(h2, 0, "Edit", "")
    
    If h = 0 Then
        h1 = FindWindowEx(h0, h1, "ToolBarWindow32", "")
        h2 = FindWindowEx(h1, 0, "ComboBox", "")
        h = FindWindowEx(h2, 0, "Edit", "")
    End If
    
    If h = 0 Then Exit Function
    
    If SendMessage(h, WM_SETTEXT, 0, ByVal m_OV_SelText) <> 1 Then Exit Function
    SendMessage h, EM_SETSEL, 0, -1
    
    m_OV_SelText = ""
    
    SetObjectViewer = True
End Function

Public Function SetBuildWindow(ByVal hWnd As Long) As Boolean
    
    SetBuildWindow = False
    
    Dim hArea As Long
    
    hArea = SetControlSize(hWnd, "#32770", "", 0, 0, 614, 439)
    
    SetControlSize hWnd, "Static", "保存在(&I):", 11, 10, 72, 14
    SetControlSize hWnd, "ComboBox", "", 86, 5, 420, 24
    SetControlSize hWnd, "ToolbarWindow32", "", 513, 5, 132, 26
    SetControlSize hWnd, "Static", "文件名(&N):", 8, 386, 66, 14
    SetControlSize hWnd, "Edit", "", 89, 383, 350, 20
    SetControlSize hWnd, "Button", "确定", 541, 380, 75, 23
    SetControlSize hWnd, "Button", "取消", 541, 408, 75, 23
    SetControlSize hWnd, "Button", "帮助(&H)", 6, 408, 70, 23
    SetControlSize hArea, "Button", "选项(&O)...", 450, 380, 77, 23
    SetControlSize hWnd, "SHELLDLL_DefView", "", 6, 36, 610, 337
    
    If CheckControlArea(hWnd, , , 630, 467) = False Then MoveWindow hWnd, 263, 136, 630, 467, True
    
    SetBuildWindow = True
    
End Function

Public Function SetOpenWindow(ByVal hWnd As Long) As Boolean
    SetOpenWindow = False
    
    Dim hComboBox As Long
    
    SetControlSize hWnd, "Static", "查找范围(&I):", 21, 36, 72, 14
    
    hComboBox = SetControlSize(hWnd, "ComboBox", "", 96, 31, 420, 24)
    
    SetControlSize hWnd, "ToolbarWindow32", "", 523, 31, 100, 26
    
    SetControlSize hWnd, "Static", "文件名(&N):", 18, 412, 66, 14
    SetControlSize hWnd, "Edit", "", 99, 409, 437, 20
    
    SetControlSize hWnd, "Static", "文件类型(&T):", 18, 440, 80, 14
    SetControlSize hWnd, "ComboBox", "", 99, 437, 350, 20, hComboBox
    
    SetControlSize hWnd, "Button", "打开(&O)", 551, 406, 75, 23
    SetControlSize hWnd, "Button", "取消", 551, 434, 75, 23
    SetControlSize hWnd, "Button", "帮助(&H)", 461, 434, 75, 23
    
    SetControlSize hWnd, "SHELLDLL_DefView", "", 16, 62, 610, 337
    
    SetControlSize hWnd, "SysListView32", "", 16, 32, 610, 367
    
    SetControlSize hWnd, "SysTabControl32", "", 6, 6, 634, 469
    
    If CheckControlArea(hWnd, , , 652, 523) = False Then MoveWindow hWnd, 258, 123, 652, 523, True
    
    SetOpenWindow = True
End Function

Public Function SetSaveWindow(ByVal hWnd As Long) As Boolean
    
    SetSaveWindow = False
    
    Dim hComboBox As Long
    
    SetControlSize hWnd, "Static", "保存在(&I):", 11, 10, 72, 14
    
    hComboBox = SetControlSize(hWnd, "ComboBox", "", 86, 5, 420, 24)
    
    SetControlSize hWnd, "ToolbarWindow32", "", 513, 5, 100, 26
    
    SetControlSize hWnd, "Static", "文件名(&N):", 8, 386, 66, 14
    SetControlSize hWnd, "Edit", "", 99 - 10, 409 - 26, 437, 20
    
    SetControlSize hWnd, "Static", "保存类型(&T):", 8, 414, 80, 14
    SetControlSize hWnd, "ComboBox", "", 89, 411, 350, 20, hComboBox
    
    SetControlSize hWnd, "Button", "保存(&S)", 541, 380, 75, 23
    SetControlSize hWnd, "Button", "取消", 541, 408, 75, 23
    SetControlSize hWnd, "Button", "帮助(&H)", 451, 408, 75, 23
    
    SetControlSize hWnd, "SHELLDLL_DefView", "", 6, 36, 610, 337
    
    If CheckControlArea(hWnd, , , 630, 467) = False Then MoveWindow hWnd, 263, 136, 630, 467, True
    
    SetSaveWindow = True
    
End Function

Public Function SetAddFileWindow(ByVal hWnd As Long) As Boolean
    
    SetAddFileWindow = False
    
    Dim hArea As Long
    Dim hComboBox As Long
    
    hArea = SetControlSize(hWnd, "#32770", "", 0, 0, 614, 466)
    
    SetControlSize hWnd, "Static", "查找范围(&I):", 11, 10, 72, 14
    
    hComboBox = SetControlSize(hWnd, "ComboBox", "", 86, 5, 420, 24)
    
    SetControlSize hWnd, "ToolbarWindow32", "", 513, 5, 100, 26
    
    SetControlSize hWnd, "Static", "文件名(&N):", 8, 386, 66, 14
    SetControlSize hWnd, "Edit", "", 89, 383, 437, 20
    
    SetControlSize hWnd, "Static", "文件类型(&T):", 8, 414, 80, 14
    SetControlSize hWnd, "ComboBox", "", 89, 411, 350, 20, hComboBox
    
    SetControlSize hWnd, "Button", "打开(&O)", 541, 380, 75, 23
    SetControlSize hWnd, "Button", "取消", 541, 408, 75, 23
    SetControlSize hWnd, "Button", "帮助(&H)", 451, 408, 75, 23
    
    SetControlSize hArea, "Button", "添加为相关文档", 8, 440, 108, 15
    
    SetControlSize hWnd, "SHELLDLL_DefView", "", 6, 36, 610, 337
    
    If CheckControlArea(hWnd, , , 630, 494) = False Then MoveWindow hWnd, 263, 123, 630, 494, True
    
    SetAddFileWindow = True
    
End Function

Public Function SetAddPartWindow(ByVal hWnd As Long) As Boolean
    
    SetAddPartWindow = False
    
    Dim hComboBox As Long
    
    SetControlSize hWnd, "Static", "查找范围(&I):", 11, 10, 72, 14
    
    hComboBox = SetControlSize(hWnd, "ComboBox", "", 86, 5, 420, 24)
    
    SetControlSize hWnd, "ToolbarWindow32", "", 513, 5, 100, 26
    
    SetControlSize hWnd, "Static", "文件名(&N):", 8, 386, 66, 14
    SetControlSize hWnd, "Edit", "", 99 - 10, 409 - 26, 437, 20
    
    SetControlSize hWnd, "Static", "文件类型(&T):", 8, 414, 80, 14
    SetControlSize hWnd, "ComboBox", "", 89, 411, 350, 20, hComboBox
    
    SetControlSize hWnd, "Button", "打开(&O)", 541, 380, 75, 23
    SetControlSize hWnd, "Button", "取消", 541, 408, 75, 23
    SetControlSize hWnd, "Button", "帮助(&H)", 451, 408, 75, 23
    
    SetControlSize hWnd, "SHELLDLL_DefView", "", 6, 36, 610, 337
    
    If CheckControlArea(hWnd, , , 630, 467) = False Then MoveWindow hWnd, 263, 136, 630, 467, True
    
    SetAddPartWindow = True
End Function

Public Sub Init()
    Set OV_Timer = New XTimer
    OV_Timer.Interval = 200
    OV_Timer.Enabled = True
End Sub

Private Sub Class_Terminate()
    Set OV_Timer = Nothing
End Sub

Private Sub OV_Timer_Timer()
    OV_Timer.Enabled = Not SetObjectViewer(OV_Hwnd)
End Sub
