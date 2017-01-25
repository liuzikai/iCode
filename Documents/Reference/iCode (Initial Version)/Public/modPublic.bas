Attribute VB_Name = "modPublic"
Option Explicit

Public Const WM_ACTIVATE = &H6
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COPY = &H301
Public Const WM_CREATE = &H1
Public Const WM_CUT = &H300
Public Const WM_ENABLE = &HA
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_KILLFOCUS = &H8
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOVE = &H3
Public Const WM_PAINT = &HF
Public Const WM_PASTE = &H302
Public Const WM_QUIT = &H12
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SHOWWINDOW = &H18
Public Const WM_SIZE = &H5
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SETFOCUS = &H7
Public Const WM_COMMAND = &H111



Public Type typControlArea
    hWnd As Long
    Left As Long
    Top As Long
    Width As Long
    Height As Long
    CodePxPerLine As Single
    bFixed As Boolean
End Type

Public VBIns As VBE
Public hVBIDE As Long
Public hCodeWnd As Long

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszCaption As String) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public iTipsBar As New clsTipsBarControler

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef pguid As guid) As Long

Private Type guid
    Data1 As Long
    Data4(0 To 7) As Byte
    Data3 As Integer
    Data2 As Integer
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long


Public Enum enm_iState
    isNone = 0
    isCode = 1
    isDesign = 2
End Enum

Public iState As enm_iState

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

Public ProjectPath As String

Public Type typSetting
    iTime As Date
    iVersion As String
    iNodes As Nodes
End Type

Public iProject As udProject
Public WndProject As Window

Public Declare Function GetLastError Lib "kernel32" () As Long

Public iCode As New clsCodeControler
Public iDesigner As clsDesignerControler

Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public frmHidden As New HiddenForm

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Public WindowsControler As New clsWindowsControler

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Public CodeOpe As New clsCodeOperator

Public iCodeMenu As CommandBarControl
Private iCodeMenuHandler As ButtonsCollection

Public Function OLEColorToRGB(ByVal nOLEColor As Long) As Long
    Dim i As Long
    OleTranslateColor nOLEColor, 0&, i
    OLEColorToRGB = i
End Function

Public Function RGBToARGB(ByVal nRGB As Long) As Long
    RGBToARGB = (-1 Xor &HFFFFFF) Or nRGB
End Function

Private Sub SetiCodeMenu()
    Set iCodeMenuHandler = New ButtonsCollection
    'iCodeMenuHandler类在需要使用时才进行实例化，可以减少不必要的空间占用
    Set iCodeMenuHandler.Target = VBIns.CommandBars("Add-Ins")
    Set iCodeMenu = iCodeMenuHandler.Add(msoControlPopup, "iCode", msoButtonCaption)
End Sub


Public Sub LoadPublicValues(objVBInstance As VBE)
    
    Set VBIns = objVBInstance '此语句优先级最高，VBIns是本程序直接操作VB IDE的通用接口
    
    hVBIDE = VBIns.MainWindow.hWnd '此语句优先级极高
    
    frmDebug.Show '此语句优先级极高，调试窗口需在执行调试输出之前进行加载
    
    
    SetiCodeMenu
    
    'iArea.CodePxPerLine = iArea.Height / VBIns.ActiveCodePane.CountOfVisibleLines
    
    
    'iCode.QDBLoadButtons
    
    'WindowsControler.Init
    
    Set frmHidden = New HiddenForm
    frmHidden.Visible = False
    
End Sub

Public Sub DBPrint(ByVal Str)
    frmDebug.txtMain.Text = frmDebug.txtMain.Text & Str & vbCrLf
    DoEvents
End Sub

Public Sub DBClear()
    frmDebug.txtMain.Text = ""
End Sub

Public Function GetGUID() As String
    Dim lRetVal As Long
    Dim udtGuid As guid
    
    Dim sPartOne As String
    Dim sPartTwo As String
    Dim sPartThree As String
    Dim sPartFour As String
    Dim iDataLen As Integer
    Dim iStrLen As Integer
    Dim iCtr As Integer
    Dim sAns As String
    
    On Error GoTo ErrorHandler
    
    sAns = ""
    
    lRetVal = CoCreateGuid(udtGuid)
    
    If lRetVal = 0 Then
        
        sPartOne = Hex$(udtGuid.Data1)
        iStrLen = Len(sPartOne)
        iDataLen = Len(udtGuid.Data1)
        sPartOne = String((iDataLen * 2) - iStrLen, "0") & Trim$(sPartOne)
        
        sPartTwo = Hex$(udtGuid.Data2)
        iStrLen = Len(sPartTwo)
        iDataLen = Len(udtGuid.Data2)
        sPartTwo = String((iDataLen * 2) - iStrLen, "0") & Trim$(sPartTwo)
        
        sPartThree = Hex$(udtGuid.Data3)
        iStrLen = Len(sPartThree)
        iDataLen = Len(udtGuid.Data3)
        sPartThree = String((iDataLen * 2) - iStrLen, "0") & Trim$(sPartThree)
        
        For iCtr = 0 To 7
            sPartFour = sPartFour & format$(Hex$(udtGuid.Data4(iCtr)), "00")
        Next
        
        sAns = sPartOne & sPartTwo & sPartThree & sPartFour
    End If
    
    GetGUID = sAns
    Exit Function
    
ErrorHandler:
    Exit Function
End Function

Public Function LeftIs(ByVal Str1 As String, ByVal Str2 As String) As Boolean
    LeftIs = (Left(LCase(Str1), Len(Str2)) = LCase(Str2))
End Function

Public Function SetControlSize(ByVal hParent As Long, ByVal sClassName As String, ByVal sCaption As String, ByVal Left, ByVal Top, ByVal Width, ByVal Height, Optional ByVal hNext As Long = 0) As Long
    Dim h As Long
    
    h = FindWindowEx(hParent, hNext, sClassName, sCaption)
    
    If h = 0 Then Exit Function
    
    If MoveWindow(h, Left, Top, Width, Height, True) = 0 Then Exit Function
    
    SetControlSize = h
End Function


Public Function GetControlRect(ByVal hWnd As Long, ByVal PHwnd As Long) As RECT
    Dim tRect As RECT
    Dim tPoint As POINTAPI
    
    GetWindowRect hWnd, tRect
    
    tPoint.X = tRect.Left
    tPoint.Y = tRect.Top
    
    ScreenToClient PHwnd, tPoint
    
    tRect.Left = tPoint.X
    tRect.Top = tPoint.Y
    
    tPoint.X = tRect.Right
    tPoint.Y = tRect.Bottom
    
    ScreenToClient PHwnd, tPoint
    
    tRect.Right = tPoint.X
    tRect.Bottom = tPoint.Y
    
    GetControlRect = tRect
End Function

Public Function GetControlArea(ByVal hWnd As Long, ByVal PHwnd As Long) As typControlArea
    Dim tRect As RECT
    tRect = GetControlRect(hWnd, PHwnd)
    GetControlArea.hWnd = hWnd
    GetControlArea.Left = tRect.Left
    GetControlArea.Top = tRect.Top
    GetControlArea.Width = tRect.Right - tRect.Left
    GetControlArea.Height = tRect.Bottom - tRect.Top
End Function

Public Function CheckControlArea(ByVal hWnd As Long, Optional ByVal Left, Optional ByVal Top, Optional ByVal Width, Optional ByVal Height) As Boolean
    Dim tArea As typControlArea
    tArea = GetControlArea(hWnd, GetParent(hWnd))
    If Not IsMissing(Left) Then If tArea.Left <> Left Then Exit Function
    If Not IsMissing(Top) Then If tArea.Top <> Top Then Exit Function
    If Not IsMissing(Width) Then If tArea.Width <> Width Then Exit Function
    If Not IsMissing(Height) Then If tArea.Height <> Height Then Exit Function
End Function

Public Function iGetClassName(ByVal hWnd As Long) As String
    On Error Resume Next
    
    Dim n As Long
    
    iGetClassName = Space$(255)
    n = GetClassName(hWnd, iGetClassName, 255)
    iGetClassName = Left$(iGetClassName, n)
End Function

Public Function iGetCaption(ByVal hWnd As Long) As String
    On Error Resume Next
    
    Dim n As Long
    
    iGetCaption = String$(255, Chr(0))
    n = GetWindowText(hWnd, iGetCaption, 255)
    iGetCaption = Left$(iGetCaption, InStr(1, iGetCaption, Chr(0)) - 1)
    
End Function

Public Sub AccelerationMotion(ByVal nStart As Single, ByVal nEnd As Single, ByVal nTimes As Long, ByVal MinVPerV0, ByRef mArray() As Single)
    
    ReDim mArray(nTimes)
    
    Dim mDis As Single, V0 As Single, a As Single, t As Long
    
    mDis = nEnd - nStart
    V0 = 2 * mDis / (nTimes) / (1 + MinVPerV0)
    a = (MinVPerV0 - 1) * V0 / (nTimes)
    
    mArray(1) = nStart
    
    For t = 2 To nTimes - 1
        mArray(t) = nStart + (V0 * t + (a * t * t) / 2)
    Next
    
    mArray(nTimes) = nEnd
End Sub

Public Sub AccelerationMotionDB(ByVal nStart As Single, ByVal nEnd As Single, ByVal nTimes As Long, ByVal MinVPerV0, ByRef mArray1() As Single, ByRef mArray2() As Single)
    AccelerationMotion nStart, nEnd, nTimes, MinVPerV0, mArray1
    AccelerationMotion nEnd, nStart, nTimes, MinVPerV0, mArray2
End Sub


Public Sub SetAlwaysOnTop(nHwnd As Long, bOnTop As Boolean)
    Dim rRect As RECT
    
    GetWindowRect nHwnd, rRect
    
    If bOnTop = True Then
        SetParent nHwnd, 0
        
        SetWindowPos nHwnd, HWND_TOPMOST, rRect.Left, rRect.Top, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Else
        Dim h
        
        h = FindWindow(vbNullString, "Program Manager")
        h = FindWindowEx(h, ByVal 0, "SHELLDLL_DefView", vbNullString)
        
        SetParent nHwnd, h
        
        SetWindowPos nHwnd, HWND_TOPMOST, rRect.Left, rRect.Top, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    End If
End Sub

Public Function iReplaceAll(ByVal Str As String, ByVal Find As String, ByVal Replace As String) As String
    
    If Str = "" Then Exit Function
    
    Dim a() As String
    a = Split(Str, Find)
    
    Dim r As String
    
    r = a(0)
    
    Dim i As Long
    For i = 1 To UBound(a)
        r = r & Replace & a(i)
    Next
    
    iReplaceAll = r
End Function
