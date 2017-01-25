VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipsBarHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public hMDIClient As Long

Private WithEvents TipsBar As TipsBar
Attribute TipsBar.VB_VarHelpID = -1
Private hTipsBar As Long

Private C_Rect As RECT

Private TB_Height As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Const SW_MAX = 10
Private Const SW_NORMAL = 1
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CLOSE = &H10

Private Const WM_SIZE = &H5
Private Const WM_SETFOCUS = &H7
Private Const WM_DESTROY = &H2

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Dim frmContainer As Form

Private m_TipsBarAvliable As Boolean

Event MDIClientChangeSize()
Event ActiavteWindow(ByVal ClassName As String, ByVal Caption As String)
Event CloseWindow(ByVal ClassName As String, ByVal Caption As String)


Public Property Get MDIClient_Rect() As RECT
    MDIClient_Rect = C_Rect
End Property

Public Property Get TipsBarAvliable() As Boolean
    TipsBarAvliable = m_TipsBarAvliable
End Property

Public Property Let TipsBarAvliable(ByVal Value As Boolean)
    If m_TipsBarAvliable <> Value Then
        m_TipsBarAvliable = Value
        If Value Then
            SetParent hTipsBar, hVBIDE
            Msg_WM_SIZE_If_MDIClient
        Else
            SetParent hTipsBar, frmContainer.hWnd
            Dim r As RECT
            r = iGetRect(hMDIClient, hVBIDE)
            r.Top = r.Top - TB_Height
            MoveWindow hMDIClient, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, True
        End If
    End If
End Property

Public Sub Msg_WM_SIZE_If_MDIClient()
    If m_TipsBarAvliable Then
        Dim r As RECT
        r = iGetRect(hMDIClient, hVBIDE)
        'DBPrint "r.top=" & r.Top & ",bottom=" & r.Bottom
        'DBPrint "C_Rect.top=" & C_Rect.Top & ",bottom=" & C_Rect.Bottom
        If r.Top <> C_Rect.Top Or r.Left <> C_Rect.Left Or r.Bottom <> C_Rect.Bottom Or r.Right <> C_Rect.Right Then
            r.Top = r.Top + TB_Height
            C_Rect = r
            MoveWindow hMDIClient, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, True
            MoveWindow hTipsBar, r.Left, r.Top - TB_Height, r.Right - r.Left, TB_Height, True
            RaiseEvent MDIClientChangeSize
        End If
    End If
End Sub


Public Sub Msg_WM_MDIDESTROY_If_IDEWindow(ByVal hWnd As Long)
    
    Dim n As Long
    
    n = TipsBar.FindIDByKey(hWnd)
    
    If n <> -1 Then
        TipsBar.Remove n
        RaiseEvent CloseWindow(iGetClassName(hWnd), iGetCaption(hWnd))
    End If
    
End Sub

Public Sub Msg_WM_MDIACTIVATE_If_IDEWindow(ByVal hWnd As Long, ByVal ClassName As String)

SubActivate:
    
    Dim n As Long
    n = TipsBar.FindIDByKey(hWnd)
    If n <> -1 Then
        TipsBar.Activate n
        RaiseEvent ActiavteWindow(ClassName, iGetCaption(hWnd))
        Exit Sub
    End If

SubJudge:
        
    Dim s As String: s = iGetCaption(hWnd)
    Dim sFinal As String
    Dim i1 As Long, i2 As Long, i3 As Long
    Dim sP As String, sC As String, sK As String
        
    Select Case ClassName
    
    Case "VbaWindow", "DesignerWindow"
        i1 = InStr(1, s, "-")
        i2 = InStr(1, s, "(")
        i3 = InStr(1, s, ")")
        
        If i1 = 0 Or i2 = 0 Or i3 = 0 Then Exit Sub
        
        sP = Trim(Mid(s, 1, i1 - 1))
        sC = Trim(Mid(s, i1 + 1, i2 - (i1 + 1)))
        sK = Trim(Mid(s, i2 + 1, i3 - (i2 + 1)))
                
        'If VBIns.VBProjects.Item(sP) Is Nothing Then Exit Sub
        'If VBIns.VBProjects.Item(sP).VBComponents.Item(sC) Is Nothing Then Exit Sub
        
        sFinal = sC
        
        Select Case sK
        Case "Code"
            'Do Nothing
        Case Else
            sFinal = sFinal & "[设计]"
        End Select
        
        RaiseEvent ActiavteWindow(ClassName, s)
        
    Case "DockingView"
        'If s = "对象浏览器" Then
            sFinal = s
            RaiseEvent ActiavteWindow(ClassName, s)
        'Else
        '    Exit Sub
        'End If
    Case Else
        Exit Sub
    End Select
    
SubAddTip:
    TipsBar.Add sFinal, hWnd
    
    Exit Sub

End Sub

Private Function iGetRect(ByVal hWnd As Long, ByVal hParent As Long) As RECT
    Dim r As RECT
    GetWindowRect hWnd, r
    
    Dim pt As POINTAPI
    pt.x = r.Left
    pt.y = r.Top
    
    ScreenToClient hParent, pt
    
    r.Left = pt.x
    r.Top = pt.y
    
    pt.x = r.Right
    pt.y = r.Bottom
    
    ScreenToClient hParent, pt
    
    r.Right = pt.x
    r.Bottom = pt.y
    
    iGetRect = r
End Function

Public Sub Initialize(ByVal IDEIns As VBE, ByVal DBForm As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode)
    
    Set VBIns = IDEIns
    hVBIDE = VBIns.MainWindow.hWnd
    
    Set DebugForm = DBForm
    
    Set frmContainer = New frmContainerX
    frmContainer.Show
    frmContainer.Visible = False
    
    Set TipsBar = frmContainer.TipsBar1
    TB_Height = frmContainer.TipsBar1.Height
    hTipsBar = TipsBar.hWnd

    hMDIClient = FindWindowEx(hVBIDE, 0, "MDIClient", "")
    If hMDIClient = 0 Then DBPrint "clsTBH:Fail To Find MDIClient"
    
    If ConnectMode <> vbext_cm_Startup Then LoadWindows
    
End Sub

Private Sub LoadWindows()
    Dim hWnd As Long
    hWnd = FindWindowEx(hMDIClient, 0&, vbNullString, vbNullString)
    Do Until hWnd = 0
        Dim ClassName As String
        ClassName = iGetClassName(hWnd)
        Select Case ClassName
        Case "VbaWindow", "DesignerWindow", "DockingView"
            Msg_WM_MDIACTIVATE_If_IDEWindow hWnd, ClassName
        End Select
        hWnd = FindWindowEx(hMDIClient, hWnd, vbNullString, vbNullString)
    Loop
End Sub

Private Sub Class_Terminate()
    
    'DBPrint "Exit: Start Terminate"

    TipsBarAvliable = False
    
    'DBPrint "Exit: SetBack Succeed!"
    
    Set TipsBar = Nothing
    
    'DBPrint "Exit: UnLoad TipsBar Succeed!"
    
End Sub



Private Sub TipsBar_mnuCengDie()
    On Error Resume Next
    VBIns.CommandBars("Window").Controls("层叠(&C)").Execute
End Sub

Private Sub TipsBar_mnuShuiPing()
    On Error Resume Next
    VBIns.CommandBars("Window").Controls("水平平铺(&H)").Execute
End Sub

Private Sub TipsBar_mnuChuiZhi()
    On Error Resume Next
    VBIns.CommandBars("Window").Controls("垂直平铺(&V)").Execute
End Sub

Private Sub TipsBar_TipClick(ByVal ID As Long)
    SetFocus TipsBar.Key(ID)
End Sub

Private Sub TipsBar_TipClose(ByVal ID As Long)
    SendMessage TipsBar.Key(ID), WM_CLOSE, 0, 0
    DoEvents
End Sub
