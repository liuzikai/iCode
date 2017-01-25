VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeWindowEnhance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents btnToCommand As CommandButton
Attribute btnToCommand.VB_VarHelpID = -1
Private WithEvents btnBack As CommandButton
Attribute btnBack.VB_VarHelpID = -1
Private WithEvents btnGo As CommandButton
Attribute btnGo.VB_VarHelpID = -1

Private Const CB_SETCURSEL = &H14E
Private Const CB_SHOWDROPDOWN = &H14F
Private Const WM_SETFOCUS = &H7
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_KILLFOCUS = &H8

Private Const cControlDisX As Long = 7

Public hButtonsContainer As Long
Private c1&, c2&
Private c1Rect As RECT, c2Rect As RECT
Private hcb1&, hcb2&, hcb3&

Public Type typProcInfo
    ProcName As String
    ProcKind As vbext_ProcKind
End Type

Private Const ProcNotesMaxCount As Long = 7
Private ProcNotes(ProcNotesMaxCount + 1) As typProcInfo
Private m_ProcNoteTarget As Long
Public ProcNotesRollingBack As Boolean



Private KeyNote As String

Private WithEvents iAPIHooker As CVBAPIHooker
Attribute iAPIHooker.VB_VarHelpID = -1
Private WithEvents m_ErrBox As EH_ErrorBox
Attribute m_ErrBox.VB_VarHelpID = -1
Private lngHookIndex As Long

Public Type MSGBOXPARAMS
    cbSize As Long
    hwndOwner As Long
    hInstance As Long
    lpszText As String
    lpszCaption As String
    dwStyle As Long
    lpszIcon As String
    dwContextHelpId As Long
    lpfnMsgBoxCallback As Long
    dwLanguageId As Long
End Type

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWngCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long

Private Const hh_display_topic = &H0                                            '上下文相关的帮助(导航窗口)
Private Const hh_display_index = &H2                                            '索引窗口
Private Const hh_help_context = &HF                                             '上下文相关的帮助
Private Const hh_display_search = &H3                                           '搜索窗口
Private Const hh_display_text_popup = &HE                                       '上下文相关的帮助

Private WithEvents timerWait As XTimer
Attribute timerWait.VB_VarHelpID = -1
Private WithEvents timerPNWait As XTimer
Attribute timerPNWait.VB_VarHelpID = -1

Friend Property Get ErrBox() As EH_ErrorBox
    Set ErrBox = m_ErrBox
End Property

Friend Property Set ErrBox(NewValue As EH_ErrorBox)
    Set m_ErrBox = NewValue
End Property

Public Sub EHHookMessageBox()
    Set iAPIHooker = New CVBAPIHooker
    Set m_ErrBox = New EH_ErrorBox
    m_ErrBox.Visible = False
    
    iAPIHooker.SetTargetTo -1
    lngHookIndex = iAPIHooker.Hook("VB6.EXE", "user32.dll", "MessageBoxIndirectA")
End Sub

Public Sub EHUnHookMessageBox()
    iAPIHooker.Unhook lngHookIndex
    
    Set iAPIHooker = Nothing
End Sub

Public Function EHGetMessageBoxInfo(ByVal lpArgs As Long) As MSGBOXPARAMS
    Dim Res(10) As Long
    
    Dim i As Long
    
    For i = 1 To 10
        'MessageBoxIndirectA真正的参数地址从APIHooker提供的参数堆栈地址之后16个字节开始
        
        CopyMemory Res(i), ByVal lpArgs + 4 * (i - 1) + 16, 4
    Next
    
    With EHGetMessageBoxInfo
        .cbSize = Res(1)
        .hwndOwner = Res(2)
        .hInstance = Res(3)
        .lpszText = iAPIHooker.iGetStrFromPtr(Res(4))
        .lpszCaption = iAPIHooker.iGetStrFromPtr(Res(5))
        .dwStyle = Res(6)
        .lpszIcon = iAPIHooker.iGetStrFromPtr(Res(7))
        .dwContextHelpId = Res(8)
        .lpfnMsgBoxCallback = Res(9)
        .dwLanguageId = Res(10)
    End With
    
End Function


Private Property Get ProcNoteTarget() As Long
    ProcNoteTarget = m_ProcNoteTarget
End Property

Private Property Let ProcNoteTarget(NewValue As Long)
    m_ProcNoteTarget = NewValue
    btnBack.Enabled = (NewValue > 1)
    btnGo.Enabled = ((NewValue <= ProcNotesMaxCount) And (ProcNotes(NewValue).ProcName <> ""))
End Property

Public Sub ProcNotesAdd(ByVal ProcName As String, ByVal ProcKind As vbext_ProcKind)
    
    If ProcNotesRollingBack Then Exit Sub
    
    If ProcName = "" Then ProcName = "Declaration"
    
    If ProcName = ProcNotes(ProcNoteTarget).ProcName And ProcKind = ProcNotes(ProcNoteTarget).ProcKind Then Exit Sub
    
    Dim i As Long
    If ProcNoteTarget >= ProcNotesMaxCount Then
        For i = 1 To ProcNotesMaxCount - 1
            ProcNotes(i) = ProcNotes(i + 1)
        Next
        ProcNotes(ProcNotesMaxCount).ProcName = ProcName
        ProcNotes(ProcNotesMaxCount).ProcKind = ProcKind
    ElseIf ProcNoteTarget < ProcNotesMaxCount Then
        ProcNoteTarget = ProcNoteTarget + 1
        
        ProcNotes(ProcNoteTarget).ProcName = ProcName
        ProcNotes(ProcNoteTarget).ProcKind = ProcKind
        
        For i = ProcNoteTarget + 1 To ProcNotesMaxCount
            ProcNotes(i).ProcName = ""
            ProcNotes(i).ProcKind = 0
        Next
    End If
End Sub

Public Sub ProcNotesBack()
    ProcNoteTarget = ProcNoteTarget - 1
    ProcNotesGoto ProcNoteTarget
End Sub

Public Sub ProcNotesGo()
    ProcNoteTarget = ProcNoteTarget + 1
    ProcNotesGoto ProcNoteTarget
End Sub


Public Sub ProcNotesGoto(ByVal NoteIndex As Long)
    On Error Resume Next
    
    Dim n As Long
    
    ProcNotesRollingBack = True
    
    If ProcNotes(NoteIndex).ProcName = "Declaration" Then
        n = 1
    Else
        n = VBIns.ActiveCodePane.CodeModule.ProcBodyLine(ProcNotes(NoteIndex).ProcName, ProcNotes(NoteIndex).ProcKind) + 1
        '这里如果找不到过程会出现错误,因此用On Error Resume Next,则n会为1
        
        If n = 1 Then
            If NoteIndex >= 2 Then
                ProcNotesBack
            Else
                ProcNoteTarget = ProcNoteTarget - 1
            End If
            Exit Sub
        End If
        
    End If
    
    VBIns.ActiveCodePane.SetSelection n, 1, n, 1
    
    ProcNotesRollingBack = False
    
End Sub

Public Sub ProcNotesNewCodePane()
    
    Erase ProcNotes
    ProcNoteTarget = 0
    iCode.CodeSort.lProc = 0
    
    Dim n As Long
    n = CodeOpe.GetActiveLine
    
    Dim s As String, k As vbext_ProcKind
    
    s = VBIns.ActiveCodePane.CodeModule.ProcOfLine(n, k)
    
    ProcNotesAdd s, k
    
End Sub

Public Sub DealMessage_NewCodePane()
    timerWait.Start
    timerPNWait.Start
End Sub


Public Sub DealMessage(ByVal Msg As Long, ByVal hWnd As Long, ByVal Caption As String, ByVal ClassName As String)
    
    Select Case Msg
        
    Case HCBT_MOVESIZE, HCBT_MINMAX
        
        Select Case ClassName
        Case "VbaWindow"
            If Caption = "对象浏览器" Or Caption = "立即" Then Exit Sub
            timerWait.Start                                                     'ECW延时器启动
        End Select
        
    Case HCBT_DESTROYWND
        
        Select Case ClassName
            
        Case "VbaWindow"
            
            If Caption = "对象浏览器" Or Caption = "立即" Then Exit Sub
            
            iCode.CWE.ECWSetButtonBack                                          '在代码窗格销毁前将按钮迁回
            
        End Select
        
    End Select
    
End Sub

Public Sub ECWLoadButton()
    Set btnToCommand = frmHidden.CodeToCommand
    hcb1 = frmHidden.cCodeToCommand.hWnd
    Set btnBack = frmHidden.CodeBack
    hcb2 = frmHidden.cCodeBack.hWnd
    Set btnGo = frmHidden.CodeGo
    hcb3 = frmHidden.cCodeGo.hWnd
    hButtonsContainer = frmHidden.hWnd
End Sub

Public Sub ECWSetButtonBack()
    DBPrint "ECWSetButtonBack"
    SetParent hcb1, frmHidden.hWnd
    SetParent hcb2, frmHidden.hWnd
    SetParent hcb3, frmHidden.hWnd
    hButtonsContainer = frmHidden.hWnd
End Sub

Public Function ECWEnhance(ByVal hWnd As Long) As Boolean
    
    c1 = FindWindowEx(hWnd, 0, "ComboBox", "")
    c2 = FindWindowEx(hWnd, c1, "ComboBox", "")
    
    If c1 = 0 Or c2 = 0 Then Exit Function
    
    Dim r1 As RECT, r2 As RECT
    
    r1 = GetControlRect(c1, hWnd)
    r2 = GetControlRect(c2, hWnd)
    
    
    If r1.Left > r2.Left Then
        Dim th As Long, tr As RECT
        th = c1:  c1 = c2:  c2 = th
        
        tr = r1:  r1 = r2:  r2 = tr
    End If
    
    If r1.Right <> c1Rect.Right Or r2.Right <> c2Rect.Right Then
        
        With frmHidden
            
            MoveWindow hcb1, r1.Left, r1.Top, .CodeToCommand.Width, .CodeToCommand.Height, True
            
            Dim wComboBox As Long
            
            wComboBox = (r2.Right - r1.Left - .CodeToCommand.Width - cControlDisX - cControlDisX - .CodeBack.Width - cControlDisX - .CodeGo.Width - cControlDisX) / 2
            
            MoveWindow c1, r1.Left + .CodeToCommand.Width + cControlDisX, r1.Top, wComboBox, r1.Bottom - r1.Top, True
            
            MoveWindow hcb2, r1.Left + .CodeToCommand.Width + cControlDisX + wComboBox + cControlDisX, r1.Top, .CodeBack.Width, .CodeBack.Height, True
            MoveWindow hcb3, r1.Left + .CodeToCommand.Width + cControlDisX + wComboBox + cControlDisX + .CodeBack.Width + cControlDisX, r1.Top, .CodeGo.Width, .CodeGo.Height, True
            
            MoveWindow c2, r1.Left + .CodeToCommand.Width + cControlDisX + wComboBox + cControlDisX + .CodeBack.Width + cControlDisX + .CodeGo.Width + cControlDisX, r2.Top, wComboBox, r2.Bottom - r2.Top, True
        End With
        
    End If
    
    If hButtonsContainer <> hWnd Then
        SetParent hcb1, hWnd
        SetParent hcb2, hWnd
        SetParent hcb3, hWnd
        hButtonsContainer = hWnd
    End If
    
    c1Rect = GetControlRect(c1, hWnd)
    c2Rect = GetControlRect(c2, hWnd)
    
End Function



Public Sub ECWSetToCommand(ByVal hWnd As Long)
    SendMessage hWnd, WM_SETFOCUS, 0&, 0&
    SendMessage hWnd, CB_SETCURSEL, 0, 0
    SendMessage hWnd, WM_KEYDOWN, vbKeyReturn, 0&
    SendMessage hWnd, WM_KEYUP, vbKeyReturn, 0&
    SendMessage hWnd, WM_KILLFOCUS, 0&, 0&
End Sub

Public Sub ECWShowPopUp(ByVal hWnd As Long)
    SendMessage hWnd, WM_SETFOCUS, 0&, 0&
    SendMessage hWnd, CB_SHOWDROPDOWN, True, 0
End Sub


Private Sub btnBack_Click()
    ProcNotesBack
    SetFocus hButtonsContainer
End Sub

Private Sub btnGo_Click()
    ProcNotesGo
    SetFocus hButtonsContainer
End Sub

Private Sub btnToCommand_Click()
    Me.ECWSetToCommand c1
End Sub

Private Sub Class_Initialize()
    Set timerWait = New XTimer
    timerWait.Enabled = False
    timerWait.Interval = 100
    Set timerPNWait = New XTimer
    timerPNWait.Enabled = False
    timerPNWait.Interval = 100
End Sub

Private Sub iAPIHooker_BeforeAction(ByVal sAPIName As String, ByVal lpArgs As Long, fBlock As Boolean, lArgCount As Long)
    Dim mbArgs As MSGBOXPARAMS
    
    mbArgs = EHGetMessageBoxInfo(lpArgs)
    
    With mbArgs
        If .lpszCaption = "Microsoft Visual Basic" And .dwStyle = 16432 And InStr(1, .lpszText, "错误") <> 0 Then
            
            fBlock = True
            
            Dim t As POINTAPI
            
            t = CodeOpe.GetCaretPoint
            
            ErrBox.Left = ErrBox.ScaleX(t.X + 10, 3, 1)
            ErrBox.Top = ErrBox.ScaleX(t.Y + 10, 3, 1)
            
            ErrBox.ShowError .lpszText, .dwContextHelpId
            
            HiddenForm.EH_timerErrorBox.Enabled = True
            
        Else
            fBlock = False
        End If
    End With
    
End Sub

Private Sub m_ErrBox_PressHelp(ByVal sHelpFile As String, ByVal nHelpContext As Long)
    sHelpFile = "E:\MSDN Library\98VS\2052\msdnvs98.col"
    HtmlHelp hVBIDE, sHelpFile, hh_display_topic, nHelpContext
End Sub

Private Sub m_ErrBox_PressOK()
    m_ErrBox.Visible = False
End Sub

Private Sub timerPNWait_Tick()
    Me.ProcNotesNewCodePane                                                     '触发ProcNotes的进入新窗格时的初始化动作
    timerPNWait.Pause
End Sub

Private Sub timerWait_Tick()
    Me.ECWEnhance hCodeWnd                                                      '触发ECW的代码窗格强化动作
    timerWait.Pause
End Sub
