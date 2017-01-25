VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIDEEnhancer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const DBFlag = "FunctionsSet::IDEEnhancer::"

Dim CommonDialog As CommonDialog
Dim MakeExeButton As CommandBarButton

Private iToolBar As CommandBar

'=== Params ===
Public ChangeScope_Button_Style As MsoButtonStyle
Public m_ChangeScope_Button_Visible As Boolean

Public m_Compile_Button_Visible As Boolean

Public ToCommon_Button_Style As MsoButtonStyle
Public m_ToCommon_Button_Visible As Boolean

'=== Buttons And Events ===

Private ToCommon_Button As CommandBarButton
Private WithEvents ToCommon_Button_Events As CommandBarEvents
Attribute ToCommon_Button_Events.VB_VarHelpID = -1

Private ChangeScope_Button As CommandBarButton
Private WithEvents ChangeScope_Button_Events As CommandBarEvents
Attribute ChangeScope_Button_Events.VB_VarHelpID = -1

Private MakeExe_Button As CommandBarButton
Private WithEvents MakeExe_Button_Events As CommandBarEvents
Attribute MakeExe_Button_Events.VB_VarHelpID = -1

Private Compile_Button As CommandBarButton
Private WithEvents Compile_Button_Events As CommandBarEvents
Attribute Compile_Button_Events.VB_VarHelpID = -1

Private AddFiles_Button1 As CommandBarButton
Private WithEvents AddFiles_Button1_Events As CommandBarEvents
Attribute AddFiles_Button1_Events.VB_VarHelpID = -1
Private AddFiles_Button2 As CommandBarButton
Private WithEvents AddFiles_Button2_Events As CommandBarEvents
Attribute AddFiles_Button2_Events.VB_VarHelpID = -1

Public Property Get ChangeScope_Button_Visible() As Boolean
    ChangeScope_Button_Visible = m_ChangeScope_Button_Visible
End Property

Public Property Let ChangeScope_Button_Visible(ByVal value As Boolean)
    m_ChangeScope_Button_Visible = value
    ChangeScope_Button.Visible = value
End Property

Public Property Get Compile_Button_Visible() As Boolean
    Compile_Button_Visible = m_Compile_Button_Visible
End Property

Public Property Let Compile_Button_Visible(ByVal value As Boolean)
    m_Compile_Button_Visible = value
    Compile_Button.Visible = value
End Property

Public Property Get ToCommon_Button_Visible() As Boolean
    ToCommon_Button_Visible = m_ToCommon_Button_Visible
End Property

Public Property Let ToCommon_Button_Visible(ByVal value As Boolean)
    m_ToCommon_Button_Visible = value
    ToCommon_Button.Visible = value
End Property

Private Sub Set_ToCommon_Button()
    Clipboard.SetData LoadResPicture(102, 0)
    Set ToCommon_Button = iToolBar.Controls.Add(MsoControlType.msoControlButton)
    With ToCommon_Button
        .Caption = "→通用"
        .ToolTipText = "返回代码“通用”"
        .PasteFace
        .Style = ToCommon_Button_Style
        .Visible = m_ToCommon_Button_Visible
    End With
    Set ToCommon_Button_Events = VBIns.Events.CommandBarEvents(ToCommon_Button)
End Sub

Private Sub Set_ChangeScope_Button()
    Clipboard.SetData LoadResPicture(101, 0)
    Set ChangeScope_Button = iToolBar.Controls.Add(MsoControlType.msoControlButton)
    With ChangeScope_Button
        .Caption = "Private←→Public"
        .ToolTipText = "互换当前行\选中内容\当前过程的“范围”标识"
        .PasteFace
        .Style = ChangeScope_Button_Style
        .Visible = m_ChangeScope_Button_Visible
    End With
    Set ChangeScope_Button_Events = VBIns.Events.CommandBarEvents(ChangeScope_Button)
End Sub

Private Function CSB_DealString(ByVal s As String) As String
    Do Until InStr(1, s, "Private ") = 0
        s = Replace(s, "Private ", "#3607#")
    Loop
    Do Until InStr(1, s, "Public ") = 0
        s = Replace(s, "Public ", "Private ")
    Loop
    Do Until InStr(1, s, "#3607#") = 0
        s = Replace(s, "#3607#", "Public ")
    Loop
    CSB_DealString = s
End Function


Private Sub Set_MakeExe_Button()
    
    Dim btn As CommandBarControl
    
    For Each btn In VBIns.CommandBars("File").Controls
        If InStr(1, btn.Caption, "生成") <> 0 Then Set MakeExeButton = btn
    Next
    
    Dim PWBar As CommandBar
    Set PWBar = VBIns.CommandBars.Item("Project Window Project")
    
    If Not (PWBar Is Nothing) Then
    
        Set MakeExe_Button = PWBar.Controls.Add(MsoControlType.msoControlButton, , , PWBar.Controls("打印(&P)...").Index, True)
        With MakeExe_Button
            .Caption = "编译生成(&K)..."
            .Style = msoButtonCaption
            .BeginGroup = True
        End With
        Set MakeExe_Button_Events = VBIns.Events.CommandBarEvents(MakeExe_Button)
    
    End If
    
End Sub

Private Sub GetCommandBars()
    Dim i As Long
    For i = 1 To VBIns.CommandBars.count
        DBPrint VBIns.CommandBars.Item(i).Name & "(" & VBIns.CommandBars.Item(i).Id & ")"
        Dim j As Long
        For j = 1 To VBIns.CommandBars.Item(i).Controls.count
            With VBIns.CommandBars.Item(i).Controls.Item(j)
                DBPrint "   " & .Caption & "(" & .Id & ")"
                If .Type = msoControlPopup Then GetPopUp .Controls, 2
            End With
        Next
    Next
End Sub

Private Sub Set_Compile_Button()
    Set Compile_Button = iToolBar.Controls.Add(MsoControlType.msoControlButton)
    With Compile_Button
        .Caption = "编译"
        .ToolTipText = "快速编译（覆盖上次编译文件）"
        .Style = msoButtonCaption
        .BeginGroup = True
        .Visible = m_Compile_Button_Visible
    End With
    Set Compile_Button_Events = VBIns.Events.CommandBarEvents(Compile_Button)
End Sub

Private Sub Call_Compile_Button()
    If Not (VBIns.ActiveVBProject Is Nothing) Then VBIns.ActiveVBProject.MakeCompiledFile
End Sub

Private Sub Set_AddFiles_Buttons()

    Set CommonDialog = frmContainer.CommonDialog1
    With CommonDialog
        .Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
        .Filter = _
            "VB 文件 (*.frm;*.ctl;*.dob;*.pag;*.dsr;*.bas;*.cls;*.res)|*.frm;*.ctl;*.dob;*.pag;*.dsr;*.bas;*.cls;*.res|" & _
            "窗体文件 (*.frm)|*.frm|" & _
            "模块文件 (*.bas)|*.bas|" & _
            "类文件 (*.cls)|*.cls|" & _
            "资源文件 (*.res)|*.res|" & _
            "用户控件文件 (*.ctl)|*.ctl|" & _
            "用户文档文件 (*.dob)|*.dob|" & _
            "属性页文件 (*.pag)|*.pag|" & _
            "设计器文件 (*.dsr)|*.dsr|" & _
            "所有文件 (*.*)|*.*"
    End With
    
    Dim Bar As CommandBar
    
    Set Bar = VBIns.CommandBars("Project Window Insert")
    Set AddFiles_Button1 = Bar.Controls.Add(MsoControlType.msoControlButton, , , Bar.Controls("添加文件(&A)...").Index, True)
    With AddFiles_Button1
        .Caption = "添加文件(增强版)(&P) ..."
        .ToolTipText = "添加文件(支持多个文件)"
        .Style = msoButtonCaption
    End With
    Set AddFiles_Button1_Events = VBIns.Events.CommandBarEvents(AddFiles_Button1)
    
    Set Bar = VBIns.CommandBars("Project")
    Set AddFiles_Button2 = Bar.Controls.Add(MsoControlType.msoControlButton, , , Bar.Controls("添加文件(&A)...").Index, True)
    With AddFiles_Button2
        .Caption = "添加文件(增强版)(&P) ..."
        .ToolTipText = "添加文件(支持多个文件)"
        .Style = msoButtonCaption
    End With
    Set AddFiles_Button2_Events = VBIns.Events.CommandBarEvents(AddFiles_Button2)
    
End Sub


Private Sub GetPopUp(C As CommandBarControls, ByVal D As Long)
    Dim i As Long
    For i = 1 To C.count
        With C.Item(i)
            DBPrint Space(D * 3) & .Caption & "(" & .Id & ")"
            If .Type = msoControlPopup Then GetPopUp .Controls, D
        End With
    Next
End Sub

Private Sub AddFiles_Button1_Events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    With CommonDialog
        .ShowOpen
        
        Dim s() As String
        s = Split(.FileName, Chr(0))
        
        If UBound(s) = 0 Then
            VBIns.ActiveVBProject.VBComponents.AddFile s(0)
        Else
            Dim i As Long
            For i = 1 To UBound(s)
                VBIns.ActiveVBProject.VBComponents.AddFile s(0) & "\" & s(i)
            Next
        End If
    End With
End Sub

Private Sub AddFiles_Button2_Events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    AddFiles_Button1_Events_Click CommandBarControl, handled, CancelDefault
End Sub

Public Sub Initialize(ByVal IDEIns As VBE, ByVal DBForm As Object, ByVal ToolBar As Object)
    
    Set VBIns = IDEIns
    Let hVBIDE = VBIns.MainWindow.hWnd
    
    Set DebugForm = DBForm
        
    Set iToolBar = ToolBar
    
    Set_MakeExe_Button
    Set_AddFiles_Buttons
    Set_ChangeScope_Button
    Set_ToCommon_Button
    Set_Compile_Button
    
End Sub

Public Sub Msg_ToExit()
    ChangeScope_Button_Style = ChangeScope_Button.Style
    m_ChangeScope_Button_Visible = ChangeScope_Button.Visible
    m_Compile_Button_Visible = Compile_Button.Visible
    ToCommon_Button_Style = ToCommon_Button.Style
    m_ToCommon_Button_Visible = ToCommon_Button.Visible
End Sub

Private Sub ChangeScope_Button_Events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    
    Dim s As String
    s = CodeOpe.Selection
    
    If s <> "" Then
        If CSB_DealString(s) <> s Then
            CodeOpe.Selection = CSB_DealString(s)
        End If
    Else
        Dim OSL As Long: OSL = CodeOpe.SL
        Dim OSC As Long: OSC = CodeOpe.SC
        If CodeOpe.bInDeclaration = True Then
            CodeOpe.ReplaceLine CodeOpe.SL, CSB_DealString(CodeOpe.Lines(CodeOpe.SL))
        Else
            Dim ProcName As String, ProcKind As vbext_ProcKind
            ProcName = CodeOpe.TargetCodePane.CodeModule.ProcOfLine(CodeOpe.SL, ProcKind)
            Dim ProcSL As Long
            ProcSL = CodeOpe.TargetCodePane.CodeModule.ProcBodyLine(ProcName, ProcKind)
            CodeOpe.ReplaceLine ProcSL, CSB_DealString(CodeOpe.Lines(ProcSL))
        End If
        CodeOpe.SetCaret OSL, OSC, OSL, OSC
    End If

End Sub

Private Sub Compile_Button_Events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error GoTo iErr
    If Not (VBIns.ActiveVBProject Is Nothing) Then
        VBIns.ActiveVBProject.MakeCompiledFile
    End If
    
    Exit Sub
iErr:
    MsgBox "错误！请通过 文件 - 编译 检查问题", vbExclamation, "Develop Tool"
End Sub

Private Sub MakeExe_Button_Events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error GoTo iErr
    If Not (MakeExeButton Is Nothing) Then MakeExeButton.Execute
    
    Exit Sub
iErr:
    MsgBox "错误！请通过 文件 - 编译 检查问题", vbExclamation, "Develop Tool"
End Sub

Private Sub ToCommon_Button_Events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    
    '注意：
    'CountOfDeclarationLines是声明部分的末行，不包含到第一个过程前的空行（属于第一个过程）
    
    '当没有声明部分时，第一个过程前的空行属于声明，但CountOfDeclarationLines为0
    '此时就需要跳转到第LinesCountBeforeProcs行（这部分空行的末行）
    
    '当存在声明部分时，第LinesCountBeforeProcs行可能不属于声明部分
    '此时就需要跳转到第CountOfDeclarationLines行
    
    Dim k As Long
    
    If CodeOpe.TargetCodePane.CodeModule.CountOfDeclarationLines = 0 Then
        If CodeOpe.LinesCountBeforeProcs = 0 Then
            CodeOpe.AddCodeToDeclaration "" '自带换行
            k = 1
        Else
            k = CodeOpe.LinesCountBeforeProcs
        End If
    Else
        k = CodeOpe.TargetCodePane.CodeModule.CountOfDeclarationLines
    End If
    
    CodeOpe.SetCaret k, Len(CodeOpe.Lines(k))
    
End Sub
