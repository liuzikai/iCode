VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeIndent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum CI_Structure
    IF_Start = 11
    IF_Middle = 31
    IF_End = 21
    Select_Start = 12
    Select_Middle = 32
    Select_End = 22
    With_Start = 13
    With_End = 23
    For_Start = 14
    For_End = 24
    Proc_Start = 15
    Proc_End = 25
    Type_Start = 16
    Type_End = 26
    Enum_Start = 17
    Enum_End = 27
    Do_Start = 18
    Do_End = 28
    While_Start = 19
    Wend_End = 29
End Enum

Private iSL As Long, iSC As Long, iEL As Long, iEC As Long

Dim Stack As New CStack

Public Var_SpacePerLevel As Long

Public Var_MulitLines_IndentSpaceCount As Long
'-1为自动匹配(尚未编写)，0为不整理，>0为默认缩进量
Public Var_ResultWindow_AutoHide As Boolean

Private m_Var_QuickButtonMode As Long

Private iCodeMenu As CommandBarPopup
Private WithEvents iCodeMenuEvent As CommandBarEvents
Attribute iCodeMenuEvent.VB_VarHelpID = -1
Private iToolBar As CommandBar
Private iQuickButton As CommandBarButton
Private WithEvents iQuickButtonEvent As CommandBarEvents
Attribute iQuickButtonEvent.VB_VarHelpID = -1

Private popIndent As CommandBarPopup
Attribute popIndent.VB_VarHelpID = -1
Private WithEvents btnIndentProjectEvent As CommandBarEvents
Attribute btnIndentProjectEvent.VB_VarHelpID = -1
Private WithEvents btnIndentModuleEvent As CommandBarEvents
Attribute btnIndentModuleEvent.VB_VarHelpID = -1
Private WithEvents btnIndentProcEvent As CommandBarEvents
Attribute btnIndentProcEvent.VB_VarHelpID = -1
Private btnErrorWindow As CommandBarButton
Private WithEvents btnErrorWindowEvent As CommandBarEvents
Attribute btnErrorWindowEvent.VB_VarHelpID = -1

Private udcResult As udcResult
Private wndResult As Window

Private Ret As String

Public Property Get Var_QuickButtonMode() As Long
    Var_QuickButtonMode = m_Var_QuickButtonMode
End Property

Public Property Let Var_QuickButtonMode(ByVal value As Long)
    iQuickButton.Visible = (value <> 0)
    m_Var_QuickButtonMode = value
    Select Case value
    Case 1
        iQuickButton.ToolTipText = "缩进当前过程代码"
    Case 2
        iQuickButton.ToolTipText = "缩进当前模块代码"
    Case 3
        iQuickButton.ToolTipText = "缩进当前工程代码"
    End Select
End Property

Private Function GetDescription(ByVal StructMark As CI_Structure) As String
    Select Case StructMark
    Case IF_Start
        GetDescription = "End If"
    Case IF_End
        GetDescription = "If"
    Case Select_Start
        GetDescription = "End Select"
    Case Select_End
        GetDescription = "Select Case"
    Case With_Start
        GetDescription = "End With"
    Case With_End
        GetDescription = "With"
    Case For_Start
        GetDescription = "Next"
    Case For_End
        GetDescription = "For"
    Case Proc_Start
        GetDescription = "End Sub/Function/Property"
    Case Proc_End
        GetDescription = "Sub/Function/Property"
    Case Type_Start
        GetDescription = "End Type"
    Case Type_End
        GetDescription = "Type"
    Case Enum_Start
        GetDescription = "End Enum"
    Case Enum_End
        GetDescription = "Enum"
    Case Do_Start
        GetDescription = "Loop"
    Case Do_End
        GetDescription = "Do"
    Case While_Start
        GetDescription = "Wend"
    Case Wend_End
        GetDescription = "While"
    End Select
    GetDescription = "缺少 " & GetDescription
End Function


Private Function IndentProc(ByVal Proc_Start As String, ByVal kProc As vbext_ProcKind) As Long
    With CodeOpe.TargetCodePane.CodeModule
        If Proc_Start <> "" Then
            IndentLines .ProcBodyLine(Proc_Start, kProc), .ProcStartLine(Proc_Start, kProc) + .ProcCountLines(Proc_Start, kProc) - 1
            IndentProc = .ProcStartLine(Proc_Start, kProc) + .ProcCountLines(Proc_Start, kProc) - 1
        Else
            IndentLines 1, .CountOfDeclarationLines
            IndentProc = .CountOfDeclarationLines
        End If
        
    End With
End Function

Private Sub PushError(ByVal Line As Long, ByVal StructMark As CI_Structure)
    udcResult.AddItem Line, GetDescription(StructMark), CodeOpe.TargetCodePane
End Sub

'未启用
Private Function GetBaseSpaceCount(ByVal s As String, ByVal ReferValue As Long) As Long
    Dim i As Long, c As String, l As String, r As String
    c = Mid(s, ReferValue, 1)
    If ReferValue - 1 > 0 Then l = Mid(s, ReferValue - 1, 1)
    If ReferValue + 1 <= Len(s) Then r = Mid(s, ReferValue + 1, 1)
End Function

'将一行代码加入Ret，并处理光标
Private Function PushToRet(ByVal i As Long, ByVal SpaceCount As Long) As Boolean
    
    Dim s As String, n As Long
    s = CodeOpe.Lines(i)
    n = GetLeftSpaceCount(s)
    
    If Ret = "" Then
        Ret = Space(SpaceCount) & LTrim(s)
    Else
        Ret = Ret & vbCrLf & Space(SpaceCount) & LTrim(s)
    End If
    
    
    
    If i = iSL Then
        iSC = iSC + (SpaceCount - n)
        If iSC < 0 Then iSC = 0
    End If
    
    If i = iEL Then
        iEC = iEC + (SpaceCount - n)
        If iEC < 0 Then iEC = 0
    End If
    
End Function


Private Function IndentLines(ByVal lStart As Long, ByVal lEnd As Long) As Boolean
    
    On Error GoTo iErr
    
    
    Dim i As Long, j As Long
    
    
    
    Dim BaseLine As Long
    'BaseLine为最终在代码窗格中处理的行序号
    Dim BaseSpaceCount As Long
    'BaseSpaceCount为“一分多”代码第二行起的空格数
    Dim ReferValue As Long
    
    
    '1.处理普通“一行一语句”：记为BaseLine①，处理字符串②，移动③
    
    '2.处理“一分多”（一条语句用“ _”写在多行）：
    '  将第一行记为BaseLine①，合并④，记第二行空格数为ReferValue（以实现保持相对位置不变）⑨，
    '  处理字符串②，按情况1移动首行③，得到第二行起缩进空格数BaseSpaceCount⑧，再移动其他行⑤
    
    '3.处理“多合一”（多条语句用“ :”写在一行）：分割⑥，以第一句决定本行缩进量②③，再处理其他语句（但不移动）⑦
    
    Ret = ""
    
    Dim s As String
    
    iSL = CodeOpe.SL
    iSC = CodeOpe.SC
    iEL = CodeOpe.EL
    iEC = CodeOpe.EC
    
    Stack.Clear
    
    For i = lStart To lEnd
        
        s = s & LTrim(CodeOpe.Lines(i))
        '通过CodeOpe.Lines(i)获取的代码不含换行符
        
        If s = "" Then
            Ret = Ret & vbCrLf
        Else
        
            If BaseLine = 0 Then
                BaseLine = i '①
            ElseIf ReferValue = 0 Then
                ReferValue = GetLeftSpaceCount(CodeOpe.Lines(i)) '⑨
            End If
    
            If Right(s, 2) <> " _" Then '单行代码/“一分多”代码末行
                
                
                Dim p() As String
                p = Split(s, ":") '⑥
                
                
                
                j = ProcessString(p(0), BaseLine) * Var_SpacePerLevel '②
                PushToRet BaseLine, j '③
                
                
                '⑧
                If Var_MulitLines_IndentSpaceCount > 0 Then
                    BaseSpaceCount = j + Var_MulitLines_IndentSpaceCount
                ElseIf Var_MulitLines_IndentSpaceCount = 0 Then
                    BaseSpaceCount = j + (ReferValue - GetLeftSpaceCount(CodeOpe.Lines(BaseLine))) '保持相对位置不变
                End If
                
                
                '⑦
                For j = 1 To UBound(p)
                    ProcessString p(j), BaseLine
                Next
                
                
                
                '⑤
                For j = BaseLine + 1 To i
                    'DBPrint "MTLines: " & BaseSpaceCount
                    PushToRet j, BaseSpaceCount
                Next
                
                
                
                s = ""
                BaseLine = 0
                BaseSpaceCount = 0
                ReferValue = 0
                
            Else
                
                s = Left(s, Len(s) - 2) '④
                
            End If
        
        End If
        
    Next
    
    CodeOpe.DeleteLines lStart, lEnd - lStart + 1
    CodeOpe.InsertLines lStart, Ret
    
    CodeOpe.SetCaret iSL, iSC, iEL, iEC
    
    IndentLines = True
    
    Exit Function
iErr:
    IndentLines = False
    'DBErr "iCode - Function IndentLines", "Code :" & vbCrLf & VBIns.ActiveCodePane.CodeModule.Lines(1, VBIns.ActiveCodePane.CodeModule.CountOfLines), "lStart = " & lStart, "lEnd = " & lEnd
End Function

'处理字符串（包括栈处理），返回缩进量
Private Function ProcessString(ByVal s As String, ByVal LineMark As Long) As Long
    
    Dim r As CI_Structure, v As CI_Structure
        
    r = CheckCode(s)
    
    If r = 0 Then
        ProcessString = Stack.count
    ElseIf (r > 10 And r < 20) Then
        ProcessString = Stack.count
        Stack.Push r
    ElseIf r > 20 Then
        
        Do Until Stack.count = 0

            v = Stack.Top
            If r < 30 Then Stack.Pop
            
            If (v Mod 10) = (r Mod 10) Then
                If r < 30 Then
                    ProcessString = Stack.count
                Else
                    
                    ProcessString = Stack.count - 1
                End If
                Exit Do
            Else
                PushError LineMark, v
            End If
            
        Loop
        
    End If
    
End Function

'获取左空格数（非缩进量）
Private Function GetLeftSpaceCount(ByVal s As String) As Long
    Dim n As Long
    For n = 1 To Len(s)
        If Mid(s, n + 1, 1) <> " " Then Exit For
    Next
    GetLeftSpaceCount = n
End Function

'匹配代码类型
Private Function CheckCode(ByVal s As String) As CI_Structure
    s = LTrim(s)
    If LeftIs(s, "If ") Or LeftIs(s, "#If ") Then
        Dim ts As String
        ts = LTrim(Right(s, Len(s) - (InStr(1, LCase(s), LCase("Then")) + 4) + 1))
        If ts = "" Or LeftIs(ts, "'") Then
            CheckCode = IF_Start
        Else
            CheckCode = 0
        End If
    ElseIf LCase(s) = LCase("Else") Or LCase(s) = LCase("#Else") Or LeftIs(s, "elseif ") Then
        CheckCode = IF_Middle
    ElseIf LeftIs(s, "End If") Or LeftIs(s, "#End If") Or LeftIs(s, "EndIf ") Or LeftIs(s, "EndIf'") Then
        CheckCode = IF_End
    ElseIf LeftIs(s, "Select Case ") Then
        CheckCode = Select_Start
    ElseIf LeftIs(s, "Case ") Then
        CheckCode = Select_Middle
    ElseIf LeftIs(s, "End Select") Then
        CheckCode = Select_End
    ElseIf LeftIs(s, "With ") Then
        CheckCode = With_Start
    ElseIf LeftIs(s, "End With") Then
        CheckCode = With_End
    ElseIf LeftIs(s, "For ") Then
        CheckCode = For_Start
    ElseIf LCase(s) = LCase("Next") Or LeftIs(s, "Next ") Or LeftIs(s, "Next'") Then
        CheckCode = For_End
    ElseIf LeftIs(s, "Private Sub ") Or LeftIs(s, "Private Function ") Or LeftIs(s, "Private Property ") _
        Or LeftIs(s, "Public Sub ") Or LeftIs(s, "Public Function ") Or LeftIs(s, "Public Property ") _
        Or LeftIs(s, "Friend Sub ") Or LeftIs(s, "Friend Function ") Or LeftIs(s, "Friend Property ") _
        Or LeftIs(s, "Sub ") Or LeftIs(s, "Function ") Or LeftIs(s, "Property ") Then
        CheckCode = Proc_Start
    ElseIf LeftIs(s, "End Sub") Or LeftIs(s, "End Function") Or LeftIs(s, "End Property") Then
        CheckCode = Proc_End
    ElseIf LeftIs(s, "Private Type ") Or LeftIs(s, "Public Type ") Or LeftIs(s, "Firend Type ") Or LeftIs(s, "Type ") Then
        CheckCode = Type_Start
    ElseIf LeftIs(s, "End Type") Then
        CheckCode = Type_End
    ElseIf LeftIs(s, "Private Enum ") Or LeftIs(s, "Public Enum ") Or LeftIs(s, "Firend Enum ") Or LeftIs(s, "Enum ") Then
        CheckCode = Enum_Start
    ElseIf LeftIs(s, "End Enum") Then
        CheckCode = Enum_End
    ElseIf LeftIs(s, "Do ") Or LCase(s) = LCase("Do") Then
        CheckCode = Do_Start
    ElseIf LeftIs(s, "While ") Then
        CheckCode = While_Start
    ElseIf LeftIs(s, "Loop ") Or LCase(s) = LCase("Loop") Then
        CheckCode = Do_End
    ElseIf LeftIs(s, "Wend ") Or LCase(s) = LCase("Wend") Then
        CheckCode = Wend_End
    End If
End Function

Public Sub Button_Init()
    
    Set iQuickButton = iToolBar.Controls.Add(msoControlButton)
    iQuickButton.Caption = "代码缩进"
    iQuickButton.Visible = True
    iQuickButton.BeginGroup = True
    Set iQuickButtonEvent = VBIns.Events.CommandBarEvents(iQuickButton)
    Clipboard.SetData LoadResPicture(101, 0)
    iQuickButton.PasteFace
    
    Set popIndent = iCodeMenu.Controls.Add(msoControlPopup)
    popIndent.Caption = "代码缩进"
    popIndent.Visible = True
    popIndent.BeginGroup = True
    
    Set btnErrorWindow = iCodeMenu.Controls.Add(msoControlButton)
    btnErrorWindow.Visible = True
    Set btnErrorWindowEvent = VBIns.Events.CommandBarEvents(btnErrorWindow)
    
    Dim tmpButton As CommandBarButton
    
    Set tmpButton = popIndent.Controls.Add(msoControlButton)
    tmpButton.Caption = "缩进当前过程代码"
    Set btnIndentProcEvent = VBIns.Events.CommandBarEvents(tmpButton)
    
    Set tmpButton = popIndent.Controls.Add(msoControlButton)
    tmpButton.Caption = "缩进当前模块代码"
    Set btnIndentModuleEvent = VBIns.Events.CommandBarEvents(tmpButton)
    
    Set tmpButton = popIndent.Controls.Add(msoControlButton)
    tmpButton.Caption = "缩进当前工程代码"
    Set btnIndentProjectEvent = VBIns.Events.CommandBarEvents(tmpButton)
    
End Sub

Public Sub Initialize(ByVal IDEIns As VBE, ByVal DBForm As Object, ByVal AddIn As AddIn, _
                      ByVal iMenu As Object, ByVal iBar As Object, _
                      ByVal WndGUID As String)
    
    Set VBIns = IDEIns
    hVBIDE = VBIns.MainWindow.hWnd
    
    Set DebugForm = DBForm
    
    Set iCodeMenu = iMenu
    Set iCodeMenuEvent = VBIns.Events.CommandBarEvents(iMenu)
    Set iToolBar = iBar
    Set AddInInst = AddIn
    
    Set wndResult = VBIns.Windows.CreateToolWindow(AddInInst, "iCode_CodeIndent.udcResult", "代码缩进 - 发现错误", WndGUID, udcResult)
    wndResult.Visible = False
    Button_Init
    
    
End Sub

Private Sub IndentInit()
    udcResult.Clear
End Sub

Private Sub IndentComplete()
    If Var_ResultWindow_AutoHide Then wndResult.Visible = (udcResult.count > 0)
End Sub

Private Sub btnErrorWindowEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    wndResult.Visible = Not wndResult.Visible
End Sub

Private Sub btnIndentModuleEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If Not (VBIns.SelectedVBComponent Is Nothing) Then
        Set CodeOpe.TargetCodePane = VBIns.SelectedVBComponent.CodeModule.CodePane
        IndentInit
        IndentLines 1, CodeOpe.TargetCodePane.CodeModule.CountOfLines
        IndentComplete
        Set CodeOpe.TargetCodePane = Nothing
    End If
End Sub

Private Sub btnIndentProcEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim ProcName As String, ProcKind As vbext_ProcKind
    CodeOpe.GetProcByLine CodeOpe.SL, ProcName, ProcKind
    IndentInit
    IndentProc ProcName, ProcKind
    IndentComplete
End Sub

Private Sub btnIndentProjectEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If Not (VBIns.ActiveVBProject Is Nothing) Then
        Dim i As Long
        IndentInit
        For i = 1 To VBIns.ActiveVBProject.VBComponents.count
            Set CodeOpe.TargetCodePane = VBIns.ActiveVBProject.VBComponents.item(i).CodeModule.CodePane
            IndentLines 1, CodeOpe.TargetCodePane.CodeModule.CountOfLines
        Next
        Set CodeOpe.TargetCodePane = Nothing
        IndentComplete
    End If
End Sub

Private Sub iCodeMenuEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If wndResult.Visible = False Then
        btnErrorWindow.Caption = "显示错误窗口"
    Else
        btnErrorWindow.Caption = "隐藏错误窗口"
    End If
End Sub

Private Sub iQuickButtonEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Select Case Var_QuickButtonMode
    Case 1
        btnIndentProcEvent_Click Nothing, False, False
    Case 2
        btnIndentModuleEvent_Click Nothing, False, False
    Case 3
        btnIndentProjectEvent_Click Nothing, False, False
    End Select
End Sub

