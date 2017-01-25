VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeStatistic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Button As CommandBarButton
Private WithEvents ButtonEvent As CommandBarEvents
Attribute ButtonEvent.VB_VarHelpID = -1

Private Name_MembersType(5) As String
Private Name_ComponentType(11) As String
Private Name_ProjectType(3) As String

Private Type Result
    Chars As Long
    NotesChar As Long
    Lines As Long
    EmptyLines As Long
    NoteLines As Long
    Members(5) As Long
    Components(11) As Long
End Type

Dim tvResult As TreeView

Private WithEvents Operator As frmOperator
Attribute Operator.VB_VarHelpID = -1
Public Canceled As Boolean

Public Property Let MenuButton_BeginGroup(ByVal Value As Boolean)
    Button.BeginGroup = Value
End Property

Private Function Stat_Group() As Boolean

    Dim Result As Result
    Dim i As Long
    
    Dim Total As Long
    
    For i = 1 To VBIns.VBProjects.Count
        Total = Total + VBIns.VBProjects.Item(i).VBComponents.Count
    Next
    
    Operator.pbTotal.Max = Total
    Operator.pbTotal.Value = 0
    
    Operator.Show
    
    tvResult.Nodes.Add , , "Group", "所有工程"
    
    tvResult.Nodes.Item("Group").Bold = True
    tvResult.Nodes.Item("Group").Expanded = True
    
    For i = 1 To VBIns.VBProjects.Count
    
        Result_Add_B2A Result, Stat_Project(VBIns.VBProjects.Item(i), "Group")
        
        DoEvents
        If Canceled Then
            Canceled = False
            Exit Function
        End If

        
    Next
    
    PrintResult "Group", Result
    
    Stat_Group = True
    
End Function

Private Function Stat_Project(ByVal VBProject As VBProject, ByVal ParentKey As String) As Result
    
    tvResult.Nodes.Add ParentKey, tvwChild, ParentKey & "." & VBProject.Name, VBProject.Name & " (" & Name_ProjectType(VBProject.Type) & ")"
    
    Dim i As Long
    
    For i = 1 To VBProject.VBComponents.Count
        
        Result_Add_B2A Stat_Project, Stat_Component(VBProject.VBComponents.Item(i), ParentKey & "." & VBProject.Name)
        
        Stat_Project.Components(VBProject.VBComponents.Item(i).Type) = Stat_Project.Components(VBProject.VBComponents.Item(i).Type) + 1
        
        DoEvents
        If Canceled Then
            Exit For
            '返回到Stat_Group恢复Canceled状态
        End If
    
    Next
    
    PrintResult ParentKey & "." & VBProject.Name, Stat_Project
    
End Function

Private Function Stat_Component(ByVal VBComponent As VBComponent, ByVal ParentKey As String) As Result
    
    If VBComponent Is Nothing Then Exit Function
    If VBComponent.CodeModule Is Nothing Then Exit Function
    '可能会出现VBComponent无效的情况，原因尚不明确
    
    Operator.pbPart.Max = 3 '预留空间避免取消延迟
    If Operator.chkChar.Value = 1 Then
        Operator.pbPart.Max = Operator.pbPart.Max + VBComponent.CodeModule.CountOfLines
    End If
    If Operator.chkProc.Value = 1 Then
        Operator.pbPart.Max = Operator.pbPart.Max + VBComponent.CodeModule.Members.Count
    End If
    
    If Operator.pbPart.Max = 3 Then
        Operator.pbPart.Value = 3
    Else
        Operator.pbPart.Value = 1
    End If
    
    tvResult.Nodes.Add ParentKey, tvwChild, ParentKey & "." & VBComponent.Name, VBComponent.Name & " (" & Name_ComponentType(VBComponent.Type) & ")"
    
    With Stat_Component
        
        Dim i As Long
        Dim s As String
        
        .Lines = VBComponent.CodeModule.CountOfLines
        
        If Operator.chkChar.Value = 1 Then
            For i = 1 To VBComponent.CodeModule.CountOfLines
                s = Trim(VBComponent.CodeModule.Lines(i, 1))
                If s = "" Then
                    .EmptyLines = .EmptyLines + 1
                Else
                    .Chars = .Chars + Len(s)
                    If InStr(1, s, "'") <> 0 Then
                        .NotesChar = .NotesChar + Len(s) - InStr(1, s, "'") + 1
                    End If
                    If Left(s, 1) = "'" Then .NoteLines = .NoteLines + 1
                End If
                Operator.pbPart.Value = i
                DoEvents
                If Canceled Then Exit Function
            Next
        End If
        
        If Operator.chkProc.Value = 1 Then
            For i = 1 To VBComponent.CodeModule.Members.Count
                .Members(VBComponent.CodeModule.Members.Item(i).Type) = .Members(VBComponent.CodeModule.Members.Item(i).Type) + 1
                Operator.pbPart.Value = Operator.pbPart.Value + 1
                DoEvents
                If Canceled Then Exit Function
            Next
        End If
    End With
    
    Operator.pbTotal.Value = Operator.pbTotal.Value + 1
    
    PrintResult ParentKey & "." & VBComponent.Name, Stat_Component
    
End Function

Private Sub Result_Add_B2A(a As Result, b As Result)
    
    a.Chars = a.Chars + b.Chars
    a.NotesChar = a.NotesChar + b.NotesChar
    
    a.Lines = a.Lines + b.Lines
    a.EmptyLines = a.EmptyLines + b.EmptyLines
    a.NoteLines = a.NoteLines + b.NoteLines
    
    Dim i As Long
    
    For i = 1 To 5
        a.Members(i) = a.Members(i) + b.Members(i)
    Next
    
    For i = 1 To 11
        a.Components(i) = a.Components(i) + b.Components(i)
    Next
    
End Sub

Private Sub PrintResult(ByVal Key As String, Result As Result)
    
    With Result
        
        If Operator.chkChar.Value = 1 Then
            tvResult.Nodes.Add(Key, tvwChild, Key & ".Chars", "字符总数：" & .Chars).Bold = True
            tvResult.Nodes.Add Key & ".Chars", tvwChild, Key & ".Chars" & ".CodeChars", "有效字符数：" & .Chars - .NotesChar & " (" & Round((.Chars - .NotesChar) / .Chars * 100, 2) & "%)"
            tvResult.Nodes.Add Key & ".Chars", tvwChild, Key & ".Chars" & ".NoteChars", "注释字符数：" & .NotesChar
        End If
        
        tvResult.Nodes.Add(Key, tvwChild, Key & ".Lines", "总行数：" & .Lines).Bold = True
        tvResult.Nodes.Add Key & ".Lines", tvwChild, Key & ".Lines" & ".CodeLines", "有效行数：" & .Lines - .EmptyLines - .NoteLines & " (" & Round((.Lines - .EmptyLines - .NoteLines) / .Lines * 100, 2) & "%)"
        tvResult.Nodes.Add Key & ".Lines", tvwChild, Key & ".Lines" & ".EmptyLines", "空行数：" & .EmptyLines
        tvResult.Nodes.Add Key & ".Lines", tvwChild, Key & ".Lines" & ".NoteLines", "纯注释行数：" & .NoteLines
        
        Dim i As Long
        Dim Total As Long
        
        If Operator.chkProc.Value = 1 Then
            tvResult.Nodes.Add(Key, tvwChild, Key & ".Members", "").Bold = True
            For i = 1 To 5
                If .Members(i) > 0 Then
                    Total = Total + .Members(i)
                    tvResult.Nodes.Add Key & ".Members", tvwChild, _
                                       Key & ".Members" & "." & Name_MembersType(i), _
                                       Name_MembersType(i) & "：" & .Members(i)
                End If
            Next
            tvResult.Nodes.Item(Key & ".Members").Text = "总成员数：" & Total
        End If
        
        Total = 0
        
        tvResult.Nodes.Add(Key, tvwChild, Key & ".Components", "").Bold = True
        For i = 1 To 11
            If .Components(i) > 0 Then
                Total = Total + .Components(i)
                tvResult.Nodes.Add Key & ".Components", tvwChild, _
                                   Key & ".Components" & "." & Name_ComponentType(i), _
                                   Name_ComponentType(i) & "：" & .Components(i)
            End If
        Next
        tvResult.Nodes.Item(Key & ".Components").Text = "总组件数：" & Total
        
    End With
    
End Sub


Private Sub LoadTranslations()

    Name_MembersType(vbext_mt_Method) = "方法"
    Name_MembersType(vbext_mt_Property) = "属性"
    Name_MembersType(vbext_mt_Variable) = "变量"
    Name_MembersType(vbext_mt_Event) = "事件"
    Name_MembersType(vbext_mt_Const) = "常数"
    
    Name_ComponentType(vbext_ct_StdModule) = "标准模块" '1
    Name_ComponentType(vbext_ct_ClassModule) = "类模块" '2
    Name_ComponentType(vbext_ct_MSForm) = "窗体" '3
    Name_ComponentType(vbext_ct_ResFile) = "资源文件" '4
    Name_ComponentType(vbext_ct_VBForm) = "标准窗体" '5
    Name_ComponentType(vbext_ct_VBMDIForm) = "MDI窗体" '6
    Name_ComponentType(vbext_ct_PropPage) = "属性页" '7
    Name_ComponentType(vbext_ct_UserControl) = "用户控件" '8
    Name_ComponentType(vbext_ct_DocObject) = "用户文档" '9
    Name_ComponentType(vbext_ct_RelatedDocument) = "相关文档" '10
    Name_ComponentType(vbext_ct_ActiveXDesigner) = "设计器" '11
    
    Name_ProjectType(vbext_pt_StandardExe) = "标准EXE"
    Name_ProjectType(vbext_pt_ActiveXDll) = "ActiveX DLL"
    Name_ProjectType(vbext_pt_ActiveXExe) = "ActiveX EXE"
    Name_ProjectType(vbext_pt_ActiveXControl) = "ActiveX 控件"
    
End Sub

Public Sub Initialize(ByVal aVBIns As VBE, ByVal aDebugForm As Object, ByVal aiCodeMenu As Object)
    
    Set VBIns = aVBIns
    Let hVBIDE = VBIns.MainWindow.hWnd
    Set DebugForm = aDebugForm
    Set iCodeMenu = aiCodeMenu
    
    Set Button = iCodeMenu.Controls.Add(MsoControlType.msoControlButton, , , , True)
    Button.Style = msoButtonIconAndCaption
    Button.Caption = "代码统计"
    Clipboard.SetData LoadResPicture(101, 0)
    Button.PasteFace
    
    Set ButtonEvent = VBIns.Events.CommandBarEvents(Button)
    
    LoadTranslations
    
    Set tvResult = frmResult.tvResult
    
    Set Operator = New frmOperator
    
End Sub

Private Sub ButtonEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Operator.iShow
End Sub

Private Sub Class_Terminate()
    Unload Operator
End Sub

Private Sub Operator_Cancel()
    Canceled = True
End Sub

Private Sub Operator_Start()
    tvResult.Nodes.Clear
    If Stat_Group = True Then
        Operator.Hide
        frmResult.Show
    End If
End Sub
