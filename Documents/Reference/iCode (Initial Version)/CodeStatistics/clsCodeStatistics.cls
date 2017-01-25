VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents BH As ButtonsCollection
Attribute BH.VB_VarHelpID = -1

Private Sub ProjectStatistics()
    Dim Head As String
    Dim cLines As Long
    Dim cChars As Long
    Dim cCom(1 To 11) As Long
    Dim cMembers(1 To 5) As Long
    
    Head = "工程" & VBIns.ActiveVBProject.Name
    
    Dim i As Long
    
    For i = 1 To VBIns.ActiveVBProject.VBComponents.Count
        cCom(VBIns.ActiveVBProject.VBComponents.item(i).type) = cCom(VBIns.ActiveVBProject.VBComponents.item(i).type) + 1
    Next
    
    Dim k As Long
    
    For k = 1 To VBIns.CodePanes.Count
        
        Set CodeOpe.TargetCodePane = VBIns.CodePanes.item(k)
        
        cLines = cLines + CodeOpe.TargetCodePane.CodeModule.CountOfLines
        
        For i = 1 To CodeOpe.TargetCodePane.CodeModule.CountOfLines
            cChars = cChars + Len(CodeOpe.Lines(i))
        Next
        
        For i = 1 To CodeOpe.TargetCodePane.CodeModule.Members.Count
            cMembers(CodeOpe.TargetCodePane.CodeModule.Members.item(i).type) = cMembers(CodeOpe.TargetCodePane.CodeModule.Members.item(i).type) + 1
        Next
        
    Next
    
    Set CodeOpe.TargetCodePane = Nothing
    
    PrintOutResult Head, cLines, cChars, VBIns.ActiveVBProject.References.Count, cCom, cMembers
    
End Sub

Private Sub ComponentStatistics()
    
    Dim Head As String
    Dim cLines As Long
    Dim cChars As Long
    Dim cCom(0) As Long
    Dim cMembers(1 To 5) As Long
    
    Head = Translate_ComponentType(VBIns.SelectedVBComponent.type) & VBIns.SelectedVBComponent.Name
    cLines = CodeOpe.TargetCodePane.CodeModule.CountOfLines
    
    Dim i As Long
    For i = 1 To cLines
        cChars = cChars + Len(CodeOpe.Lines(i))
    Next
    
    For i = 1 To CodeOpe.TargetCodePane.CodeModule.Members.Count
        cMembers(CodeOpe.TargetCodePane.CodeModule.Members.item(i).type) = cMembers(CodeOpe.TargetCodePane.CodeModule.Members.item(i).type) + 1
    Next
    
    PrintOutResult Head, cLines, cChars, -1, cCom, cMembers
    
End Sub

Private Sub PrintOutResult(ByVal Head As String, _
    ByVal cLines As Long, _
    ByVal cChars As Long, _
    ByVal cReferences As Long, _
    ByRef cCom() As Long, _
    ByRef cMembers() As Long)
    
    Dim s As String
    s = "对" & Head & "的代码统计如下：" & vbCrLf & vbCrLf
    s = s & "代码行数：" & cLines & vbCrLf & vbCrLf
    s = s & "字符数：" & cChars & vbCrLf & vbCrLf
    If cReferences <> -1 Then s = s & "引用/部件数：" & cReferences & vbCrLf & vbCrLf
    
    Dim n As Long, i As Long, p As String
    
    If UBound(cCom) > 0 Then
        
        p = " "
        
        For i = 1 To UBound(cCom)
            n = n + cCom(i)
            If cCom(i) <> 0 Then p = p & Translate_ComponentType(i) & ":" & cCom(i) & " "
        Next
        
        s = s & "模块总数：" & n & "(" & p & ")" & vbCrLf & vbCrLf
        
    End If
    
    n = 0
    p = " "
    
    For i = 1 To UBound(cMembers)
        n = n + cMembers(i)
        If cMembers(i) <> 0 Then p = p & Translate_MemberType(i) & ":" & cMembers(i) & " "
    Next
    
    s = s & "成员总数：" & n & "(" & p & ")" & vbCrLf & vbCrLf
    
    MsgBox s, vbInformation Or vbYes, "iCode"
    
End Sub

Private Function Translate_MemberType(ByVal c As vbext_MemberType) As String
    If c = vbext_mt_Method Then
        Translate_MemberType = "方法"
    ElseIf c = vbext_mt_Property Then
        Translate_MemberType = "属性"
    ElseIf c = vbext_mt_Variable Then
        Translate_MemberType = "变量"
    ElseIf c = vbext_mt_Event Then
        Translate_MemberType = "事件"
    ElseIf c = vbext_mt_Const Then
        Translate_MemberType = "常数"
    End If
End Function

Private Function Translate_ComponentType(ByVal c As vbext_ComponentType) As String
    If c = vbext_ct_StdModule Then                                              'c=1
        Translate_ComponentType = "标准模块"
    ElseIf c = vbext_ct_ClassModule Then                                        'c=2
        Translate_ComponentType = "类模块"
    ElseIf c = vbext_ct_MSForm Then                                             'c=3
        Translate_ComponentType = "窗体"                                        '#此处存疑#
    ElseIf c = vbext_ct_ResFile Then                                            'c=4
        Translate_ComponentType = "资源文件"
    ElseIf c = vbext_ct_VBForm Then                                             'c=5
        Translate_ComponentType = "标准窗体"
    ElseIf c = vbext_ct_VBMDIForm Then                                          'c=6
        Translate_ComponentType = "MDI窗体"
    ElseIf c = vbext_ct_PropPage Then                                           'c=7
        Translate_ComponentType = "属性页"
    ElseIf c = vbext_ct_UserControl Then                                        'c=8
        Translate_ComponentType = "用户控件"
    ElseIf c = vbext_ct_DocObject Then                                          'c=9
        Translate_ComponentType = "用户文档"
    ElseIf c = vbext_ct_RelatedDocument Then                                    'c=10
        Translate_ComponentType = "RelatedDocument"
    ElseIf c = vbext_ct_ActiveXDesigner Then                                    'c=11
        Translate_ComponentType = "设计器"
    End If
End Function

Private Sub BH_ButtonClick(ByVal CommandBarControl As Office.CommandBarControl)
    Select Case CommandBarControl.Caption
    Case "代码统计(当前工程)"
        ProjectStatistics
    Case "代码统计(当前部件)"
        ComponentStatistics
    End Select
End Sub

Private Sub Class_Initialize()
    Set BH = New ButtonsCollection
    Set BH.Target = iCodeMenu
    BH.Add msoControlButton, "代码统计(当前工程)", msoButtonCaption
    BH.Add msoControlButton, "代码统计(当前部件)", msoButtonCaption
End Sub

Private Sub Class_Terminate()
    Set BH = Nothing
End Sub
