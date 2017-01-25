VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodeOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const DBFlag = "CodeOpe::"

Public Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'TargetCodePane不受工程限制，可跨工程使用

Private m_SL As Long
Private m_SC As Long
Private m_EL As Long
Private m_EC As Long
Private m_hCodeWnd As Long

Private m_TargetCodePane As CodePane

Public Property Get hCodeWnd() As Long
    If m_hCodeWnd <> 0 Then
        hCodeWnd = m_hCodeWnd
    Else
        hCodeWnd = TargetCodePane.Window.hWnd
    End If
End Property

Public Property Get SL() As Long
    UpdataSelectionInfo
    SL = m_SL
End Property

'※ CodeOperator封装的SC属性：代表光标在第SC个字符后
Public Property Get SC() As Long
    UpdataSelectionInfo
    SC = m_SC - 1
End Property

'※ CodeOperator封装的EL属性：在选择大于等于两行时无视末行（原数据是倒数第二行）的换行符
Public Property Get EL() As Long
    UpdataSelectionInfo
    If (m_EC = 1 And m_SC > 1 And m_EL - m_SL > 1) Or (m_EC = 1 And m_SC = 1 And m_EL - m_SL >= 1) Then '判断语句使用原数据(m_)
        '选中多于2行，此时选中了第EL行末的换行符，可视为无意之举
        EL = m_EL - 1
    Else
        '选中单行或刻意选中了第SL行末的换行符（EL-SL=1）（选择整行除外）
        EL = m_EL
    End If
End Property

'※ CodeOperator封装的EC属性：代表光标在第EC个字符后，在选择大于等于两行时无视末行（原数据是倒数第二行）的换行符
Public Property Get EC() As Long
    UpdataSelectionInfo
    If (m_EC = 1 And m_SC > 1 And m_EL - m_SL > 1) Or (m_EC = 1 And m_SC = 1 And m_EL - m_SL >= 1) Then '判断语句使用原数据(m_)
        '选中多于2行，此时选中了第EL行末的换行符，可视为无意之举
        EC = Len(Lines(m_EL - 1))
    Else
        '选中单行或刻意选中了第SL行末的换行符（EL-SL=1）
        EC = m_EC - 1
    End If
End Property

Private Sub UpdataSelectionInfo()
    On Error Resume Next
    VBIns.ActiveCodePane.GetSelection m_SL, m_SC, m_EL, m_EC
End Sub

Public Property Get Lines(ByVal StartLine As Long, Optional ByVal count As Long = 1) As String
    Dim s As String
    s = TargetCodePane.CodeModule.Lines(StartLine, count)
    If Right(s, 1) = Chr(13) Or Right(s, 1) = vbCrLf Then
        s = Left(s, Len(s) - 1)
    End If
    Lines = s
End Property

'此过程未经测试
Public Property Let Lines(ByVal StartLine As Long, Optional ByVal count As Long = 1, ByVal value As String)
    If count = 1 Then
        ReplaceLine StartLine, value
    Else
        DeleteLines StartLine, count
        InsertLines StartLine, value
    End If
End Property


'向声明部分插入代码
Public Sub AddCodeToDeclaration(ByVal Code As String)
    '使用CodeModule.AddFromString应该也可以实现
    InsertLines LinesCountBeforeProcs, Code
    '自带换行
End Sub

Public Property Get TargetCodePane() As CodePane
    On Error Resume Next
    If Not HaveStaticTarget Then
        Set TargetCodePane = VBIns.ActiveCodePane
    Else
        Set TargetCodePane = m_TargetCodePane
    End If
End Property

Public Property Get HaveStaticTarget() As Boolean
    HaveStaticTarget = Not (m_TargetCodePane Is Nothing)
End Property

Public Property Set TargetCodePane(ByVal value As CodePane)
    Set m_TargetCodePane = value
End Property

'若同时省略EL、EC，则视为设置光标而无选中文本
'否则，省略项将以当前坐标替代
Public Sub SetCaret(Optional ByVal SL, Optional ByVal SC, Optional ByVal EL, Optional ByVal EC)
    If IsMissing(SL) Then SL = Me.SL
    If IsMissing(SC) Then SC = Me.SC
    
    If IsMissing(EL) And IsMissing(EC) Then
        TargetCodePane.SetSelection SL, SC + 1, SL, SC + 1
    Else
        If IsMissing(EL) Then EL = Me.EL
        If IsMissing(EC) Then EC = Me.EC
        TargetCodePane.SetSelection SL, SC + 1, EL, EC + 1
    End If
End Sub

Public Function GetCaretPoint(ByVal hWnd As Long) As POINTAPI
    Dim tPoint As POINTAPI
    GetCaretPos tPoint
    If tPoint.x < 0 Then '已选中文字，会导致横坐标不能正常获取，则坐标以当前行首字符为准
        tPoint.x = 34
    End If
    ClientToScreen hWnd, tPoint
    GetCaretPoint = tPoint
End Function

'注意：LinesCountBeforeProcs除了声明部分，还包括到第一个过程前的空行
'（尽管他们可能不属于声明部分（声明部分不为空时））
Public Property Get LinesCountBeforeProcs() As Long
    
    If IsIDEMode Then If TargetCodePane Is Nothing Then LinesCountBeforeProcs = -1: Exit Property
    'IDE模式下可能不能正确获取TargetCodePane
    
    Dim k As Long
    
    k = TargetCodePane.CodeModule.CountOfDeclarationLines
    
    Do While Trim(TargetCodePane.CodeModule.Lines(k + 1, 1)) = "" And k <= TargetCodePane.CodeModule.CountOfLines
        k = k + 1
    Loop
    
    LinesCountBeforeProcs = k
End Property

'注意：bInDeclaration返回光标是否处于“声明部分以及到第一个过程前的空行”中
'（尽管他们可能不属于声明部分，但执行插入代码时，置光标于此可视为“意图”插入到声明部分）
Public Property Get bInDeclaration() As Boolean
    bInDeclaration = (LinesCountBeforeProcs = -1) Or (Me.SL <= LinesCountBeforeProcs)
End Property


Public Property Get Selection() As String
    If SL = EL Then
        Selection = Mid(Lines(SL), SC + 1, EC - SC)
    Else
        Selection = Right(Lines(SL), Len(Lines(SL)) - SC)
        If EL - SL > 1 Then Selection = Selection & vbCrLf & Lines(SL + 1, EL - SL - 1) & vbCrLf
        If EC > 0 Then Selection = Selection & vbCrLf & Left(Lines(EL), EC)
    End If
End Property

'直接替换选择的文本，并配置对应光标
Public Property Let Selection(ByVal value As String)
    
    Dim s As String
    
    s = Left(Lines(SL), SC)
    s = s & value
    s = s & Right(Lines(EL), Len(Lines(EL)) - EC)
    
    Dim oldSL As Long: oldSL = SL
    Dim oldSC As Long: oldSC = SC
    Dim ToNextLine As Boolean: ToNextLine = (m_EL <> EL)
    
    If EL > SL Then DeleteLines SL + 1, EL - SL
    ReplaceLine SL, s
    
    Dim l() As String: l = Split(value, vbCrLf)
    
    If Not ToNextLine Then
        If UBound(l) = 0 Then
            SetCaret oldSL, oldSC, oldSL, oldSC + Len(l(UBound(l)))
        Else
            SetCaret oldSL, oldSC, oldSL + UBound(l), Len(l(UBound(l)))
        End If
    Else
        SetCaret oldSL, oldSC, oldSL + UBound(l) + 1, 0
    End If
    
End Property


Public Sub InsertToCaret(ByVal sCode As String)
    
    Dim s As String, tSL As Long, tEL As Long, tSC As Long, tEC As Long
    
    tSL = SL: tSC = SC
    tEL = EL: tEC = EC
    
    If tSL = tEL Then
        
        If tSL > TargetCodePane.CodeModule.CountOfLines Then
            TargetCodePane.CodeModule.InsertLines tSL, sCode
        Else
            s = Me.Lines(tSL)
            ReplaceLine tSL, Left(s, tSC) & sCode & Right(s, Len(s) - tEC)
        End If
        SetCaret tSL, tSC + Len(sCode), tEL, tSC + Len(sCode)
        '使用ReplaceLine或InsertLine焦点会改变
    Else
        
        Me.Selection = sCode
        
    End If
End Sub

Public Sub GetProcByLine(ByVal n As Long, ByRef sProc As String, ByRef kProc As vbext_ProcKind)
    On Error Resume Next
    Dim k As vbext_ProcKind: k = kProc
    sProc = TargetCodePane.CodeModule.ProcOfLine(n, k)
End Sub

Public Sub ReplaceAll_InProject(ByVal Project As VBProject, ByVal Find As String, ByVal Replace As String)
    Dim i As Long
    For i = 1 To Project.VBComponents.count
        Dim TragetCodeModule As CodeModule
        Set TragetCodeModule = Project.VBComponents.Item(i).CodeModule
        Dim s As String
        s = TragetCodeModule.Lines(1, TragetCodeModule.CountOfLines)
        If InStr(1, s, Find) <> 0 Then
            s = iReplaceAll(s, Find, Replace)
            TragetCodeModule.DeleteLines 1, TragetCodeModule.CountOfLines
            TragetCodeModule.InsertLines 1, s
        End If
    Next
End Sub


'简单继承于CodePane和CodeModule的方法和属性

Public Sub ReplaceLine(ByVal Line As Long, ByVal Code As String)
    TargetCodePane.CodeModule.ReplaceLine Line, Code
End Sub

Public Sub DeleteLines(ByVal StartLine As Long, Optional ByVal count As Long = 1)
    TargetCodePane.CodeModule.DeleteLines StartLine, count
End Sub

Public Sub InsertLines(ByVal Line As Long, ByVal Code As String)
    TargetCodePane.CodeModule.InsertLines Line, Code
End Sub

