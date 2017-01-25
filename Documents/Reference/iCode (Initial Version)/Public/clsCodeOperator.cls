VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long


Private m_SL As Long
Private m_SC As Long
Private m_EL As Long
Private m_EC As Long

Private m_TargetCodePane As CodePane

Public m As Variant

Public Property Get SL() As Long
    SL = m_SL
End Property

Public Property Let SL(ByVal Value As Long)
    m_SL = Value
End Property

Public Property Get SC() As Long
    SC = m_SC
End Property

Public Property Let SC(ByVal Value As Long)
    m_SC = Value
End Property

Public Property Get EL() As Long
    EL = m_EL
End Property

Public Property Let EL(ByVal Value As Long)
    m_EL = Value
End Property

Public Property Get EC() As Long
    EC = m_EC
End Property

Public Property Let EC(ByVal Value As Long)
    m_EC = Value
End Property

Public Sub UpdataSelectionInfo()
    VBIns.ActiveCodePane.GetSelection m_SL, m_SC, m_EL, m_EC
End Sub

Public Property Get Lines(ByVal StartLine As Long, Optional ByVal Count As Long = 1) As String
    Lines = TargetCodePane.CodeModule.Lines(StartLine, Count)
End Property

'此过程未经测试
Public Property Let Lines(ByVal StartLine As Long, Optional ByVal Count As Long = 1, ByVal Value As String)
    If Count = 1 Then
        TargetCodePane.CodeModule.ReplaceLine StartLine, Value
    Else
        TargetCodePane.CodeModule.DeleteLines StartLine, Count
        TargetCodePane.CodeModule.InsertLines StartLine, Value
    End If
End Property

Public Sub ReplaceLine(ByVal Line As Long, ByVal Code As String)
    TargetCodePane.CodeModule.ReplaceLine Line, Code
End Sub


'向声明部分插入代码
'返回值：声明部分行数
Public Function AddCodeToDeclaration(ByVal Code As String) As Long
    TargetCodePane.CodeModule.InsertLines TargetCodePane.CodeModule.CountOfDeclarationLines + 1, Code
    AddCodeToDeclaration = TargetCodePane.CodeModule.CountOfDeclarationLines
End Function

Public Property Get TargetCodePane() As CodePane
    If m_TargetCodePane Is Nothing Then
        Set TargetCodePane = VBIns.ActiveCodePane
    Else
        Set TargetCodePane = m_TargetCodePane
    End If
End Property

Public Property Set TargetCodePane(ByVal Value As CodePane)
    If Value Is Nothing Then
        Set m_TargetCodePane = Nothing
    Else
        Set m_TargetCodePane = Value
    End If
End Property

Public Sub SetCaret(Optional ByVal SL, Optional ByVal SC, Optional ByVal EL, Optional ByVal EC)
    If IsMissing(SL) Then SL = Me.SL
    If IsMissing(SC) Then SC = Me.SC
    If IsMissing(EL) Then EL = Me.EL
    If IsMissing(EC) Then EC = Me.EC
    TargetCodePane.SetSelection SL, SC, EL, EC
    Me.UpdataSelectionInfo
End Sub

Friend Function GetCaretPoint() As POINTAPI
    Dim tPoint As POINTAPI
    GetCaretPos tPoint
    ClientToScreen CodeOpe.TargetCodePane.Window.hWnd, tPoint
    GetCaretPoint = tPoint
End Function

Public Property Get bInDeclaration() As Boolean
    Me.UpdataSelectionInfo
    
    Dim k As Long
    
    k = VBIns.ActiveCodePane.CodeModule.CountOfDeclarationLines
    
    Do While Trim(VBIns.ActiveCodePane.CodeModule.Lines(k + 1, 1)) = "" And k <= VBIns.ActiveCodePane.CodeModule.CountOfLines
        k = k + 1
    Loop
    
    bInDeclaration = (Me.SL <= k)
End Property

Public Property Get Selection() As String
    Dim i As Long
    
    Me.UpdataSelectionInfo
    
    If SL = EL Then
        Selection = Mid(Me.Lines(SL), SC, EC - SC)
    Else
        Selection = Right(Me.Lines(SL), Len(Me.Lines(SL)) - SC + 1) & vbCrLf
        If EL - SL >= 2 Then Selection = Selection & Me.Lines(SL + 1, EL - SL - 1)
        '注意EC是指光标在第EC个字符前
        If EC = 1 Then
            If EL - SL >= 2 Then Selection = Selection & vbCrLf
        Else
            Selection = Selection & Left(Me.Lines(EL), EC - 1)
        End If
    End If
    
End Property

Public Property Let Selection(ByVal Value As String)
    Me.UpdataSelectionInfo
    
    If Me.SL = Me.EL Then
        Me.InsertCode Value
    Else
        
        Dim nS As Long, nE As Long, sF As String
        
        If SC = 1 Then
            nS = Me.SL
        Else
            '#对此处SC是否需要减一保留怀疑#
            sF = Left(Me.Lines(Me.SL), Me.SC - 1)
            nS = Me.SL + 1
        End If
        
        If Me.EC = Len(Me.Lines(Me.EL)) - 1 Then
            nE = Me.EL
            '如果选择末端包括了EL整行，则整行删去
        ElseIf Me.EC = 1 Then
            nE = Me.EL - 1
            '如果选择末端包含EL-1行的行末换行，则删除至EL-1行
        Else
            
            If SC <> 1 Then
                sF = sF & Right(Me.Lines(Me.EL), Len(Me.Lines(Me.EL)) - (Me.EC - 1))
                nE = Me.EL
                '如果选择起始前与选择末端后都各有一段，则将选择末行整行删去，将选择末行中选择末端右侧，与选择起始行中选择起始左端合并，再往其中间插入代码
            Else
                Me.Lines(Me.EL) = Right(Me.Lines(Me.EL), Len(Me.Lines(Me.EL)) - (Me.EC - 1))
                nE = Me.EL - 1
                '如果选择起始行需整行删去，则只需直接处理末行部分
                '这里删除至EL-1行，之前行会为空而不会保留空行
            End If
        End If
        
        If sF <> "" Then
            TargetCodePane.CodeModule.ReplaceLine Me.SL, sF
            '这里应用Me.SL而不是nS
        End If
        
        If nS <= nE Then
            TargetCodePane.CodeModule.DeleteLines nS, nE - nS + 1
            'nS已经考虑过sF是否存在的情况了
        End If
        
        'SL、SC、EL、EC都还保持删改前的状态
        
        Dim OldEC As Long: OldEC = Me.EC
        
        TargetCodePane.SetSelection Me.SL, Me.SC, Me.SL, Me.SC
        '#此语句仍存在Bug#
        Me.UpdataSelectionInfo
        
        'SL、SC、EL、EC为插入代码的状态
        
        Me.InsertCode Value
        
        If InStr(1, Value, vbCrLf) = 0 Then
            TargetCodePane.SetSelection Me.SL, Me.SC, Me.SL, Me.SC + Len(Value)
        Else
            Dim a() As String, cL As Long
            a = Split(Value, vbCrLf)
            cL = UBound(a)
            '这里的Ubound(a)即是代码中换行符的个数，即是除了插入行外的增加的行数
            'If OldEC = 1 Then cL = cL - 1
            TargetCodePane.SetSelection Me.SL, Me.SC, Me.SL + cL, Len(a(UBound(a))) + 1
        End If
        
    End If
End Property



Public Sub InsertCode(ByVal sCode As String)
    UpdataSelectionInfo
    
    Dim s As String
    
    If SL = EL Then
        s = Me.Lines(SL)
        TargetCodePane.CodeModule.ReplaceLine SL, Left(s, SC - 1) & sCode & Right(s, Len(s) - EC + 1)
        TargetCodePane.SetSelection SL, SC, EL, SC + Len(sCode)
    Else
        
        Me.Selection = sCode
        
    End If
End Sub

Public Function GetProcByLine(ByVal n As Long, ByRef sProc As String, ByRef kProc As vbext_ProcKind) As Boolean
    Dim k As vbext_ProcKind: k = kProc
    sProc = TargetCodePane.CodeModule.ProcOfLine(n, k)
    GetProcByLine = True
End Function

Public Function GetActiveLine() As Long
    Me.UpdataSelectionInfo
    GetActiveLine = Me.SL
End Function

Public Sub ReplaceAll(ByVal Find As String, ByVal Replace As String)
    Dim i As Long
    For i = 1 To VBIns.CodePanes.Count
        Set TargetCodePane = VBIns.CodePanes(i)
        Dim j As Long
        For j = 1 To TargetCodePane.CodeModule.CountOfLines
            Dim s As String
            s = iReplaceAll(Me.Lines(j), Find, Replace)
            If s <> Me.Lines(j) Then TargetCodePane.CodeModule.ReplaceLine j, s
        Next
    Next
    Set TargetCodePane = Nothing '将TargetCodePane设回默认，即激活的CodePane
End Sub






