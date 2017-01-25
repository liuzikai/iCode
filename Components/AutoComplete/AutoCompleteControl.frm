VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAutoComplete 
   Appearance      =   0  'Flat
   BackColor       =   &H80000011&
   BorderStyle     =   0  'None
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   614
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chkInsLock 
      BackColor       =   &H00FFFFFF&
      Caption         =   "锁定插入位置"
      Height          =   255
      Left            =   3375
      TabIndex        =   4
      Top             =   3855
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.PictureBox Container 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   0
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   3
      Top             =   0
      Width           =   3255
   End
   Begin VB.PictureBox Container_D 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   1
      Top             =   1980
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label lblIns 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "插入位置 →"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   30
         Width           =   1695
      End
   End
   Begin VB.Timer Timer1 
      Left            =   1320
      Top             =   2880
   End
   Begin MSComctlLib.ImageList iImageList 
      Left            =   270
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":02F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":05E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":08D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":0BC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":0EBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":11AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":149E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":1790
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":1A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":1D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":2066
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":2358
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvModule 
      Height          =   1875
      Left            =   3270
      TabIndex        =   0
      Top             =   1980
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3307
      _Version        =   393217
      Indentation     =   0
      LabelEdit       =   1
      Style           =   1
      FullRowSelect   =   -1  'True
      ImageList       =   "ilModule"
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ilModule 
      Left            =   5100
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65535
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":264A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":27A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":28FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":2A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":2BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":2D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":2E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":2FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":311A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":3274
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":36C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoCompleteControl.frx":3B18
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblBlock 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3270
      TabIndex        =   6
      Top             =   3855
      Width           =   105
   End
   Begin VB.Label lblFailToFindModule 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "找不到先前的插入模块"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2265
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "frmAutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


'明确各组件范围:
'UI:处理控件大小位置等参数
'TV:处理TreeView内部信息，适用于当前TreeView
'STV:固态TreeView接口
'Action:处理响应动作(分配至UI动作、事件触发等)
'F:处理匹配


Private STV_Width() As Long
Private STV_Count As Long
Private tvTarget As TreeView

Private Const UI_Form_Dis_X = 7
Private Const UI_Form_Dis_Y = 5

Private Const UI_Dis = 3

Private Const UI_TV_DefultTextLen = 17
Private UI_TV_MaxTextLen As Long
Private Const UI_TV_ExtraWidth = 40


Public hParent As Long

Event ItemChoose(ByVal itsName As String, ByVal itsType As AC_ItemKind, ByVal Code As String, ByVal Scope As String, ByVal Module As Long)
Event ToHide()
Event GoBack()

Private F_Key As String

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_CLIPCHILDREN = &H2000000

Private Const UI_WidthFromCarpet = 10
Private Const UI_HeightFromCarpet = 15

Private Const UI_RetainOuterWidth = 34 + UI_WidthFromCarpet
Private Const UI_RetainOuterHeight = 34 + UI_HeightFromCarpet


'设置透明
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const UI_AlphaColor = vbGreen

Private m_LockAsPrivate As Boolean

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Public Property Get LockAsPrivate() As Boolean
    LockAsPrivate = m_LockAsPrivate
End Property

Public Property Let LockAsPrivate(ByVal value As Boolean)
    m_LockAsPrivate = value
    If value Then cmdScope.Caption = "Private"
End Property

Private Function SetOnTop(ByVal hWnd As Long, ByVal IsOnTop As Boolean)
    Dim rtn As Long
    If IsOnTop = True Then
        '将窗口置于最上面
        rtn = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Else
        rtn = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If
End Function

'启动
Public Sub iShow()

    If tvModule.DropHighlight Is Nothing Then Set tvModule.DropHighlight = tvModule.Nodes("C0")
    
    UI_UpdataLocation
    UI_Dec_Hide
    Form_MouseMove 0, 0, 0, 0 '隐藏插入位置选择器
    
    SetOnTop Me.hWnd, True
    Me.Show
    iSetFocus hParent
    
End Sub

Public Sub iClearCur()
    TV_Clear
    UI_Dec_Hide
    F_Clear
End Sub

Public Sub iClear()
    TV_SetTarget 0
    iClearCur
    tvModule.Nodes.Clear
    tvModule.Nodes.Add , , "C" & 0, "(当前模块)", 12, 12
    Form_MouseMove 0, 0, 0, 0 '隐藏插入位置选择器
End Sub

'设置透明
Private Sub UI_SetLayered()
    
    Me.BackColor = UI_AlphaColor
    
    SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, UI_AlphaColor, 0, 1
    
End Sub

'模块选择
Public Sub Moudle_ReadyToAdd()
    Set tvTarget = tvModule
    TV_ReadyToAdd 0
    '借助TV_AddItem的大小调整功能
End Sub

Public Sub Moudle_AddItem(ByVal itsName As String, ByVal Index As String, ByVal itsType As AC_ItemKind)
    On Error Resume Next
    tvTarget.Nodes.Add , , "C" & Index, itsName, itsType, itsType
    If Len(itsName) > UI_TV_MaxTextLen Then UI_TV_MaxTextLen = Len(itsName)
End Sub

Public Sub Module_FinishAdding()
    TV_FinishAdding
    TV_SetTarget 0, True
    tvModule.Visible = True
End Sub

Public Sub Moudle_Select(ByVal Index As Long)
    DBPrint "tvModule Select " & Index
    Set tvModule.DropHighlight = tvModule.Nodes.Item("C" & Index)
End Sub

'查找
Public Property Get F_Len() As Long
    F_Len = Len(F_Key)
End Property

Private Sub F_Clear()
    F_Key = ""
End Sub

Private Sub F_Find()
    
    If Len(F_Key) = 0 Then Exit Sub
    
    DBPrint "F_Key = " & F_Key
    
    Dim i As Long, j As Long, m As Long, n As Long
    
    For i = 1 To tvTarget.Nodes.count
        If Len(tvTarget.Nodes.Item(i).Key) >= Len(F_Key) Then
            j = 1
            Do
                If j > Len(F_Key) Then
                    Action_iSelect tvTarget.Nodes.Item(i)
                    Exit Sub
                End If
                
                If j > Len(tvTarget.Nodes.Item(i).Key) Then GoTo GoNext

                If Mid(F_Key, j, 1) = LCase(Mid(tvTarget.Nodes.Item(i).Key, j, 1)) Then
                    j = j + 1
                Else
GoNext:
                    If j - 1 > m Then
                        n = i
                        m = j - 1
                    End If
                    Exit Do
                End If
            Loop
        End If
    Next
    
    If n <> 0 Then
        Set tvTarget.SelectedItem = tvTarget.Nodes.Item(n)
        Set tvTarget.SelectedItem = Nothing
    End If
    
End Sub

Public Property Get TV() As TreeView
    Set TV = tvTarget
End Property

'静态TV
Public Function STV_Create(ByVal Tag As String) As Long

    STV_Count = STV_Count + 1
    
    ReDim Preserve STV_Width(STV_Count)
    
    Load iTreeView(STV_Count)
    iTreeView(STV_Count).Top = iTreeView(0).Top
    iTreeView(STV_Count).Left = iTreeView(0).Left
    iTreeView(STV_Count).Height = iTreeView(0).Height
    iTreeView(STV_Count).Visible = False
    
    iTreeView(STV_Count).Tag = Tag
    
    STV_Create = STV_Count
    
End Function

Public Function STV_Find(ByVal Tag As String) As Long
    Dim i As Long
    For i = 1 To STV_Count
        If iTreeView(i).Tag = Tag Then Exit For
    Next
    If i > STV_Count Then
        STV_Find = -1
    Else
        STV_Find = i
    End If
End Function

Private Sub UI_TV_ChangeWidth(ByVal n As Long)
    If hParent = 0 Then
        tvTarget.Width = n
    Else
        Dim rc As RECT
        GetWindowRect hParent, rc
        
        tvTarget.Width = Min(n, (rc.Right - rc.Left) - Me.ScaleX(Me.Left, 1, 3) - UI_RetainOuterWidth - UI_Form_Dis_X * 2)
    End If
    
    Container.Width = tvTarget.Width + tvTarget.Left * 2
    
    Container_D.Width = Container.Width
    lblIns.Width = Container_D.Width - lblIns.Left - 3
    
    If lblFailToFindModule.Visible Then
        lblFailToFindModule.Top = tvTarget.Top + tvTarget.Height
        lblFailToFindModule.Width = tvTarget.Width
    End If
End Sub

Public Sub TV_SetTarget(ByVal Index As Long, Optional ByVal NoClear As Boolean = False)
    
    If Index < 0 Then Exit Sub
    
    
    If Not (tvTarget Is Nothing) Then
    
        tvTarget.Visible = False
        
        '扫尾工作，保证下次开启这一列表时选中状态为无
        If NoClear = False And tvTarget.Nodes.count > 0 Then
            Set tvTarget.SelectedItem = tvTarget.Nodes.Item(1)
            Set tvTarget.SelectedItem = Nothing
            Set tvTarget.DropHighlight = Nothing
        End If
    
    End If
    
    Set tvTarget = iTreeView(Index)
    
    tvTarget.Visible = True
    
    UI_TV_ChangeWidth tvTarget.Width
    
End Sub

Public Sub TV_Clear()
    tvTarget.Visible = False
    tvTarget.Nodes.Clear
    Set tvTarget.DropHighlight = Nothing
    tvTarget.Visible = True
End Sub

Public Sub TV_ReadyToAdd(ByVal PBarMaxium As Long)
    tvTarget.Visible = False
    UI_TV_MaxTextLen = UI_TV_DefultTextLen
    If PBarMaxium > 100 Then
        PBar.Max = PBarMaxium + 3
        PBar.value = 0
        PBar.Visible = True
        PBar.Top = Container.Top + (Container.Height - PBar.Height) / 2
        PBar.Left = Container.Left + UI_Dis
        PBar.Width = Container.Width - 2 * UI_Dis
        
        lblLoading.Visible = True
        lblLoading.Top = PBar.Top - 2 * UI_Dis - lblLoading.Height
        lblLoading.Left = PBar.Left + (PBar.Width - lblLoading.Width) / 2
    End If
    DoEvents
End Sub

Public Sub TV_AddItem(ByVal itsName As String, ByVal Code As String, ByVal itsType As AC_ItemKind)
    On Error Resume Next
    tvTarget.Nodes.Add(, , itsName, itsName, itsType, itsType).Tag = Code
    If Len(itsName) > UI_TV_MaxTextLen Then UI_TV_MaxTextLen = Len(itsName)
    If PBar.Visible Then PBar.value = PBar.value + 1
    'DoEvents
End Sub

Public Sub TV_FinishAdding()
    'iTreeView.Sort慎用,F_Find需要用到项目名称的递增性,因此加载项目时应保证顺序
    tvTarget.Visible = True
    UI_TV_ChangeWidth UI_TV_MaxTextLen * Me.TextWidth("A") + UI_TV_ExtraWidth
    PBar.Visible = False
    lblLoading.Visible = False
End Sub

Private Sub cmdScope_Click()
    If cmdScope.Caption = "Public" Or LockAsPrivate Then cmdScope.Caption = "Private" Else cmdScope.Caption = "Public"
    'tvTarget.SetFocus
End Sub

Private Sub Container_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove Button, Shift, x, y
End Sub

Private Sub Form_Load()
    iClear
    UI_SetLayered
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If tvModule.Visible = True Then
        tvModule.Visible = False
        lblBlock.Visible = False
        chkInsLock.Visible = False
    End If
End Sub



Private Sub Action_iSelect(ByVal Node As Node)
    
    If tvTarget.Visible Then
        
        Set tvTarget.DropHighlight = Node
        
        Set tvTarget.SelectedItem = tvTarget.DropHighlight
        Set tvTarget.SelectedItem = Nothing
        
        If (Node.Image > 10 Or _
            (CodeOpe.bInDeclaration = True And ( _
                Node.Image = kMethod Or _
                Node.Image = kType Or _
                Node.Image = kConst))) _
            And Container_D.Visible = False Then
            
            UI_Dec_Show
        End If
        
        If (Node.Image <= 10 And CodeOpe.bInDeclaration = False) And Container_D.Visible = True Then
            UI_Dec_Hide
        End If
    End If
    
End Sub

Private Sub Action_iChoose(ByVal Node As Node)
    If Container_D.Visible Then
        RaiseEvent ItemChoose(Node.Key, Node.Image, Node.Tag, cmdScope.Caption, Right(tvModule.DropHighlight.Key, Len(tvModule.DropHighlight.Key) - 1))
    Else
        RaiseEvent ItemChoose(Node.Key, Node.Image, Node.Tag, "", 0)
    End If
End Sub

Private Sub UI_Dec_Show()
    Container_D.Visible = True
    lblFailToFindModule.Visible = False
End Sub

Private Sub UI_Dec_Hide()
    Container_D.Visible = False
    tvModule.Visible = False
    lblBlock.Visible = False
    chkInsLock.Visible = False
End Sub



Private Sub lblIns_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If tvModule.Visible = False Then
        
        tvModule.Visible = True
        tvModule.Left = Container_D.Left + Container_D.Width
        tvModule.Top = lblIns.Top + Container_D.Top
        
        lblBlock.Visible = True
        lblBlock.Left = tvModule.Left
        lblBlock.Top = tvModule.Top + tvModule.Height
        
        chkInsLock.Visible = True
        chkInsLock.Top = tvModule.Top + tvModule.Height
        chkInsLock.Left = lblBlock.Left + lblBlock.Width
        chkInsLock.Width = tvModule.Left + tvModule.Width - chkInsLock.Left
        
        lblFailToFindModule.Visible = False
        
    End If
End Sub

Private Sub tvModule_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub UI_UpdataLocation()

    If Not IsIDEMode Then
        Dim pt As POINTAPI
        pt = CodeOpe.GetCaretPoint(hParent)
        
'        Dim rc As RECT
'        GetWindowRect hParent, rc
        
        '避免超出边界
        'ACControl的宽度变化幅度很大，为避免空余空间过多，以当前宽度为准
'        If pt.X + TtoPx(tvTarget.Width) + UI_RetainOuterWidth > rc.Right - rc.Left Then _
'            pt.X = (rc.Right - rc.Left) - (TtoPx(tvTarget.Width) + UI_RetainOuterWidth)
        
        Me.Left = Me.ScaleX(pt.x + UI_WidthFromCarpet, 3, 1)
        
        'ACControl高度变化有限，以最大高度为准
'        If pt.Y + lblScope.Top + lblScope.Height + UI_RetainOuterHeight > rc.Bottom - rc.Top Then _
'            pt.Y = (rc.Bottom - rc.Top) - (lblScope.Top + lblScope.Height + UI_RetainOuterHeight)
            
        Me.Top = Me.ScaleY(pt.y + UI_HeightFromCarpet, 3, 1)
        
        DBPrint "UI_UpdataLocation:"
        DBPrint "   X = " & pt.x
        DBPrint "   Y = " & pt.y
    Else
        Me.Left = 0
        Me.Top = 0
    End If
    
End Sub

Public Sub Msg_Char(ByVal KeyAscii As Long)
    If InStr(1, SignOfChoose, Chr(KeyAscii)) <> 0 Then
        If Not (tvTarget.DropHighlight Is Nothing) Then Action_iChoose tvTarget.DropHighlight
        RaiseEvent ToHide
    ElseIf InStr(1, SignOfEnd, Chr(KeyAscii) <> 0) Then
        RaiseEvent ToHide
    ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 56 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 95 Then
        F_Key = F_Key & LCase(Chr(KeyAscii))
        F_Find
        
        UI_UpdataLocation
    End If
End Sub

'返回值：0表示保留消息，1表示阻止消息传递
Public Function Msg_Key(ByVal KeyCode As Long) As Long
    
    If KeyCode = 0 Then Exit Function
    
    Select Case KeyCode
        
    Case vbKeyBack
        
        If Len(F_Key) <= 0 Then
            RaiseEvent ToHide
            Exit Function
        End If
        
        F_Key = Left(F_Key, Len(F_Key) - 1)
        F_Find
        
    Case vbKeyUp
        
        Msg_Key = 1
        
        If tvTarget.DropHighlight Is Nothing Then
            'Do Nothing
        ElseIf tvTarget.DropHighlight.Index <= 1 Then
            'Do Nothing
        Else
            Action_iSelect tvTarget.Nodes.Item(tvTarget.DropHighlight.Index - 1)
        End If
        
    Case vbKeyDown
        
        Msg_Key = 1
        
        If tvTarget.DropHighlight Is Nothing Then
            Action_iSelect tvTarget.HitTest(1, 1)
        ElseIf tvTarget.DropHighlight.Index >= tvTarget.Nodes.count Then
            'Do Nothing
        Else
            Action_iSelect tvTarget.Nodes.Item(tvTarget.DropHighlight.Index + 1)
        End If
        
    Case vbKeyLeft
        
        RaiseEvent GoBack
        
        Msg_Key = 1
        
    Case vbKeyRight
        
        If tvTarget.DropHighlight Is Nothing Then
            RaiseEvent ToHide
        Else
            Action_iChoose tvTarget.DropHighlight
            Msg_Key = 1
        End If
    
    Case vbKeyTab
    
        If lblScope.Visible Then
            cmdScope_Click
            Msg_Key = 1
        End If
    
    Case vbKeyEscape
        
        RaiseEvent ToHide
    
    Case vbKeySpace, vbKeyReturn
        
        If Not (tvTarget.DropHighlight Is Nothing) Then
            Msg_Key = 1
            Action_iChoose tvTarget.DropHighlight
        Else
            Msg_Key = 0
            RaiseEvent ToHide
        End If
        
    End Select
    
End Function

Private Sub tvModule_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set tvModule.DropHighlight = tvModule.HitTest(1, y)
End Sub

Public Sub iUnLoad()
    
    Container.Left = 0
    Container.Top = 0
    
    Me.Width = Me.ScaleX(Container.Width, 3, 1)
    Me.Height = Me.ScaleY(Container.Height, 3, 1)

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    Dim i As Long, j As Long, k As Long
    
    PBar.Max = 1
    For i = 0 To STV_Count
        PBar.Max = PBar.Max + iTreeView(i).Nodes.count
        iTreeView(i).Visible = False
    Next
    
    k = 200
    
    PBar.value = 0
    PBar.Visible = True
    PBar.Top = Container.Top + (Container.Height - PBar.Height) / 2
    PBar.Left = Container.Left + UI_Dis
    PBar.Width = Container.Width - 2 * UI_Dis
    
    lblLoading.Caption = "正在清理..."
    lblLoading.Visible = True
    lblLoading.Top = PBar.Top - 2 * UI_Dis - lblLoading.Height
    lblLoading.Left = PBar.Left + (PBar.Width - lblLoading.Width) / 2
    
    Me.Show
    
    DoEvents
    
    For i = 0 To STV_Count
        Do Until iTreeView(i).Nodes.count = 0
            iTreeView(i).Nodes.Remove 1
            PBar.value = PBar.value + 1
            If PBar.value = k Then
                DoEvents
                k = k + 200
            End If
        Loop
        If i <> 0 Then Unload iTreeView(i)
        DoEvents
    Next
    
    Unload Me
    
End Sub

