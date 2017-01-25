VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AC_AutoCompleteControl 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1845
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   166
   Begin iCode_Project.SelectBox SB_Scope 
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1530
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   450
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer mTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1140
      Top             =   990
   End
   Begin VB.PictureBox ColIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   60
      Picture         =   "AC_AutoCompleteControl.frx":0000
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList iImageList 
      Left            =   1830
      Top             =   870
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
            Picture         =   "AC_AutoCompleteControl.frx":02E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AC_AutoCompleteControl.frx":05D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AC_AutoCompleteControl.frx":08C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AC_AutoCompleteControl.frx":0BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AC_AutoCompleteControl.frx":0EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AC_AutoCompleteControl.frx":119E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AC_AutoCompleteControl.frx":1490
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AC_AutoCompleteControl.frx":1782
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AC_AutoCompleteControl.frx":1A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AC_AutoCompleteControl.frx":1D66
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AC_AutoCompleteControl.frx":205A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AC_AutoCompleteControl.frx":234C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AC_AutoCompleteControl.frx":263E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView iTreeView 
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   2566
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   0
      Style           =   1
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "iImageList"
      Appearance      =   0
   End
   Begin VB.Label SB_Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "声明："
      Height          =   180
      Left            =   60
      TabIndex        =   3
      Top             =   1530
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Line ColLine 
      BorderColor     =   &H80000000&
      X1              =   4
      X2              =   162
      Y1              =   17
      Y2              =   17
   End
   Begin VB.Label ColText 
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   330
      TabIndex        =   2
      Top             =   45
      Visible         =   0   'False
      Width           =   2115
   End
End
Attribute VB_Name = "AC_AutoCompleteControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum AC_ItemKind
    kMethod = 1
    kProperty = 2
    kVar = 3
    kEvent = 4
    kConst = 5
    kClass = 6
    kModule = 7
    kType = 8
    kEnum = 9
    kCollection = 10
    kUnUsedMethod = 11
    kUnUsedConst = 12
    kUnUsedType = 13
End Enum

Private sKey As String

Private mTreeViewDown() As Single, mTreeViewUp() As Single
Private mFormDown() As Single, mFormUp() As Single
Private mbColShow As Boolean, mbSBShow As Boolean
Private mMoveType As Long, t As Long
Private Const mTimes As Long = 10

Private m_Items() As New AC_Item

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_KILLFOCUS = &H8
Private Const WM_SETFOCUS = &H7

Private Const SignOfChoose As String = "!^*()-+=;:,<.>/"
Private Const SignOfEnd As String = "`~@#$%&[{]}\|'?"""

Private WithEvents ATimer_Width As AccelerationMotionTimer
Attribute ATimer_Width.VB_VarHelpID = -1
Private WithEvents ATimer_Height As AccelerationMotionTimer
Attribute ATimer_Height.VB_VarHelpID = -1
Private WithEvents ATimer_Top As AccelerationMotionTimer
Attribute ATimer_Top.VB_VarHelpID = -1

Event ItemChoose(item As AC_Item, ByVal Scope As String, bClose As Boolean)
Event CollectionEvent(Collection() As AC_Item, ByVal Count As Long)
Event GotFocus()

Private m_Collection() As New AC_Item
Private ItemsMaxLen As Long

Private nFormBorderWidth As Single, nFormBorderHeight As Single
Private Const nColTVDis As Single = 3

Private Const pDataWin32API As String = "\API Data\WIN32API.MDB"
Private Const pDataGDIP As String = "\API Data\GDIP.MDB"

Private DataCnn As New ADODB.Connection
Private DataRS As New ADODB.Recordset

Private AC_StartL As Long, AC_StartC As Long, AC_CharCount As Long

Dim bReadyToShowACBox As Long

Public Function DealMessage(ByVal Msg As Long, ByVal wParam As Long, ByVal Caption As String, ByVal ClassName As String) As Boolean
    
    DealMessage = True
    
    Select Case Msg
        
    Case WM_KEYUP
        
        Select Case hWnd
            
        Case vbKeyControl
            
            If bReadyToShowACBox = True Then Me.ACShow
            
            bReadyToShowACBox = False
            
        End Select
        
    Case WM_CHAR
        
        If Me.Visible = True Then Me.CharAction wParam
        
    Case WM_KEYDOWN
        
        Select Case wParam
            
        Case vbKeyBack, vbKeyReturn, vbKeySpace, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
            
            If iState = isCode And Me.Visible = True Then
                DealMessage = Me.KeyAction(wParam)
            End If
            
        Case vbKeyControl
            
            bReadyToShowACBox = True
            
        Case Else
            
            bReadyToShowACBox = False
            
        End Select
        
    Case WM_LBUTTONUP
        
        If Me.Visible = True Then iCode.AC.HideBox
        
    End Select
    
End Function

Private Sub DataConnectWin32Api()
    DataCnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & pDataWin32API
    DataRS.CursorLocation = adUseClient
End Sub

Private Sub DataConnectGDIP()
    DataCnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & pDataGDIP
    DataRS.CursorLocation = adUseClient
End Sub

Private Sub DataClose()
    DataRS.Close
    DataCnn.Close
End Sub

Public Sub ACShow()
    
    Dim tPoint As POINTAPI
    
    tPoint = CodeOpe.GetCaretPoint
    
    Me.Left = Me.ScaleX(tPoint.X + 3, 3, 1)
    Me.Top = Me.ScaleY(tPoint.Y + 20, 3, 1)
    
    ACSetFirstLevel
    
    Me.AddItemsToTreeview
    Me.iTreeView.Sorted = True
    
    Me.Show
    
    modPublic.SetFocus VBIns.ActiveCodePane.Window.hWnd
    
    CodeOpe.UpdataSelectionInfo
    
    AC_StartL = CodeOpe.SL
    AC_StartC = CodeOpe.SC
    AC_CharCount = 0
    
    'Me.SB_Scope.SelectItem 1
    'Me.SB_Scope.Draw
End Sub

Public Sub ACSetFirstLevel()
    Me.AddItem "# Win32Api", kCollection
    Me.AddItem "# Gdip", kCollection
    
    ACGetUsedMembers
End Sub


'这个过程是依据TreeView来设置其他控件
Private Sub SetControlsArea(Optional ByVal OnlyHeight As Boolean = False)
    SB_Scope.Top = iTreeView.Top + iTreeView.Height
    SB_Label.Top = SB_Scope.Top + (SB_Scope.Height - SB_Label.Height) / 2
    
    Me.Height = Me.ScaleY(iTreeView.Top + iTreeView.Height + 18, 3, 1)
    
    If Not OnlyHeight Then
        ColText.Width = iTreeView.Left + iTreeView.Width - ColText.Left
        ColLine.X2 = ColText.Left + ColText.Width
        
        SB_Scope.Width = iTreeView.Left + iTreeView.Width - SB_Scope.Left
        
        Me.Width = Me.ScaleX(iTreeView.Left * 2 + iTreeView.Width + nFormBorderWidth, 3, 1)
        
    End If
    
    SB_Scope.Cls
    SB_Scope.Draw
End Sub

Private Sub ATimer_Width_Timer(ByVal Value As Double)
    iTreeView.Width = Value
    SetControlsArea
End Sub

Public Property Get ItemsTotal() As Long
    ItemsTotal = UBound(m_Items)
End Property

Public Function FindItemByStr(ByVal Str As String) As Long
    Dim i As Long
    
    For i = 1 To Me.ItemsTotal
        If m_Items(i).Str = Str Then Exit For
    Next
    
    If Not IsNumeric(i) Or Val(i) > Me.ItemsTotal Then
        FindItemByStr = 0
    Else
        FindItemByStr = i
    End If
End Function

Friend Property Get Items(ByVal n) As AC_Item
    If Not (IsNumeric(n) And n <= Me.ItemsTotal) Then
        Dim i As Long
        i = FindItemByStr(n)
        If i = 0 Then Exit Property
    End If
    
    Set Items = m_Items(n)
End Property

Friend Property Set Items(ByVal n, ByVal NewValue As AC_Item)
    If Not (IsNumeric(n) And n <= Me.ItemsTotal) Then
        Dim i As Long
        i = FindItemByStr(n)
        If i = 0 Then Exit Property
    End If
    
    Set m_Items(n) = NewValue
End Property

Public Property Get Collection(ByVal Index As Long) As AC_Item
    If Index = 0 Or Index > CollectionCount Then Exit Property
    
    Set Collection = m_Collection(Index)
    
End Property

Private Sub SetTimerWidth(Optional ByVal nItemsMaxLen)
    If Not IsMissing(nItemsMaxLen) Then ItemsMaxLen = nItemsMaxLen
    If ItemsMaxLen < 19 Then ItemsMaxLen = 19
    ATimer_Width.StartValue = iTreeView.Width
    ATimer_Width.EndValue = Me.TextWidth(Space(ItemsMaxLen)) + 43
    ATimer_Width.Strat
End Sub

Public Sub Clear()
    ReDim m_Items(0)
    
    iTreeView.Visible = False
    
    iTreeView.Nodes.Clear
    
    iTreeView.Visible = True
    
    sKey = ""
    
    SetTimerWidth 19
End Sub

'Function CharAction
'输入：ByVal KeyAscii As Long —— KeyAscii值
'注意：此过程只接受 WM_CHAR 事件产生的键码
Public Sub CharAction(ByVal KeyAscii As Long)
    
    If InStr(1, SignOfChoose, Chr(KeyAscii)) <> 0 Then
        
        If Not (iTreeView.DropHighlight Is Nothing) Then CallEventItemChoose iTreeView.DropHighlight
        
        HideBox
        
        AC_CharCount = 0
        
    ElseIf InStr(1, SignOfEnd, Chr(KeyAscii) <> 0) Then
        
        HideBox
        
        AC_CharCount = 0
        
    Else
        
        sKey = sKey & Chr(KeyAscii)
        
        FindItem
        
        AC_CharCount = AC_CharCount + 1
        
    End If
End Sub

'Function KeyAction
'输入：ByVal KeyCode As Long —— KeyCode值
'返回值： Bool —— 是否保留按键信息
'注意：此过程只接受 WM_KEY 事件产生的键码
Public Function KeyAction(ByVal KeyCode As Long) As Boolean
    
    If KeyCode = 0 Then Exit Function
    
    KeyAction = True
    
    Dim bNoFind As Boolean
    
    Select Case KeyCode
        
    Case vbKeyBack
        
        DBPrint "vbKeyBack"
        
        If Len(sKey) <= 0 Then
            HideBox
            AC_CharCount = 0
            'KeyAction = -36
            Exit Function
        End If
        
        sKey = Left(sKey, Len(sKey) - 1)
        
        AC_CharCount = AC_CharCount - 1
        
        FindItem
        
        
    Case vbKeyUp
        
        If iTreeView.SelectedItem.Index = 1 Then Exit Function
        
        If Not (iTreeView.DropHighlight Is Nothing) Then Set iTreeView.SelectedItem = iTreeView.Nodes.item(iTreeView.SelectedItem.Index - 1)
        
        TreeView_SetFocus
        
    Case vbKeyDown
        
        If iTreeView.SelectedItem.Index = iTreeView.Nodes.Count Then Exit Function
        
        
        If Not (iTreeView.DropHighlight Is Nothing) Then
            Set iTreeView.SelectedItem = iTreeView.SelectedItem.Next
        End If
        
        TreeView_SetFocus
        
        bNoFind = True
        
        KeyAction = False
        'KeyAction = 0
        
    Case vbKeyLeft
        
        If CollectionCount >= 1 Then
            CollectionBack
            If mbSBShow = True Then
                mMoveType = 4
                mTimer.Enabled = True
                SB_Label.Visible = False
                SB_Scope.Visible = False
            End If
        End If
        
        KeyAction = False
        'KeyAction = 0
        
    Case vbKeyRight
        
        If iTreeView.DropHighlight Is Nothing Then
            
            HideBox
            
            AC_CharCount = 0
            'KeyAction = -3607
            
        Else
            
            If Items(iTreeView.DropHighlight.Index).Kind = kCollection Then CollectionAdd Items(iTreeView.DropHighlight.Index)
            
            KeyAction = False
            'KeyAction = 0
            
        End If
        
    Case vbKeySpace, vbKeyReturn
        
        If Not (iTreeView.DropHighlight Is Nothing) Then CallEventItemChoose iTreeView.DropHighlight
        
        HideBox
        
        AC_CharCount = 0
        'KeyAction = -3607
        
    End Select
    
End Function

Private Sub TreeView_KillFocus()
    Set iTreeView.DropHighlight = Nothing
End Sub

Private Sub TreeView_SetFocus()
    Set iTreeView.DropHighlight = iTreeView.SelectedItem
    Call Event_GotFocus
    
    If iTreeView.SelectedItem Is Nothing Then Exit Sub
    If (iTreeView.SelectedItem.Tag > 10 Or CodeOpe.bInDeclaration = True) And mbSBShow = False Then
        'mMoveType = 3
        'mTimer.Enabled = True
        
        ATimer_Height.Strat
        
        SB_Label.Visible = True
        SB_Scope.Visible = True
        'SB_Scope.Draw
        
        SetControlsArea
        
        mbSBShow = True
    ElseIf iTreeView.SelectedItem.Tag <= 10 And mbSBShow = True Then
        'mMoveType = 4
        'mTimer.Enabled = True
        ATimer_Height.RollBack
        
        SB_Label.Visible = False
        SB_Scope.Visible = False
        
        mbSBShow = False
    End If
    
    If mbSBShow = True Then SB_Scope.Draw
End Sub

Private Sub ATimer_Height_Timer(ByVal Value As Double)
    Me.Height = Value
End Sub

Private Sub ATimer_Top_Timer(ByVal Value As Double)
    iTreeView.Top = Value
    SetControlsArea True
End Sub

Private Sub FindItem()
    
    If sKey = "" Then
        TreeView_SelectItem 1
        TreeView_KillFocus
        Exit Sub
    End If
    
    Dim ResItem As Long, MaxChar As Long, i As Long, j As Long
    
    j = 1
    
    For i = 1 To iTreeView.Nodes.Count
        With iTreeView.Nodes.item(i)
            Dim c1 As Long, c2 As Long
            
            c1 = Asc(LCase(Mid(Items(i).Str, j, 1)))
            c2 = Asc(LCase(Mid(sKey, j, 1)))
            
            If c1 < c2 Then
                
            ElseIf c1 = c2 Then
                ResItem = i
                If j >= Len(sKey) Then
                    TreeView_SelectItem ResItem
                    TreeView_SetFocus
                    Exit For
                ElseIf j >= Len(Items(i).Str) Then
                    TreeView_SelectItem ResItem
                    TreeView_KillFocus
                    Exit For
                Else
                    j = j + 1
                    i = i - 1
                End If
            Else
                If ResItem = 0 Then ResItem = 1
                TreeView_SelectItem ResItem
                TreeView_KillFocus
                Exit For
            End If
        End With
    Next
    
End Sub

Private Sub TreeView_SelectItem(ByVal nItem As Long)
    
    If nItem > Me.ItemsTotal Then Exit Sub
    
    If iTreeView.SelectedItem Is iTreeView.Nodes.item(nItem) Then Exit Sub
    
    If nItem <> 1 Then Set iTreeView.SelectedItem = iTreeView.Nodes.item(iTreeView.Nodes.Count)
    
    Set iTreeView.SelectedItem = iTreeView.Nodes.item(nItem)
End Sub

Private Sub Form_GotFocus()
    Call Event_GotFocus
End Sub

Private Sub CallEventItemChoose(n As Node)
    If Items(n.Index).Kind = kCollection Then
        
        CollectionAdd Items(n.Index)
        
    Else
        
        Dim bRes As Boolean: bRes = True
        
        If mbSBShow = True Then
            Call Event_ItemChoose(Items(n.Index), SB_Scope.Selection, bRes)
        Else
            Call Event_ItemChoose(Items(n.Index), "", bRes)
        End If
        
        If bRes = True Then HideBox
        
    End If
End Sub

Private Sub Event_ItemChoose(item As AC_Item, ByVal Scope As String, bClose As Boolean)
    On Error GoTo iErr
    
    CodeOpe.UpdataSelectionInfo
    
    Dim s As String, r As String
    
    s = CodeOpe.Lines(AC_StartL)
    
    r = Left(s, (AC_StartC - 1))
    
    If CodeOpe.bInDeclaration = True Then
        r = r & Scope & " " & item.Code
    Else
        r = r & item.Str
    End If
    
    r = r & Right(s, Len(s) - AC_CharCount - (AC_StartC - 1))
    
    
    CodeOpe.ReplaceLine AC_StartL, r
    
    VBIns.ActiveCodePane.SetSelection CodeOpe.SL, AC_StartC + Len(item.Code), CodeOpe.EL, AC_StartC + Len(item.Code)
    
    If CodeOpe.bInDeclaration = False And Scope <> "" Then
        CodeOpe.AddCodeToDeclaration Scope & " " & item.Code
    End If
    
    Exit Sub
    
iErr:
    DBErr "AutoComplect_ItemChoose", "item.Str = " & item.Str, "item.Kind = " & item.Kind, "item.Kind = " & item.Code
    
End Sub

Private Sub Event_GotFocus()
    modPublic.SetFocus VBIns.ActiveCodePane.Window.hWnd
End Sub

Public Sub HideBox()
    Me.Hide
    CollectionClear
    Call Event_GotFocus
    Me.Clear
End Sub

Private Sub Form_Load()
    nFormBorderWidth = Me.ScaleX(Me.Width, 1, 3) - Me.ScaleWidth
    nFormBorderHeight = Me.ScaleY(Me.Height, 1, 3) - Me.ScaleHeight
    
    ReDim m_Items(0)
    ReDim m_Collection(0)
    
    SetAlwaysOnTop Me.hWnd, True
    
    SB_Scope.Init
    '这里的手动初始化并不是很恰当
    SB_Scope.AddItem "Private"
    SB_Scope.AddItem "Public"
    SB_Scope.SelectItem 1
    'SB_Scope.Draw
    
    Set ATimer_Width = New AccelerationMotionTimer
    
    Set ATimer_Height = New AccelerationMotionTimer
    ATimer_Height.StartValue = Me.ScaleY(iTreeView.Top * 2 + ColText.Height + nColTVDis + iTreeView.Height + nFormBorderHeight, 3, 1)
    ATimer_Height.EndValue = Me.ScaleY(iTreeView.Top * 2 + ColText.Height + nColTVDis + iTreeView.Height + SB_Scope.Height + nFormBorderHeight, 3, 1)
    
    Set ATimer_Top = New AccelerationMotionTimer
    ATimer_Top.StartValue = iTreeView.Top
    ATimer_Top.EndValue = iTreeView.Top + ColText.Height + nColTVDis
    
    
    'iTreeView.Width = Me.TextWidth(Space(ItemsMaxLen)) + 43
    SetControlsArea
    
    Me.Height = Me.ScaleY(iTreeView.Top * 2 + iTreeView.Height + nFormBorderHeight, 3, 1)
End Sub

Private Sub iTreeView_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub iTreeView_DblClick()
    If Not (iTreeView.DropHighlight Is Nothing) Then
        CallEventItemChoose iTreeView.DropHighlight
    End If
    
End Sub

Private Sub iTreeView_GotFocus()
    Call Event_GotFocus
End Sub

Private Sub iTreeView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TreeView_KillFocus
End Sub

Public Sub ReDimItems(ByVal nCount As Long)
    ReDim Preserve m_Items(nCount)
End Sub

Public Sub CollectionClear()
    
    If ColText.Caption = "" Then Exit Sub
    
    ColText.Caption = ""
    
    ReDim m_Collection(0)
    
    RaiseEvent CollectionEvent(m_Collection, CollectionCount)
    
    ColIcon.Visible = False
    ColText.Visible = False
    mbColShow = False
    
    ATimer_Top.RollBack
    
End Sub

Public Sub CollectionBack()
    
    If CollectionCount = 0 Then Exit Sub
    
    
    If CollectionCount = 1 Then
        CollectionClear
    Else
        ColText.Caption = Left(ColText.Caption, InStrRev(ColText.Caption, " > ") - 1)
        
        ReDim Preserve m_Collection(CollectionCount - 1)
        
        RaiseEvent CollectionEvent(m_Collection, CollectionCount)
    End If
    
    
End Sub

Public Property Get CollectionCount() As Long
    
    CollectionCount = UBound(m_Collection)
    
End Property

Public Sub CollectionAdd(ByVal item As AC_Item)
    
    If ColText.Caption = "" Then
        
        ColText.Caption = RTrim(item.Code)
        
        ATimer_Top.Strat
        
    Else
        ColText.Caption = ColText.Caption & " > " & RTrim(Right(item.Code, Len(item.Code) - 2))
    End If
    
    ReDim Preserve m_Collection(CollectionCount + 1)
    
    Set m_Collection(CollectionCount) = item
    
    Call Event_CollectionEvent(m_Collection, CollectionCount)
    
End Sub

Private Sub iTreeView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TreeView_SetFocus
    Call Event_GotFocus
End Sub

Public Sub AddItem(ByVal sItem As String, ByVal kItem As AC_ItemKind, Optional ByVal sCode As String = "", Optional ByVal Index As Long = 0)
    
    If Index = 0 Then
        ReDim Preserve m_Items(Me.ItemsTotal + 1)
        Index = Me.ItemsTotal
    End If
    
    m_Items(Index).Str = sItem
    m_Items(Index).Kind = kItem
    
    If Len(sItem) > ItemsMaxLen Then
        
        ItemsMaxLen = Len(sItem)
        
    End If
    
    If sCode = "" Then
        m_Items(Index).Code = sItem
    Else
        m_Items(Index).Code = sCode
    End If
    
End Sub

Private Sub CalcAndSetWidth()
    
End Sub

Public Sub AddItemsToTreeview(Optional ByVal bNoSort As Boolean = False, Optional bHaveScrollBar As Boolean = False)
    On Error GoTo iErr
    
    Dim j As Long, m As Long, p As Long, temp As AC_Item
    
    If bNoSort = False Then
        
        For j = 1 To Me.ItemsTotal - 1
            
            m = j
            
            For p = j + 1 To Me.ItemsTotal
                If Items(m).Str > Items(p).Str Or (Items(m).Str = Items(p).Str And Items(m).Kind > Items(p).Kind) Then m = p
            Next p
            
            If m <> j Then
                Set temp = Items(j)
                Set Items(j) = Items(m)
                Set Items(m) = temp
            End If
            
            
        Next
        
    End If
    
    iTreeView.Visible = False
    
    SetControlsArea
    
    For j = 1 To Me.ItemsTotal
        
        Randomize
        
        Dim i As Long, s As String
        
        s = Items(j).Str
        
        If Items(j).Kind = kCollection Then
            i = ItemsMaxLen - Len(Items(j).Str) - 1
            If i < 0 Then i = 0
            s = s & Space(i) & ">"
        End If
        
        iTreeView.Nodes.Add(, tvwChild, "Item - " & Items(j).Str & " - " & Items(j).Kind & " - " & CStr(Rnd), s, Items(j).Kind).Tag = Items(j).Kind
        
    Next
    
    iTreeView.Visible = True
    
    TreeView_SelectItem 1
    TreeView_KillFocus
    
    SetTimerWidth
    
    Exit Sub
    
iErr:
    DBErr "AC_AutoCompleteControl - AddItemsToTreeview"
End Sub

Private Sub Event_CollectionEvent(Collection() As AC_Item, ByVal Count As Long)
    
    On Error GoTo iErr
    
    Me.Clear
    
    If Count = 0 Then
        ACSetFirstLevel
        Me.AddItemsToTreeview
    Else
        
        Select Case Collection(Count).Str
            
        Case "# Win32Api", "# Gdip"
            
            Me.AddItem "# Constants 常数", kCollection, "# Constants"
            Me.AddItem "# Declares 方法", kCollection, "# Declares"
            Me.AddItem "# Types 类型", kCollection, "# Types"
            
            Me.AddItemsToTreeview
            
        Case Else
            
            Select Case Collection(Count - 1).Str
                
            Case "# Win32Api"
                DataConnectWin32Api
            Case "# Gdip"
                DataConnectGDIP
            End Select
            
            Dim DataRecordSet As String, DataFieldName As String, DataFieldCode As String, DataItemsType As AC_ItemKind
            
            Select Case Collection(Count).Str
                
            Case "# Constants 常数"
                DataRecordSet = "Constants"
                DataFieldName = "Name"
                
                DataItemsType = kUnUsedConst
                
                'If bInDeclaration = True Then
                DataFieldCode = "FullName"
                'Else
                '    DataFieldCode = ""
                'End If
                
            Case "# Declares 方法"
                DataRecordSet = "Declares"
                DataFieldName = "Name"
                
                DataItemsType = kUnUsedMethod
                
                'If bInDeclaration = True Then
                DataFieldCode = "FullName"
                'Else
                '    DataFieldCode = ""
                'End If
                
            Case "# Types 类型"
                DataRecordSet = "Types"
                DataFieldName = "Name"
                
                DataItemsType = kUnUsedType
                
                'If bInDeclaration = True Then
                'DataFieldCode = "#"
                'Else
                '    DataFieldCode = ""
                'End If
            End Select
            
            DataRS.Open "select * from " & DataRecordSet & " order by " & DataFieldName & " ASC", DataCnn, 3, 3
            
            Me.ReDimItems DataRS.RecordCount
            
            '            If DataFieldCode = "" Then
            '
            '                Do While Not DataRS.EOF
            '                    me.AddItem DataRS.Fields(DataFieldName), DataItemsType, , DataRS.AbsolutePosition
            '                    DataRS.MoveNext
            '                Loop
            '
            '            Else
            If Collection(Count).Str = "# Types 类型" Then
                
                Dim DataRsTypeItems As New ADODB.Recordset
                
                Dim SC As String
                
                Do While Not DataRS.EOF
                    
                    DataRsTypeItems.Open "select * from TypeItems where TypeID = " & DataRS.Fields("ID"), DataCnn, 3, 3
                    
                    SC = "Type " & DataRS.Fields(DataFieldName) & vbCrLf
                    
                    Do While Not DataRsTypeItems.EOF
                        SC = SC & DataRsTypeItems.Fields("TypeItem") & vbCrLf
                        DataRsTypeItems.MoveNext
                    Loop
                    
                    SC = SC & "End Type"
                    
                    DataRsTypeItems.Close
                    
                    Me.AddItem DataRS.Fields(DataFieldName), DataItemsType, SC, DataRS.AbsolutePosition
                    
                    DataRS.MoveNext
                Loop
            Else
                Do While Not DataRS.EOF
                    Me.AddItem DataRS.Fields(DataFieldName), DataItemsType, DataRS.Fields(DataFieldCode), DataRS.AbsolutePosition
                    DataRS.MoveNext
                Loop
            End If
            '        End If
            
            ReplaceUsedMembers DataItemsType
            
            Me.AddItemsToTreeview True, True
            
            DataClose
            
            
            
        End Select
        
        modPublic.SetFocus VBIns.ActiveCodePane.Window.hWnd
        
    End If
    
    Exit Sub
    
iErr:
    DBErr "clsCodeControler - me_CollectionEvent", "Count = " & Count, "Collection(Count).Str = " & Collection(Count).Str, "Collection(Count-1).Str = " & Collection(Count - 1).Str
End Sub

Public Sub ACGetUsedMembers()
    
    Dim i As Long, j As Long, v As CodePane
    
    For i = 1 To VBIns.CodePanes.Count
        
        Set v = VBIns.CodePanes.item(i)
        
        For j = 1 To v.CodeModule.Members.Count
            With v.CodeModule.Members.item(j)
                If v Is VBIns.ActiveCodePane Or .Scope = vbext_Public Or .Scope = vbext_Friend Then Me.AddItem .Name, .type
            End With
        Next
        
    Next
    
End Sub


Private Sub ReplaceUsedMembers(ByVal ItemsType As AC_ItemKind)
    
    If ItemsType = kUnUsedMethod Then
        ItemsType = kMethod
    ElseIf ItemsType = kUnUsedConst Then
        ItemsType = kConst
    ElseIf ItemsType = kUnUsedType Then
        ItemsType = kType
    End If
    
    
    Dim i As Long, j As Long, k As Long, v As CodePane
    
    For i = 1 To VBIns.CodePanes.Count
        
        Set v = VBIns.CodePanes.item(i)
        
        For j = 1 To v.CodeModule.Members.Count
            With v.CodeModule.Members.item(j)
                If v Is VBIns.ActiveCodePane Or .Scope = vbext_Public Or .Scope = vbext_Friend Then
                    If .type = ItemsType Then
                        k = Me.FindItemByStr(.Name)
                        If k <> 0 Then
                            Me.Items(k).Kind = .type
                        End If
                    End If
                End If
            End With
        Next
    Next
    
End Sub

