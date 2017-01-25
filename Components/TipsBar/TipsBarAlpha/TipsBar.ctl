VERSION 5.00
Begin VB.UserControl TipsBar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F2F2F2&
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   ForeColor       =   &H00585858&
   ScaleHeight     =   139
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   Begin VB.PictureBox XImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2820
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   4
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox XImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2280
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   3
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox XButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1140
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   2
      Top             =   1260
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox Tip 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00CFCFCF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   1
      Top             =   1260
      Width           =   855
   End
   Begin VB.Timer Timer_Cursor 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3900
      Top             =   1260
   End
   Begin VB.PictureBox XImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1800
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      Top             =   1260
      Width           =   315
   End
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   1260
   End
   Begin VB.Menu mnuTipsBar 
      Caption         =   "TipsBar"
      Visible         =   0   'False
      Begin VB.Menu mnuCloseThis 
         Caption         =   "关闭此标签(&T)"
      End
      Begin VB.Menu mnuCloseOthers 
         Caption         =   "关闭其它标签(&E)"
      End
      Begin VB.Menu mnuCloseLeft 
         Caption         =   "关闭左侧标签(&L)"
      End
      Begin VB.Menu mnuCloseRight 
         Caption         =   "关闭右侧标签(&R)"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "关闭所有标签(&A)"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShuiPing 
         Caption         =   "水平平铺(&H)"
      End
      Begin VB.Menu mnuChuiZhi 
         Caption         =   "垂直平铺(&V)"
      End
      Begin VB.Menu mnuCengDie 
         Caption         =   "层叠(&C)"
      End
   End
End
Attribute VB_Name = "TipsBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'用户操作触发TipClick和TipClose事件，不会自动切换焦点/移除标签
Event TipClick(ByVal ID As Long)
Event TipClose(ByVal ID As Long)
Event mnuCengDie()
Event mnuShuiPing()
Event mnuChuiZhi()



Private Const Dis_DefultFirstTip = 5   '第一个标签与UserControl.Left的默认水平距离
Private Const Dis_TextBorder_X = 10    '文字与Tip.Left\Right水平距离
'Private Const Dis_TextBorder_Y = 4     '文字与Tip.Top\Bottom竖直距离

Private Const Dis_XButton_Cut = 8      'XButton与Tip重叠的水平长度
Private Const Dis_XButton_Middle_X = 8 '"X"字样以此常量为中心X坐标绘制,XButton右侧空白长度可通过控件调整
Private Const Dis_XButton_R = 7        '"X"字样底部圆形半径
Private Const Dis_XButton_L = 3        '"X"字样外接正方形的边长的一半

'Tip背景色
Private Const Color_Back_Normal = &HF2F2F2
Private Const Color_Back_Active = vbWhite
Private Const Color_Back_Point = &HF6F6F6
'Tip文字前景色
Private Const Color_Fonts_Active = &H333333
Private Const Color_Fonts_Normal = &H585858



'"X"字样颜色
Private Const Color_X = &HB6B6B6

'"X"字样底部圆形背景颜色
Private Const Color_XButton_BackColor = &HF2F2F2
Private Const Color_XButton_BackColor_Pressed = &HE0E0E0

Private Enum TB_X_Style
    Normal = 0   '一般状态
    Active = 1   '指针指向状态
    Pressed = 2  '鼠标按下但未弹起状态
End Enum
Private X_State As TB_X_Style




Private Const Timer_Interval = 30    'Timer参数,决定Tip移动效果是否流畅(适中为佳)
Private Const Timer_TotalTime = 150  'Tip移动效果总时间

Private Enum TB_Drag_States
    Drag_None = 0           '没有按下鼠标左键(开始拖动)
    Drag_Pressed = 1        '按下了鼠标左键,但是尚未触发拖动操作
    Drag_ing = 2            '拖动中,Tip随鼠标移动
End Enum
Private Drag_State As TB_Drag_States

Private Const Drag_StartValue = 25   '触发拖动操作所需的最小X坐标移动量(相对于进入Drag_Pressed状态时)

Private Drag_ID As Long          '正在被拖动的Tip的ID,未拖动时需置零
Private Drag_CursorLeft As Long  '进入Drag_Pressed状态时,光标X坐标,用于判断进入Drag_ing状态
Private Drag_ID_Left As Long     '进入Drag_Pressed状态时,光标与Drag_Tip.Left的距离,用于Drag_ing状态时Tip的移动


'★关于序号的约定:
'ID(UID\ActiveID\PointID\Drag_ID):均为"控件ID",实际操作控件所用编号,UID持续增加,Tip销毁后其ID亦不再使用,对应总数UID
'Queue(i):目前从左到右各个Tip对应的ID,对应总数Count

Private UID As Long
Private aCaption(3000) As String '标题
Private Queue(300) As Long, Count As Long

Public ActiveID As Long     '激活ID,字体加粗,背景为Color_Back_Active,带有XButton
Public PointID As Long      '指向ID,背景为Color_Back_Point

Private mnuClickID As Long  '弹出菜单时光标指向的TipID

Private FirstTipLeft As Long


'★关于动画效果的主体思路:

'为了保证流畅度,动画效果尽可能简单迅速

'Drag_Swap\Drag_Drop操作发生后,直接交换Queue项,需要展示动画效果的ID赋值给Timer_ID,独立出来运行
'如果中途另一端动画要开始播放,直接执行Adjust操作,则正在运行的Tip立即归位

'即是说,Tip序号操作即刻完成,动画效果单独运行,即使被打断,不会影响正常Tip排布

Private Timer_ID As Long                                  '正在运行动画效果的Timer_ID
Private Timer_Time As Long                                '已运行时间
Private Timer_StartValue As Long, Timer_EndValue As Long  '起始\结束坐标

Public MouseInside As Boolean '记录光标是否在UserControl内

Private Sub UserControl_Initialize()
    FirstTipLeft = Dis_DefultFirstTip
    Timer1.Interval = Timer_Interval
    Timer1.Enabled = False
End Sub

Private Sub UserControl_Resize()
    '以下过程依赖UserControl.ScaleHeight，不可放在Initialize过程中
    XButton.Top = 0
    XButton.Width = XImg(0).Width
    XButton.Height = UserControl.ScaleHeight
    PrepareXImg
End Sub


Private Sub Timer_Start(ByVal ID As Long, ByVal StartValue As Long, ByVal EndValue As Long)
    Timer_ID = ID
    Timer_StartValue = StartValue
    Timer_EndValue = EndValue
    Timer_Time = 0
    Timer1.Enabled = True
End Sub

Private Sub ChangePointID(ByVal ID As Long)
    If PointID <> ActiveID And PointID <> 0 Then
        If Tip(PointID).BackColor <> Color_Back_Normal Then
            DrawTip PointID, Color_Back_Normal, Tip(PointID).ForeColor, Tip(PointID).FontBold
        End If
    End If
    If ID Then
        If ID <> ActiveID Then
            If Tip(ID).BackColor <> Color_Back_Point Then
                DrawTip ID, Color_Back_Point, Tip(ID).ForeColor, Tip(ID).FontBold
            End If
        End If
    End If
    PointID = ID
End Sub

'监测光标是否移出UserControl的Timer
Private Sub Timer_Cursor_Timer()
    Dim pt As POINTAPI
    If GetCursorPos(pt) <> 0 Then
        ScreenToClient UserControl.hWnd, pt
        If pt.x < 0 Or pt.x > UserControl.ScaleWidth Or pt.y < 0 Or pt.y > UserControl.ScaleHeight Then
            MouseInside = False
            If Drag_State = Drag_ing Then
                Drag_Drop pt.x
            End If
            ChangePointID 0
            If X_State <> Normal Then
                XButton.Picture = XImg(TB_X_Style.Normal).Image
                X_State = Normal
            End If
            Timer_Cursor.Enabled = False
        End If
    End If
End Sub

'进行移动操作的Timer
Private Sub Timer1_Timer()
    Timer_Time = Timer_Time + Timer_Interval
    If Timer_Time >= Timer_TotalTime Then
        SetTipLeft Timer_ID, Timer_EndValue
        Timer1.Enabled = False
        Timer_ID = 0
    Else
        SetTipLeft Timer_ID, Timer_StartValue + (Timer_EndValue - Timer_StartValue) * Timer_Time / Timer_TotalTime
    End If
End Sub

'设置Tip.Left,以及XButton.Left
Private Sub SetTipLeft(ByVal ID As Long, ByVal Left As Long)
    Tip(ID).Left = Left
    If ActiveID = ID Then
        XButton.Left = Tip(ID).Left + Tip(ID).Width - Dis_XButton_Cut - 1
    End If
End Sub

'根据Queue安排元素，但不设置Drag_ID(由UserControl_MouseMove事件驱动)和IDToAdjust的坐标,并返回后者坐标(用于Timer驱动)
Private Function Adjust(Optional ByVal IDToAdjust As Long) As Long

    If Count = 0 Then Exit Function '否则计算p时会出现除数为0

    Dim l As Long
    l = FirstTipLeft
    
    Dim i As Long, w As Long, p As Double
    
    '假设完全显示,统计总长度
    For i = 1 To Count
        w = w + Tip(Queue(i)).TextWidth(aCaption(Queue(i))) + Dis_TextBorder_X * 2 - 1
    Next
    If ActiveID <> 0 Then w = w - Dis_XButton_Cut + XButton.Width - 1
    
    '如果不能完全显示,各项按百分比缩小,否则百分比为1
    p = CDbl(UserControl.ScaleWidth) / CDbl(w)
    If p > 1 Then p = 1
    
    For i = 1 To Count
        SetTipWidth Queue(i), p
        If Queue(i) = IDToAdjust Then
            Adjust = l
        ElseIf Drag_State <> Drag_ing Or Queue(i) <> Drag_ID Then
            SetTipLeft Queue(i), l
        End If
        l = l + Tip(Queue(i)).Width - 1 '切除(覆盖)后边界
        If Queue(i) = ActiveID Then l = l - Dis_XButton_Cut + XButton.Width - 1 '由于XButton不带边框，补回下一项的前边框
    Next
    
End Function

'设置Tips.Width(期间可能涉及Tip重绘)
Private Sub SetTipWidth(ByVal ID As Long, ByVal Precentage As Double)
    Dim t As Long
    t = (Tip(ID).TextWidth(aCaption(ID)) * Precentage) + Dis_TextBorder_X * 2
    If t <> Tip(ID).Width Then
        Tip(ID).Width = t
        DrawTip ID, Tip(ID).BackColor, Tip(ID).ForeColor, Tip(ID).FontBold
    End If
End Sub

'Tip重绘
Public Sub DrawTip(ByVal ID As Long, ByVal BackColor As Long, ByVal ForeColor As Long, ByVal Bold As Boolean)
    With Tip(ID)
        .Cls
        .BackColor = BackColor
        .ForeColor = ForeColor
        .FontBold = Bold
        .CurrentX = Dis_TextBorder_X
        .CurrentY = (.ScaleHeight - .TextHeight(aCaption(ID))) / 2
        Tip(ID).Print aCaption(ID)
    End With
End Sub

Public Sub Activate(ByVal ID As Long)

    If ID <= 0 Or ID > UID Or ID = ActiveID Then Exit Sub
    
    If ActiveID <> 0 Then
        DrawTip ActiveID, Color_Back_Normal, Color_Fonts_Normal, False
    End If
    
    DrawTip ID, Color_Back_Active, Color_Fonts_Active, True
    
    ActiveID = ID
    DoEvents
    Adjust
    
    XButton.Picture = XImg(TB_X_Style.Normal).Image
    XButton.ZOrder 0
    
End Sub

Public Sub Add(ByVal Caption As String, ByVal Key As Long, Optional ByVal Active As Boolean = True)
    
    Count = Count + 1
    
    UID = UID + 1
    Load Tip(UID)

    aCaption(UID) = Caption

    With Tip(UID)
    
        .Top = -1
        .Height = UserControl.ScaleHeight + 2
        
        .Visible = True
        
        .Tag = Key
        
    End With
    
    Queue(Count) = UID
    
    If Active Then
        Activate UID
    Else
        DrawTip UID, Color_Back_Normal, Color_Fonts_Normal, False
        Adjust
    End If
    
    XButton.Visible = True
    XButton.ZOrder 0
    
End Sub

Public Sub Remove(ByVal ID As Long)
    
    If PointID = ID Then
        PointID = 0
    End If
    If ActiveID = ID Then
        ActiveID = 0
        If NextID(ID) <> 0 Then
            Activate NextID(ID)
        ElseIf PreviousID(ID) <> 0 Then
            Activate PreviousID(ID)
        End If
    End If
    
    '转移激活需要用到Queue()，Unload后置
    
    Unload Tip(ID)
    
    Dim i As Long
    For i = GetQueue(ID) To Count - 1
        Queue(i) = Queue(i + 1)
    Next
    Count = Count - 1
    
    XButton.Visible = (Count > 0)
    
    Adjust
    
End Sub

Private Property Get GetQueue(ByVal ID As Long)
    Dim i As Long
    For i = 1 To Count
        If Queue(i) = ID Then
            GetQueue = i
            Exit For
        End If
    Next
End Property

Public Property Get NextID(ByVal ID As Long)
    Dim i As Long
    i = GetQueue(ID)
    If i < Count Then NextID = Queue(i + 1)
End Property

Public Property Get PreviousID(ByVal ID As Long)
    Dim i As Long
    i = GetQueue(ID)
    If i > 1 Then PreviousID = Queue(i - 1)
End Property

Private Sub Drag_Start(ByVal x As Long)

    Drag_State = Drag_ing
    
    SetTipLeft Drag_ID, x - Drag_CursorLeft
    
    '移动项置顶
    Tip(Drag_ID).ZOrder 0
    If Drag_ID = ActiveID Then XButton.ZOrder 0
    
End Sub

Private Sub Drag_Swap(ByVal ID As Long)
    
    Dim NowLeft As Long
    NowLeft = Adjust(ID)
    
    Dim a As Long, b As Long
    a = GetQueue(Drag_ID)
    b = GetQueue(ID)
    Queue(a) = ID
    Queue(b) = Drag_ID
    
    Dim NewLeft As Long
    NewLeft = Adjust(ID)
    
    Timer_Start ID, NowLeft, NewLeft
    
End Sub

Private Sub Drag_Drop(ByVal x As Long)

    Timer_Start Drag_ID, x - Drag_CursorLeft, Adjust(Drag_ID)

    Drag_State = Drag_None
    Drag_ID = 0
    
End Sub

Private Sub Tip_MouseDown(ID As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If Drag_State = Drag_None And Timer_ID = 0 Then
            Drag_State = Drag_Pressed
            Drag_ID = ID
            Drag_ID_Left = Adjust(ID)
            Drag_CursorLeft = x
        End If
    End If
End Sub

Private Sub Tip_MouseMove(ID As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Drag_State = Drag_None Then
        If PointID <> ID Then
            ChangePointID ID
        End If
    Else
        '当Drag_State<>Drag_None时,统一转由UserControl处理,以便处理跨Tip移动情况,下同
        Call UserControl_MouseMove(Button, Shift, Tip(ID).Left + x, Tip(ID).Top + y)
    End If
    If X_State <> Normal Then
        XButton.Picture = XImg(TB_X_Style.Normal).Image
        X_State = Normal
    End If
End Sub

Private Sub Tip_MouseUp(ID As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If x > Tip(ID).ScaleWidth Or y > Tip(ID).ScaleHeight Then Exit Sub
    If Drag_State <> Drag_ing Then
        If Button = vbLeftButton Then
            RaiseEvent TipClick(ID)
        ElseIf Button = vbRightButton Then
            If PointID <> ID Then
                ChangePointID ID
            End If
            mnuClickID = ID
            UserControl.PopupMenu mnuTipsBar
            mnuCloseThis.Visible = True
            mnuCloseOthers.Visible = True
            mnuCloseLeft.Visible = True
            mnuCloseRight.Visible = True
            mnuCloseAll.Visible = True
        End If
        Drag_State = Drag_None
        Drag_ID = 0
    Else
        Call UserControl_MouseUp(Button, Shift, Tip(ID).Left + x, Tip(ID).Top + y)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MouseInside = False Then
        Timer_Cursor.Enabled = True
        MouseInside = True
    End If
    If Drag_State = Drag_None Then
        If X_State <> Normal Then
            XButton.Picture = XImg(TB_X_Style.Normal).Image
            X_State = Normal
        End If
        ChangePointID 0
    ElseIf Drag_State = Drag_Pressed Then
        If Abs((x - Drag_CursorLeft) - Drag_ID_Left) > Drag_StartValue Then
            Drag_Start x
        End If
    Else
        SetTipLeft Drag_ID, x - Drag_CursorLeft
        Dim NID As Long, PID As Long
        NID = NextID(Drag_ID): PID = PreviousID(Drag_ID)
        If NID <> 0 And x > Tip(NID).Left + Tip(NID).Width * 0.5 Then
            Drag_Swap NID
        ElseIf PID <> 0 And x < Tip(PID).Left + Tip(PID).Width * 0.5 Then
            Drag_Swap PID
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        UserControl.PopupMenu mnuTipsBar
        mnuCloseThis.Visible = False
        mnuCloseOthers.Visible = False
        mnuCloseLeft.Visible = False
        mnuCloseRight.Visible = False
        mnuCloseAll.Visible = True
    Else
        If Drag_State = Drag_ing Then
            Drag_Drop x
        End If
    End If
End Sub

Private Sub XButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    XButton.Picture = XImg(TB_X_Style.Pressed).Image
End Sub

Private Sub XButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Drag_State <> Drag_ing Then
        If X_State = Normal Then
            XButton.Picture = XImg(TB_X_Style.Active).Image
            X_State = Active
        End If
        
        ChangePointID 0
    End If
End Sub

Private Sub XButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Drag_State <> Drag_ing Then
        XButton.Picture = XImg(TB_X_Style.Normal).Image
        X_State = Normal
        RaiseEvent TipClose(ActiveID)
    End If
End Sub




'代码绘制"X"图形,包括XButton的右边界
Private Sub PrepareXImg()

    Dim i As Long
    For i = 0 To 2
    
        XImg(i).Width = XButton.Width
        XImg(i).Height = XButton.Height
        XImg(i).BackColor = Color_Back_Active
        
        Dim BK As Long, x As Long, y As Long
    
        Select Case i
        Case TB_X_Style.Normal
            BK = Color_Back_Active
        Case TB_X_Style.Active
            BK = Color_XButton_BackColor
        Case TB_X_Style.Pressed
            BK = Color_XButton_BackColor_Pressed
        End Select
        
        With XImg(i)
        
            x = Dis_XButton_Middle_X
            y = .ScaleHeight / 2
            
            .FillStyle = 0
            .FillColor = BK
            XImg(i).Circle (x, y), Dis_XButton_R, BK
            .FillStyle = 1
                
            .DrawWidth = 2
            XImg(i).Line (x - Dis_XButton_L, y - Dis_XButton_L)-(x + Dis_XButton_L, y + Dis_XButton_L), Color_X
            XImg(i).Line (x - Dis_XButton_L, y + Dis_XButton_L)-(x + Dis_XButton_L, y - Dis_XButton_L), Color_X
            .DrawWidth = 1
        
            XImg(i).Line (.ScaleWidth - 1, -2)-(.ScaleWidth - 1, .ScaleHeight), &H80000006
        
            .Refresh
            
        End With
        
    Next
    
End Sub

Public Property Get Key(ByVal ID As Long) As Long
    On Error Resume Next
    Key = Tip(ID).Tag
End Property

Public Function FindIDByKey(ByVal Key As Long) As Long
    FindIDByKey = -1
    Dim i As Long
    For i = 1 To Count
        If CLng(Tip(Queue(i)).Tag) = Key Then
            FindIDByKey = Queue(i)
            Exit For
        End If
    Next
End Function




Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "返回一个句柄到(from Microsoft Windows)一个对象的窗口。"
    hWnd = UserControl.hWnd
End Property

Private Sub mnuCengDie_Click()
    RaiseEvent mnuCengDie
End Sub

'引入Target()，避免删除过程中Queue变动引发错误（事件驱动，同步性不强）
'引入专有mnuClickID，避免删除过程中PointID变动引发错误

Private Sub mnuCloseAll_Click()
    On Error Resume Next
    
    Dim i As Long, j As Long
    j = Count
    
    If j >= 1 Then
    
        Dim Target() As Long
        ReDim Target(1 To j)
        
        For i = 1 To j
            Target(i) = Queue(i)
        Next
        
        For i = 1 To j
            RaiseEvent TipClose(Target(i))
            DoEvents
        Next
        
    End If
End Sub

Private Sub mnuCloseLeft_Click()

    On Error Resume Next
    
    Dim i As Long, j As Long
    j = GetQueue(mnuClickID) - 1
    
    If j >= 1 Then
    
        Dim Target() As Long
        ReDim Target(1 To j)
        
        For i = 1 To j
            Target(i) = Queue(i)
        Next
        
        For i = 1 To j
            RaiseEvent TipClose(Target(i))
            DoEvents
        Next
        
    End If
    
End Sub

Private Sub mnuCloseOthers_Click()
    mnuCloseRight_Click
    mnuCloseLeft_Click
End Sub

Private Sub mnuCloseRight_Click()

    On Error Resume Next
    
    Dim i As Long, j As Long, k As Long
    j = GetQueue(mnuClickID) + 1
    k = Count
    
    If k >= j Then
    
        Dim Target() As Long
        ReDim Target(j To k)
        
        For i = j To k
            Target(i) = Queue(i)
        Next
        
        For i = j To k
            RaiseEvent TipClose(Target(i))
            DoEvents
        Next
        
    End If
    
End Sub

Private Sub mnuCloseThis_Click()
    RaiseEvent TipClose(PointID)
End Sub

Private Sub mnuShuiPing_Click()
    RaiseEvent mnuShuiPing
End Sub

Private Sub mnuChuiZhi_Click()
    RaiseEvent mnuChuiZhi
End Sub


