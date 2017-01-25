VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form CI_frmCodeInsert 
   Caption         =   "插入代码"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox ContainerProperty1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   150
      ScaleHeight     =   315
      ScaleWidth      =   6165
      TabIndex        =   33
      Top             =   1080
      Width           =   6165
      Begin VB.TextBox txtMemVarName 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1020
         TabIndex        =   35
         Text            =   "m_"
         Top             =   0
         Width           =   2115
      End
      Begin VB.TextBox txtDefaultValue 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   3930
         TabIndex        =   34
         Top             =   0
         Width           =   2115
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "属性变量名:"
         Height          =   180
         Left            =   0
         TabIndex        =   37
         Top             =   45
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "默认值:"
         Height          =   180
         Left            =   3270
         TabIndex        =   36
         Top             =   45
         Width           =   630
      End
   End
   Begin VB.Timer mTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2340
      Top             =   7830
   End
   Begin VB.PictureBox ContainerProperty2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   1440
      ScaleHeight     =   1875
      ScaleWidth      =   3405
      TabIndex        =   27
      Top             =   2430
      Width           =   3405
      Begin iCode_Project.SelectBox SB_PropertyType 
         Height          =   1515
         Left            =   1680
         TabIndex        =   28
         Top             =   300
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   2672
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin iCode_Project.SelectBox SB_LetSetScope 
         Height          =   1485
         Left            =   390
         TabIndex        =   29
         Top             =   330
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   2619
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   1
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Let:"
         Height          =   180
         Left            =   330
         TabIndex        =   32
         Top             =   30
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "属性:"
         Height          =   180
         Left            =   1740
         TabIndex        =   31
         Top             =   0
         Width           =   450
      End
      Begin MSForms.ToggleButton chkPropertyScopeLinked 
         Height          =   405
         Left            =   0
         TabIndex        =   30
         Top             =   870
         Width           =   315
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "556;714"
         Value           =   "1"
         PicturePosition =   262148
         Picture         =   "CI_frmCodeInsert.frx":0000
         FontName        =   "宋体"
         FontHeight      =   180
         FontCharSet     =   134
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
   End
   Begin VB.PictureBox ContainerType 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3180
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   23
      Top             =   630
      Width           =   3015
      Begin VB.ComboBox txtType 
         Height          =   300
         Left            =   1110
         TabIndex        =   25
         Top             =   0
         Width           =   1875
      End
      Begin VB.CheckBox chkTypeNew 
         Caption         =   "New"
         Height          =   225
         Left            =   480
         TabIndex        =   24
         Top             =   30
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Left            =   0
         TabIndex        =   26
         Top             =   37
         Width           =   450
      End
   End
   Begin VB.PictureBox ContainerParam 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   150
      ScaleHeight     =   645
      ScaleWidth      =   6105
      TabIndex        =   17
      Top             =   1620
      Width           =   6105
      Begin iCode_Project.SelectBox SB_PAddTransfer 
         Height          =   285
         Left            =   1800
         TabIndex        =   38
         Top             =   330
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         BackColor       =   -2147483643
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
      Begin VB.TextBox txtParam 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   0
         Width           =   5445
      End
      Begin VB.CheckBox chkPAddOptional 
         Caption         =   "Optional"
         Height          =   285
         Left            =   630
         TabIndex        =   20
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtPAdd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3150
         TabIndex        =   19
         Top             =   330
         Width           =   2535
      End
      Begin VB.CommandButton btnPAdd 
         Caption         =   "+"
         Height          =   285
         Left            =   5760
         TabIndex        =   18
         Top             =   330
         Width           =   285
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "参数:"
         Height          =   180
         Left            =   0
         TabIndex        =   22
         Top             =   30
         Width           =   450
      End
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   5220
      TabIndex        =   12
      Top             =   3810
      Width           =   945
   End
   Begin VB.CommandButton btnInsertAndNew 
      Caption         =   "下一个(&N)"
      Height          =   345
      Left            =   5220
      TabIndex        =   11
      Top             =   3270
      Width           =   945
   End
   Begin VB.CommandButton btnComplete 
      Caption         =   "完成(&I)"
      Height          =   345
      Left            =   5220
      TabIndex        =   10
      Top             =   2730
      Width           =   945
   End
   Begin VB.Frame frmSetting 
      Caption         =   "附加"
      Height          =   2355
      Left            =   90
      TabIndex        =   6
      Top             =   4470
      Width           =   6135
      Begin VB.CheckBox chkMarkAsDefault 
         Caption         =   "作为省缺属性"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1920
         Width           =   3945
      End
      Begin VB.CheckBox chkPUsePropBag 
         Caption         =   "作为用户控件的属性使用 （使用 PropBag ）"
         Height          =   285
         Left            =   180
         TabIndex        =   15
         Top             =   1620
         Width           =   5055
      End
      Begin VB.CheckBox chkPUseInitProperties 
         Caption         =   "使用InitProperties代替Initalize"
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   1290
         Width           =   4605
      End
      Begin VB.CheckBox chkPDestoryWhenUnLoad 
         Caption         =   "实例终止()时,删除变量"
         Enabled         =   0   'False
         Height          =   315
         Left            =   450
         TabIndex        =   13
         Top             =   960
         Width           =   5625
      End
      Begin VB.CheckBox chkPInitWhenStartUp 
         Caption         =   "实例初始化()时,创建变量 "
         Enabled         =   0   'False
         Height          =   315
         Left            =   450
         TabIndex        =   9
         Top             =   630
         Width           =   5625
      End
      Begin iCode_Project.SelectBox SB_ProprrtyLetOrSet 
         Height          =   255
         Left            =   3300
         TabIndex        =   7
         Top             =   330
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   450
         BackColor       =   -2147483643
         Enabled         =   0   'False
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
      Begin VB.CheckBox chkPIsObject 
         Caption         =   "这是一个 Object 对象 （将会使用                    ）"
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   5265
      End
   End
   Begin iCode_Project.SelectBox SB_Scope 
      Height          =   1485
      Left            =   180
      TabIndex        =   4
      Top             =   2790
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2619
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   720
      TabIndex        =   2
      Top             =   630
      Width           =   2295
   End
   Begin iCode_Project.SelectBox SB_Type 
      Height          =   285
      Left            =   780
      TabIndex        =   0
      Top             =   150
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   503
      BackColor       =   -2147483643
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "类型:"
      Height          =   180
      Left            =   180
      TabIndex        =   5
      Top             =   210
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Get:"
      Height          =   180
      Left            =   150
      TabIndex        =   3
      Top             =   2490
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   660
      Width           =   450
   End
End
Attribute VB_Name = "CI_frmCodeInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const V1PerV0 = 0.01
Private Const tTotal = 400

Private mType1() As Single, mProperty1() As Single, mForm1() As Single
Private mType2() As Single, mProperty2() As Single, mForm2() As Single

Private t As Long, tStep As Long

Private mbType As Boolean, mbProperty As Boolean, mbForm As Boolean

Private sPartsName  As String, nPartsType As Long

Private Sub Complete()
    
End Sub

Private Sub btnCancel_Click()
    UnLoad Me
End Sub

Private Sub btnComplete_Click()
    Complete
    UnLoad Me
End Sub

Private Sub btnInsertAndNew_Click()
    Complete
    If SB_Type.SelectIndex <> 3 Then SB_Type_ItemMouseUp 3, SB_Type.Items(3), 0, 0, 0, 0
    Me.Pop sPartsName, nPartsType
End Sub

Private Sub btnPAdd_Click()
    If txtPAdd = "" Then Exit Sub
    If txtParam <> "" Then txtParam = txtParam & ", "
    If chkPAddOptional.Value = 1 Then txtParam = txtParam & "Optional "
    txtParam = txtParam & SB_PAddTransfer.Selection & " "
    txtParam = txtParam & txtPAdd
    txtPAdd = ""
End Sub

Private Sub chkPIsObject_Click()
    SB_ProprrtyLetOrSet.SelectItem chkPIsObject.Value + 1
    chkPInitWhenStartUp.Enabled = chkPIsObject.Value
    chkPDestoryWhenUnLoad.Enabled = chkPIsObject.Value
    Updata_chkPInitWhenStartUp
    Updata_chkPDestoryWhenUnLoad
End Sub

Private Sub chkPropertyScopeLinked_Click()
    If chkPropertyScopeLinked.Value = True Then SB_LetSetScope.SelectItem SB_Scope.SelectIndex
End Sub

Private Sub chkTypeNew_Click()
    If chkTypeNew.Value = 1 Then
        chkPIsObject.Value = 1
    End If
End Sub

Private Sub LoadValues()
    'GDIP.InitGDIPlus
    
    'SB_Type.Init
    SB_Type.AddItem "Sub"
    SB_Type.AddItem "Function"
    SB_Type.AddItem "Property"
    
    'SB_Scope.Init
    SB_Scope.AddItem "Public"
    SB_Scope.AddItem "Private"
    SB_Scope.AddItem "Friend"
    
    'SB_LetSetScope.Init
    SB_LetSetScope.AddItem "Public"
    SB_LetSetScope.AddItem "Private"
    SB_LetSetScope.AddItem "Friend"
    SB_LetSetScope.SelectItem 1
    SB_LetSetScope.Draw
    
    'SB_PropertyType.Init
    SB_PropertyType.AddItem "Get & Let/Set"
    SB_PropertyType.AddItem "Get          "
    SB_PropertyType.AddItem "Let/Set      "
    
    'SB_ProprrtyLetOrSet.Init
    SB_ProprrtyLetOrSet.AddItem "Let"
    SB_ProprrtyLetOrSet.AddItem "Set"
    
    'SB_PAddTransfer.Init
    SB_PAddTransfer.AddItem "ByVal"
    SB_PAddTransfer.AddItem "ByRef"
    
    'chkPInitWhenStartUp.Caption = Replace(chkPInitWhenStartUp.Caption, "变量", "变量" & vbCrLf)
    'chkPDestoryWhenUnLoad.Caption = Replace(chkPDestoryWhenUnLoad.Caption, "变量", "变量" & vbCrLf)
    
    
    LoadMoveInfo ContainerType.Left, Me.ScaleWidth + 3, mType1(), mType2()
    LoadMoveInfo ContainerProperty1.Left, -ContainerProperty1.Width - 3, mProperty1(), mProperty2()
    LoadMoveInfo Me.Height, 5000, mForm1(), mForm2()
End Sub

Private Sub Form_Load()
    LoadValues
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'SB_Type.Terminate
    'SB_Scope.Terminate
    'SB_LetSetScope.Terminate
    'SB_PropertyType.Terminate
    'SB_ProprrtyLetOrSet.Terminate
    'SB_PAddTransfer.Terminate
    
    'GDIP.TerminateGDIPlus
End Sub

Private Sub mTimer_Timer()
    t = t + 1
    
    If t >= (tTotal / mTimer.Interval) Then
        mbType = False
        mbProperty = False
        mbForm = False
        mTimer.Enabled = False
    End If
    
    If tStep > 0 Then
        
        If mbType = True Then
            ContainerType.Left = mType1(t)
        End If
        
        If mbProperty = True Then
            ContainerProperty1.Left = mProperty1(t)
            frmSetting.Left = ContainerProperty1.Left
            ContainerProperty2.Left = ContainerProperty1.Left + 86
        End If
        
        If mbForm = True Then
            Me.Height = mForm1(t)
        End If
        
    ElseIf tStep < 0 Then
        
        If mbType = True Then
            ContainerType.Left = mType2(t)
        End If
        
        If mbProperty = True Then
            ContainerProperty1.Left = mProperty2(t)
            frmSetting.Left = ContainerProperty1.Left
            ContainerProperty2.Left = ContainerProperty1.Left + 86
        End If
        
        If mbForm = True Then
            Me.Height = mForm2(t)
        End If
        
    End If
    
End Sub

Private Sub SB_LetSetScope_ItemMouseUp(Index As Long, item As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkPropertyScopeLinked.Value = True Then SB_Scope.SelectItem Index
End Sub

Private Sub SB_Scope_ItemMouseUp(Index As Long, item As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkPropertyScopeLinked.Value = True Then SB_LetSetScope.SelectItem Index
End Sub

Private Sub LoadMoveInfo(ByVal nStart As Single, ByVal nEnd As Single, ByRef mArray1() As Single, ByRef mArray2() As Single)
    Dim tTimes As Long: tTimes = tTotal / mTimer.Interval
    
    ReDim mArray1(tTimes)
    ReDim mArray2(tTimes)
    
    
    Dim mDis As Single, V0 As Single, a As Single, t As Long
    
    mDis = nEnd - nStart
    V0 = 2 * mDis / (tTimes) / (1 + V1PerV0)
    a = (V1PerV0 - 1) * V0 / (tTimes)
    
    mArray1(1) = nStart
    For t = 2 To tTimes - 1
        mArray1(t) = nStart + (V0 * t + (a * t * t) / 2)
    Next
    mArray1(tTimes) = nEnd
    
    mDis = nStart - nEnd
    V0 = 2 * mDis / tTimes / (1 + V1PerV0)
    a = (V1PerV0 - 1) * V0 / tTimes
    
    mArray2(1) = nEnd
    For t = 2 To tTimes - 1
        mArray2(t) = nEnd + (V0 * t + (a * t * t) / 2)
    Next
    mArray2(tTimes) = nStart
End Sub

Private Sub SB_Type_ItemMouseUp(Index As Long, item As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SB_Type.SelectIndex = -1 Then Index = 0: Exit Sub
    If Index = SB_Type.SelectIndex Then Exit Sub
    
    If Index = 1 Then
        
        mbType = True
        
        If SB_Type.SelectIndex = 3 Then
            mbProperty = True
            mbForm = True
        End If
        
        Label3.Caption = "范围:"
        
        tStep = 1
        
    ElseIf Index = 2 Then
        
        If SB_Type.SelectIndex = 1 Then
            mbType = True
            tStep = -1
        ElseIf SB_Type.SelectIndex = 3 Then
            mbProperty = True
            mbForm = True
            tStep = 1
        End If
        
        Label3.Caption = "范围:"
        
    ElseIf Index = 3 Then
        
        If SB_Type.SelectIndex = 1 Then
            mbType = True
        End If
        
        mbProperty = True
        mbForm = True
        
        Label3.Caption = "Get:"
        
        tStep = -1
        
    End If
    
    t = 0
    mTimer.Enabled = True
    
End Sub

Private Sub txtName_Change()
    txtMemVarName.Text = "m_" & txtName.Text
    Updata_chkPInitWhenStartUp
    Updata_chkPDestoryWhenUnLoad
End Sub

Private Sub Updata_chkPInitWhenStartUp()
    Dim i As Long
    i = InStr(1, chkPInitWhenStartUp.Caption, "(Set")
    
    If i = 0 Then
        If txtName <> "" And txtType.Text <> "" And chkPInitWhenStartUp.Enabled = True Then
            chkPInitWhenStartUp.Caption = chkPInitWhenStartUp.Caption & "(Set " & txtName & " = New " & txtType.Text & ")"
        End If
    Else
        If txtName <> "" And txtType.Text <> "" And chkPInitWhenStartUp.Enabled = True Then
            chkPInitWhenStartUp.Caption = Left(chkPInitWhenStartUp.Caption, i - 1) & "(Set " & txtName & " = New " & txtType.Text & ")"
        Else
            chkPInitWhenStartUp.Caption = Left(chkPInitWhenStartUp.Caption, i - 1)
        End If
    End If
End Sub

Private Sub Updata_chkPDestoryWhenUnLoad()
    Dim i As Long
    i = InStr(1, chkPDestoryWhenUnLoad.Caption, "(Set")
    
    If i = 0 Then
        If txtName <> "" And chkPDestoryWhenUnLoad.Enabled = True Then
            chkPDestoryWhenUnLoad.Caption = chkPDestoryWhenUnLoad.Caption & "(Set " & txtName & " = Nothing)"
        End If
    Else
        If txtName <> "" And chkPDestoryWhenUnLoad.Enabled = True Then
            chkPDestoryWhenUnLoad.Caption = Left(chkPDestoryWhenUnLoad.Caption, i - 1) & "(Set " & txtName & " = Nothing)"
        Else
            chkPDestoryWhenUnLoad.Caption = Left(chkPDestoryWhenUnLoad.Caption, i - 1)
        End If
    End If
    
End Sub

Public Sub Pop(ByVal PartsName As String, ByVal PartsType As String)
    
    Me.Show
    
    sPartsName = PartsName
    nPartsType = PartsType
    
    chkTypeNew.Value = 0
    txtType.Text = "Long"
    txtName.Text = ""
    chkPIsObject.Value = 0
    txtParam = ""
    chkPAddOptional.Value = 0
    txtPAdd = ""
    chkPropertyScopeLinked.Value = True
    chkPUseInitProperties.Value = 0
    chkPUsePropBag.Value = 0
    chkMarkAsDefault.Value = 0
    
    chkPUseInitProperties.Enabled = (PartsType = 8)
    chkPUsePropBag.Enabled = (PartsType = 8)
    
    chkPInitWhenStartUp.Caption = Replace(chkPInitWhenStartUp.Caption, "()", "(" & PartsName & "_Initialize" & ")")
    chkPDestoryWhenUnLoad.Caption = Replace(chkPDestoryWhenUnLoad.Caption, "()", "(" & PartsName & "_Terminate" & ")")
    
    SB_Type.SelectItem 3
    SB_Type.Draw
    
    SB_Scope.SelectItem 1
    SB_Scope.Draw
    
    SB_PropertyType.SelectItem 1
    SB_PropertyType.Draw
    
    SB_ProprrtyLetOrSet.SelectItem 1
    SB_ProprrtyLetOrSet.Draw
    
    
    SB_PAddTransfer.SelectItem 1
    SB_PAddTransfer.Draw
    
    
End Sub

Private Sub txtType_Change()
    Updata_chkPInitWhenStartUp
End Sub
