VERSION 5.00
Begin VB.Form frmSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "iCode 设置"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5460
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   526
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame4 
      Caption         =   "代码缩进"
      Height          =   2475
      Left            =   60
      TabIndex        =   11
      Top             =   3900
      Width           =   5235
      Begin VB.Frame Frame6 
         Caption         =   "快捷按钮功能选择"
         Height          =   675
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   4995
         Begin VB.OptionButton CI_QuickButton_Mode 
            Caption         =   "不显示"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton CI_QuickButton_Mode 
            Caption         =   "整理工程"
            Height          =   195
            Index           =   3
            Left            =   3180
            TabIndex        =   21
            Top             =   300
            Width           =   1095
         End
         Begin VB.OptionButton CI_QuickButton_Mode 
            Caption         =   "整理模块"
            Height          =   195
            Index           =   2
            Left            =   2100
            TabIndex        =   20
            Top             =   300
            Width           =   1095
         End
         Begin VB.OptionButton CI_QuickButton_Mode 
            Caption         =   "整理过程"
            Height          =   195
            Index           =   1
            Left            =   1020
            TabIndex        =   19
            Top             =   300
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "多行代码"
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   4995
         Begin VB.TextBox CI_MultiLine_IndentSpaceCount 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1560
            TabIndex        =   17
            Text            =   "4"
            Top             =   555
            Width           =   315
         End
         Begin VB.OptionButton CI_MultiLine_Orginal 
            Caption         =   "保持相对位置不变"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   300
            Value           =   -1  'True
            Width           =   2115
         End
         Begin VB.OptionButton CI_MultiLine_Indent 
            Caption         =   "第二行起推进     个空格"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   2475
         End
      End
      Begin VB.TextBox CI_SpacePerLevel 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1140
         TabIndex        =   13
         Text            =   "4"
         Top             =   292
         Width           =   315
      End
      Begin VB.Label Label3 
         Caption         =   "每一级缩进     个空格"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "工作区标签栏"
      Height          =   735
      Left            =   60
      TabIndex        =   9
      Top             =   6480
      Width           =   5235
      Begin VB.CheckBox TBH_Enabled 
         Caption         =   "启用"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   360
         Value           =   1  'Checked
         Width           =   2475
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "IDE工具"
      Height          =   2835
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5175
      Begin VB.CheckBox IDEE_AddFile_Buttons_Enabled 
         Caption         =   "在“工程”菜单与“工程资源管理器”菜单中添加“添加文件(增强版)”按钮"
         Height          =   360
         Left            =   180
         TabIndex        =   24
         Top             =   1980
         Width           =   4815
      End
      Begin VB.CheckBox IDEE_MakeExeButton_Enabled 
         Caption         =   "在“工程资源管理器”中工程的右键菜单内添加编译按钮"
         Height          =   180
         Left            =   180
         TabIndex        =   23
         Top             =   1680
         Width           =   4815
      End
      Begin VB.CheckBox IDEE_Compile_Button_Visible 
         Caption         =   "显示“编译”按钮"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   900
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox IDEE_ToCommon_Button_Visible 
         Caption         =   "显示“→通用”按钮"
         Height          =   180
         Left            =   180
         TabIndex        =   6
         Top             =   600
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox IDEE_ChangeScope_Button_Visible 
         Caption         =   "显示“Private←→Public”按钮"
         Height          =   180
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "以上两个选项可能导致VB原生按钮消失，重启VB生效"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   2460
         Width           =   4755
      End
      Begin VB.Label Label1 
         Caption         =   "可通过 菜单栏 - 自定义 调整按钮样式"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   1200
         Width           =   3195
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   435
      Left            =   4440
      TabIndex        =   3
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认(&S)"
      Height          =   435
      Left            =   3420
      TabIndex        =   2
      Top             =   7320
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "文件窗口加强"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CheckBox W_Redesign 
         Caption         =   "启用文件窗口重布局"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If Not IsNumeric(CI_MultiLine_IndentSpaceCount) Then
        MsgBox "代码缩进：缩进量无效", vbExclamation, "错误"
        CI_MultiLine_IndentSpaceCount.SetFocus
        Exit Sub
    End If
    
    If CLng(CI_MultiLine_IndentSpaceCount) <= 0 Then
        MsgBox "代码缩进：缩进量需大于0", vbExclamation, "错误"
        CI_MultiLine_IndentSpaceCount.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(CI_SpacePerLevel) Then
        MsgBox "代码缩进：缩进量无效", vbExclamation, "错误"
        CI_SpacePerLevel.SetFocus
        Exit Sub
    End If
    
    If CLng(CI_SpacePerLevel) <= 0 Then
        MsgBox "代码缩进：缩进量需大于0", vbExclamation, "错误"
        CI_SpacePerLevel.SetFocus
        Exit Sub
    End If
    
    Settings_Write "Windows_Linker", "ResizeWindow", CStr(CBool(W_Redesign.Value))
    Windows_Linker.Mode = CBool(W_Redesign.Value)
    
    IDEEnhancer.ChangeScope_Button_Visible = CBool(IDEE_ChangeScope_Button_Visible.Value)
    IDEEnhancer.Compile_Button_Visible = CBool(IDEE_Compile_Button_Visible.Value)
    IDEEnhancer.ToCommon_Button_Visible = CBool(IDEE_ToCommon_Button_Visible.Value)
    IDEEnhancer.m_MakeExeButton_Enabled = CBool(IDEE_MakeExeButton_Enabled.Value)
    IDEEnhancer.m_AddFile_Buttons_Enabled = CBool(IDEE_AddFile_Buttons_Enabled.Value)
    
    Settings_Write "TipsBarHandler", "Enabled", CStr(CBool(TBH_Enabled.Value))
    TipsBarHandler.TipsBarAvliable = CBool(TBH_Enabled.Value)
    
    Settings_Write "CodeIndent", "SpacePerLevel", CI_SpacePerLevel
    CodeIndent.Var_SpacePerLevel = CI_SpacePerLevel
    Settings_Write "CodeIndent", "MulitLines_IndentSpaceCount", CI_MultiLine_IndentSpaceCount
    CodeIndent.Var_MulitLines_IndentSpaceCount = CI_MultiLine_IndentSpaceCount
    
    Dim i As Long
    For i = 0 To 3
        If CI_QuickButton_Mode(i).Value = True Then
            Settings_Write "CodeIndent", "MulitLines_QuickButtonMode", i
            CodeIndent.Var_QuickButtonMode = i
            Exit For
        End If
    Next
        
    Unload Me
    
End Sub

Private Sub Form_Load()
    Me.Show
    CI_SpacePerLevel = CodeIndent.Var_SpacePerLevel
    If CodeIndent.Var_MulitLines_IndentSpaceCount = 0 Then
        CI_MultiLine_Orginal.Value = True
    Else
        CI_MultiLine_Indent.Value = True
        CI_MultiLine_IndentSpaceCount = CodeIndent.Var_MulitLines_IndentSpaceCount
    End If
    CI_QuickButton_Mode(CodeIndent.Var_QuickButtonMode).Value = True
    TBH_Enabled.Value = -CLng(TipsBarHandler.TipsBarAvliable)
    W_Redesign.Value = -CLng(Windows_Linker.Mode)
    IDEE_ChangeScope_Button_Visible = -CLng(IDEEnhancer.m_ChangeScope_Button_Visible)
    IDEE_Compile_Button_Visible = -CLng(IDEEnhancer.m_Compile_Button_Visible)
    IDEE_ToCommon_Button_Visible = -CLng(IDEEnhancer.m_ToCommon_Button_Visible)
    IDEE_MakeExeButton_Enabled = -CLng(IDEEnhancer.m_MakeExeButton_Enabled)
    IDEE_AddFile_Buttons_Enabled = -CLng(IDEEnhancer.m_AddFile_Buttons_Enabled)
End Sub

