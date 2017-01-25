VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form HiddenForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   569
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   731
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSScriptControlCtl.ScriptControl CQ_SC 
      Left            =   660
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin iCode_Project.TipsBar TipsBar 
      Height          =   270
      Left            =   360
      Top             =   300
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   476
   End
   Begin VB.Timer EH_timerErrorBox 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   180
      Top             =   2400
   End
   Begin VB.Timer FW_timerFileWindow 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4.91505e5
      Top             =   465
   End
   Begin VB.Timer Code_timerSortMouseKeyEvents 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2010
      Top             =   2370
   End
   Begin VB.PictureBox cCodeGo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3390
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   5
      Top             =   2340
      Width           =   300
      Begin VB.CommandButton CodeGo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   0
         Picture         =   "frmHidden.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.PictureBox cCodeToCommand 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   270
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   3
      Top             =   1680
      Width           =   750
      Begin VB.CommandButton CodeToCommand 
         Caption         =   "→通用"
         Height          =   300
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   750
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   28140
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   0
      TabIndex        =   2
      Top             =   555
      Width           =   0
   End
   Begin VB.PictureBox cCodeBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   420
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   1290
      Width           =   300
      Begin VB.CommandButton CodeBack 
         Enabled         =   0   'False
         Height          =   300
         Left            =   0
         Picture         =   "frmHidden.frx":0344
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.Menu mnuTipsBar 
      Caption         =   "Tips"
      Begin VB.Menu mnuTipsClose 
         Caption         =   "关闭标签(&C)"
      End
      Begin VB.Menu mnuTipsCloseOther 
         Caption         =   "关闭其他(&O)"
      End
      Begin VB.Menu mnuTipsCloseLeft 
         Caption         =   "关闭左侧标签(&L)"
      End
      Begin VB.Menu mnuTipsCloseRight 
         Caption         =   "关闭右侧标签(&R)"
      End
      Begin VB.Menu mnuLine7 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTipsLock 
         Caption         =   "锁定(&K)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTipsMin 
         Caption         =   "最小化所有窗口(&I)"
      End
      Begin VB.Menu mnuTipsNormal 
         Caption         =   "常态化所有窗口(&N)"
      End
      Begin VB.Menu mnuTipsMax 
         Caption         =   "最大化所有窗口(&A)"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTipsShuiPing 
         Caption         =   "水平平铺窗口(&H)"
      End
      Begin VB.Menu mnuTipsChuiZhi 
         Caption         =   "垂直平铺窗口(&V)"
      End
      Begin VB.Menu mnuTipsCengDie 
         Caption         =   "层叠窗口(&D)"
      End
   End
   Begin VB.Menu mnuiProject 
      Caption         =   "iProject"
      Begin VB.Menu mnuProjectStart 
         Caption         =   "设置为启动(&U)"
      End
      Begin VB.Menu mnuProjectProperty 
         Caption         =   "工程属性(&P)"
      End
      Begin VB.Menu mnuViewObject 
         Caption         =   "查看对象(&O)"
      End
      Begin VB.Menu mnuViewCode 
         Caption         =   "查看代码(&C)"
      End
      Begin VB.Menu mnuFolderOpen 
         Caption         =   "展开(&E)"
      End
      Begin VB.Menu mnuFolderClose 
         Caption         =   "收缩(&C)"
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "添加(&A)"
         Begin VB.Menu mnuAddC 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuAddC 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuAddC 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuAddC 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuAddC 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuAddC 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuAddC 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuAddC 
            Caption         =   ""
            Index           =   13
         End
      End
      Begin VB.Menu mnuAddFolder 
         Caption         =   "添加文件夹(&F)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "保存 (&S)"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "另存为 (&V)"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "移除 (&R)"
      End
      Begin VB.Menu mnuRemoveFolder 
         Caption         =   "移除 (&R)"
         Visible         =   0   'False
         Begin VB.Menu mnuRemoveFolderDelete 
            Caption         =   "清除内容(&D)"
         End
         Begin VB.Menu mnuRemoveFolderUpdata 
            Caption         =   "移动内容至上一级(&U)"
         End
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "打印(&T)"
      End
      Begin VB.Menu mnuLine6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModeClassic 
         Caption         =   "标准模式(&N)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuModeI 
         Caption         =   "自定义模式(&I)"
      End
   End
End
Attribute VB_Name = "HiddenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TB_TipsIndex As Long
Public Code_hCodeWindow As Long
Public OV_hObjectViewer As Long
Public FW_Hwnd As Long
Public FW_Caption As String

Private TimerCount As Long

Private Sub EH_timerErrorBox_Timer()
    iCode.CWE.ErrBox.Show
    modPublic.SetFocus iCode.CWE.hButtonsContainer
    EH_timerErrorBox.Enabled = False
End Sub

Private Sub mnuAddC_Click(Index As Integer)
    iProject.mnuAddC_Click (Index)
End Sub

Private Sub mnuAddFolder_Click()
    iProject.mnuAddFolder_Click
End Sub

Private Sub mnuTipsCloseOther_Click()
    iTipsBar.DelectRight TB_TipsIndex
    iTipsBar.DelectLeft TB_TipsIndex
End Sub

Private Sub mnuFolderClose_Click()
    iProject.mnuFolderClose_Click
End Sub

Private Sub mnuFolderOpen_Click()
    iProject.mnuFolderOpen_Click
End Sub

Private Sub mnuModeClassic_Click()
    iProject.mnuModeClassic_Click
End Sub

Private Sub mnuModeI_Click()
    iProject.mnuModeI_Click
End Sub

Private Sub mnuPrint_Click()
    iProject.mnuPrint_Click
End Sub

Private Sub mnuProjectProperty_Click()
    iProject.mnuProjectProperty_Click
End Sub

Private Sub mnuProjectStart_Click()
    iProject.mnuProjectStart_Click
End Sub

Private Sub mnuRemove_Click()
    iProject.mnuRemove_Click
End Sub

Private Sub mnuRemoveFolderDelete_Click()
    iProject.mnuRemoveFolder_Click (0)
End Sub

Private Sub mnuRemoveFolderUpdata_Click()
    iProject.mnuRemoveFolder_Click (1)
End Sub

Private Sub mnuSave_Click()
    iProject.mnuSave_Click
End Sub

Private Sub mnuSaveAs_Click()
    iProject.mnuSaveAs_Click
End Sub

Private Sub mnuTipsClose_Click()
    SendMessage iTipsBar.TipsBar.Tips(TB_TipsIndex).Key, WM_CLOSE, 0, 0
End Sub

Private Sub mnuTipsCloseLeft_Click()
    iTipsBar.DelectLeft TB_TipsIndex
End Sub

Private Sub mnuTipsCloseRight_Click()
    iTipsBar.DelectRight TB_TipsIndex
End Sub

Private Sub mnuViewCode_Click()
    iProject.mnuViewCode_Click
End Sub

Private Sub mnuViewObject_Click()
    iProject.mnuViewObject_Click
End Sub


Private Sub Code_timerSortMouseKeyEvents_Timer()
    iCode.CodeSort.SortMouseKeyEvent
    Code_timerSortMouseKeyEvents.Enabled = False
End Sub

