VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "iCode 配色方案集设置"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4890
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   326
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame framePlan 
      Caption         =   "Frame1"
      Height          =   5115
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   4635
      Begin VB.TextBox txtCaption 
         Height          =   270
         Left            =   840
         TabIndex        =   43
         Top             =   240
         Width           =   3555
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1440
         TabIndex        =   42
         Text            =   "FFFFFFFF"
         Top             =   1440
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   41
         Text            =   "FFFFFFFF"
         Top             =   1800
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   1440
         TabIndex        =   40
         Text            =   "FFFFFFFF"
         Top             =   2160
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   1440
         TabIndex        =   39
         Text            =   "FFFFFFFF"
         Top             =   2520
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   1440
         TabIndex        =   38
         Text            =   "FFFFFFFF"
         Top             =   2880
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   1440
         TabIndex        =   37
         Text            =   "FFFFFFFF"
         Top             =   3240
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   1440
         TabIndex        =   36
         Text            =   "FFFFFFFF"
         Top             =   3600
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   1440
         TabIndex        =   35
         Text            =   "FFFFFFFF"
         Top             =   3960
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   1440
         TabIndex        =   34
         Text            =   "FFFFFFFF"
         Top             =   4320
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   9
         Left            =   1440
         TabIndex        =   33
         Text            =   "FFFFFFFF"
         Top             =   4680
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   10
         Left            =   2520
         TabIndex        =   32
         Text            =   "FFFFFFFF"
         Top             =   1440
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   2520
         TabIndex        =   31
         Text            =   "FFFFFFFF"
         Top             =   1800
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   2520
         TabIndex        =   30
         Text            =   "FFFFFFFF"
         Top             =   2160
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   13
         Left            =   2520
         TabIndex        =   29
         Text            =   "FFFFFFFF"
         Top             =   2520
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   14
         Left            =   2520
         TabIndex        =   28
         Text            =   "FFFFFFFF"
         Top             =   2880
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   15
         Left            =   2520
         TabIndex        =   27
         Text            =   "FFFFFFFF"
         Top             =   3240
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   16
         Left            =   2520
         TabIndex        =   26
         Text            =   "FFFFFFFF"
         Top             =   3600
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   17
         Left            =   2520
         TabIndex        =   25
         Text            =   "FFFFFFFF"
         Top             =   3960
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   18
         Left            =   2520
         TabIndex        =   24
         Text            =   "FFFFFFFF"
         Top             =   4320
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   19
         Left            =   2520
         TabIndex        =   23
         Text            =   "FFFFFFFF"
         Top             =   4680
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   20
         Left            =   3600
         TabIndex        =   22
         Text            =   "FFFFFFFF"
         Top             =   1440
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   21
         Left            =   3600
         TabIndex        =   21
         Text            =   "FFFFFFFF"
         Top             =   1800
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   22
         Left            =   3600
         TabIndex        =   20
         Text            =   "FFFFFFFF"
         Top             =   2160
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   23
         Left            =   3600
         TabIndex        =   19
         Text            =   "FFFFFFFF"
         Top             =   2520
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   24
         Left            =   3600
         TabIndex        =   18
         Text            =   "FFFFFFFF"
         Top             =   2880
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   25
         Left            =   3600
         TabIndex        =   17
         Text            =   "FFFFFFFF"
         Top             =   3240
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   26
         Left            =   3600
         TabIndex        =   16
         Text            =   "FFFFFFFF"
         Top             =   3600
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   27
         Left            =   3600
         TabIndex        =   15
         Text            =   "FFFFFFFF"
         Top             =   3960
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   28
         Left            =   3600
         TabIndex        =   14
         Text            =   "FFFFFFFF"
         Top             =   4320
         Width           =   795
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   29
         Left            =   3600
         TabIndex        =   13
         Text            =   "FFFFFFFF"
         Top             =   4680
         Width           =   795
      End
      Begin VB.TextBox txtCreator 
         Height          =   270
         Left            =   1380
         TabIndex        =   12
         Top             =   600
         Width           =   3015
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   900
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "iCode配色方案设置（*.ic_c）|*.ic_c|所有文件|*.*"
      End
      Begin VB.Label lblTypeName 
         Alignment       =   2  'Center
         Caption         =   "前景色"
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   58
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label lblTypeName 
         Alignment       =   2  'Center
         Caption         =   "背景色"
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   57
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label lblTypeName 
         Alignment       =   2  'Center
         Caption         =   "标识色"
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   56
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "标准文本"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   55
         Top             =   1485
         Width           =   720
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "选定文本"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   54
         Top             =   1845
         Width           =   720
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "语法错误文本"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   53
         Top             =   2205
         Width           =   1080
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "执行点文本"
         Height          =   180
         Index           =   3
         Left            =   180
         TabIndex        =   52
         Top             =   2565
         Width           =   900
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "断点文本"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   51
         Top             =   2925
         Width           =   720
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "注释文本"
         Height          =   180
         Index           =   5
         Left            =   180
         TabIndex        =   50
         Top             =   3285
         Width           =   720
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "关键字文本"
         Height          =   180
         Index           =   6
         Left            =   180
         TabIndex        =   49
         Top             =   3645
         Width           =   900
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "标识符文本"
         Height          =   180
         Index           =   7
         Left            =   180
         TabIndex        =   48
         Top             =   4005
         Width           =   900
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "书签文本"
         Height          =   180
         Index           =   8
         Left            =   180
         TabIndex        =   47
         Top             =   4365
         Width           =   720
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "调用返回文本"
         Height          =   180
         Index           =   9
         Left            =   180
         TabIndex        =   46
         Top             =   4725
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名称："
         Height          =   180
         Left            =   180
         TabIndex        =   45
         Top             =   285
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "创建者标识："
         Height          =   180
         Left            =   180
         TabIndex        =   44
         Top             =   645
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdSafe 
      Caption         =   "保存所选到文件"
      Height          =   315
      Left            =   2100
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "从文件加载方案"
      Height          =   315
      Left            =   540
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "重置方案集"
      Height          =   435
      Left            =   3600
      TabIndex        =   6
      Top             =   7380
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存方案集"
      Height          =   435
      Left            =   2340
      TabIndex        =   5
      Top             =   7380
      Width           =   1155
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除方案"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   780
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "新建方案"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      ToolTipText     =   "以所选方案为模板创建新方案"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "↓"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   315
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "↑"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   315
   End
   Begin VB.ListBox listProg 
      Height          =   1140
      Left            =   540
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "“-1”即为“自动”"
      Height          =   180
      Left            =   3120
      TabIndex        =   8
      Top             =   7020
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "双击调用调色板"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   7020
      Width           =   1260
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Color_Count As Long
Private Colors(Color_Max) As Color_Set

Public clsParent As clsColorCode

Private Changed As Boolean

Private Sub cmdDelete_Click()
    Dim i As Long
    For i = listProg.ListIndex + 1 To Color_Count
        Colors(i - 1) = Colors(i)
        listProg.List(i - 1) = listProg.List(i)
        listProg.RemoveItem listProg.ListCount - 1
    Next
    Color_Count = Color_Count - 1
End Sub

Private Sub cmdDown_Click()
    If listProg.ListIndex < listProg.ListCount - 1 Then
        Colors(Color_Count + 1) = Colors(listProg.ListIndex + 1)
        Colors(listProg.ListIndex + 1) = Colors(listProg.ListIndex)
        Colors(listProg.ListIndex) = Colors(Color_Count + 1)
        listProg.ListIndex = listProg.ListIndex + 1
        SetProgList
        listProg.SetFocus
    End If
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo iErr
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" And Dir(CommonDialog1.FileName) <> "" Then
        Open CommonDialog1.FileName For Input As #1
        Dim i As Long, j As Long, s As String, p() As String
        Input #1, i
        For i = 1 To i
            Input #1, Colors(Color_Count + i).Caption
            Input #1, Colors(Color_Count + i).Creator
            Input #1, s
            p = Split(s, ",")
            For j = 0 To 29
                Colors(Color_Count + i).Color(j) = CLng(p(j))
            Next
        Next
        Color_Count = Color_Count + i
        Close #1
        SetProgList
    End If
    Exit Sub
iErr:
    MsgBox "读取文件错误！", vbExclamation, "iCode"
End Sub

Private Sub cmdNew_Click()

    If listProg.ListIndex >= 0 Then Colors(Color_Count) = Colors(listProg.ListIndex)
    Colors(Color_Count).Caption = "新建方案"
    
    listProg.AddItem Colors(Color_Count).Caption
    listProg.ListIndex = listProg.ListCount - 1
    
    Color_Count = Color_Count + 1
    
    Changed = True
    
End Sub

Private Sub cmdReload_Click()
    If Changed Then
        If MsgBox("还没有保存当前方案集，确定要重新读取吗？", vbInformation Or vbYesNo, "iCode") = vbYes Then
            Form_Load
        End If
    End If
End Sub

Private Sub cmdSafe_Click()
    If listProg.SelCount > 0 Then
        CommonDialog1.ShowSave
        If CommonDialog1.FileName <> "" Then
            Open CommonDialog1.FileName For Output As #1
            Dim i As Long, j As Long, s As String
            Write #1, listProg.SelCount
            For i = 0 To listProg.ListCount - 1
                If listProg.Selected(i) Then
                    Write #1, Colors(i).Caption
                    Write #1, Colors(i).Creator
                    s = ""
                    For j = 0 To 28
                        s = s & "&H" & Hex(Colors(i).Color(j)) & ","
                    Next
                    s = s & "&H" & Hex(Colors(i).Color(29))
                    Write #1, s
                End If
            Next
            Close #1
        End If
    Else
        MsgBox "还没有选择导出项！", vbInformation, "iCode"
    End If
End Sub

Private Sub cmdSave_Click()
    modPublic.Color_Count = Color_Count
    Dim i As Long
    For i = 0 To Color_Count - 1
        modPublic.Colors(i) = Colors(i)
    Next
    clsParent.WriteColors
    Changed = False
    MsgBox "保存完成！", vbInformation, "iCode"
End Sub

Private Sub cmdUp_Click()
    If listProg.ListIndex > 0 Then
        Colors(Color_Count + 1) = Colors(listProg.ListIndex - 1)
        Colors(listProg.ListIndex - 1) = Colors(listProg.ListIndex)
        Colors(listProg.ListIndex) = Colors(Color_Count + 1)
        listProg.ListIndex = listProg.ListIndex - 1
        SetProgList
        listProg.SetFocus
    End If
End Sub

Private Sub SetProgList()
    listProg.Clear
    Dim i As Long
    For i = 0 To Color_Count - 1
        listProg.AddItem Colors(i).Caption
    Next
End Sub

Private Sub Form_Load()
    Color_Count = modPublic.Color_Count
    Dim i As Long
    For i = 0 To Color_Count - 1
        Colors(i) = modPublic.Colors(i)
    Next
    SetProgList
    If listProg.ListCount >= 1 Then
        listProg.ListIndex = 0
        listProg_Click
    End If
    Changed = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Changed Then
        If MsgBox("还没有保存当前方案集，确定要退出吗？", vbInformation Or vbYesNo, "iCode") = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub listProg_Click()
    Dim i As Long, k As Boolean
    k = Changed '避免触发txtColor_Change导致Changed被改变
    If listProg.ListIndex <> -1 Then
        txtCaption = Colors(listProg.ListIndex).Caption
        framePlan.Caption = txtCaption
        txtCreator = Colors(listProg.ListIndex).Creator
        For i = 0 To 29
            If Colors(listProg.ListIndex).Color(i) = -1 Then
                txtColor(i) = "-1"
                txtColor(i).BackColor = vbWhite
            Else
                txtColor(i) = Hex(Colors(listProg.ListIndex).Color(i))
                txtColor(i).BackColor = Colors(listProg.ListIndex).Color(i)
            End If
            'txtColor(i).ForeColor的改变由txtColor_Change完成
            Call txtColor_LostFocus(CInt(i))
        Next
        cmdDelete.Enabled = (listProg.List(listProg.ListIndex) <> "默认")
    End If
    Changed = k
End Sub

Private Sub txtCaption_Change()
    Colors(listProg.ListIndex).Caption = txtCaption
    listProg.List(listProg.ListIndex) = txtCaption
End Sub

Private Sub txtColor_Change(Index As Integer)
    'On Error GoTo iErr
    If txtColor(Index) = "-1" Then
        txtColor(Index).BackColor = vbWhite
        Colors(listProg.ListIndex).Color(Index) = &HFFFFFFFF
    Else
        txtColor(Index).BackColor = CLng("&H" & txtColor(Index))
        Colors(listProg.ListIndex).Color(Index) = CLng("&H" & txtColor(Index))
    End If
    txtColor(Index).ForeColor = &HFFFFFF Xor txtColor(Index).BackColor
    Changed = True
    Exit Sub
iErr:
End Sub

Private Sub txtColor_DblClick(Index As Integer)
    CommonDialog1.Color = txtColor(Index).BackColor
    CommonDialog1.ShowColor
    txtColor(Index) = Hex(CommonDialog1.Color)
    txtColor(Index).BackColor = CommonDialog1.Color
End Sub

Private Sub txtColor_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("f") Then
        KeyAscii = KeyAscii - Asc("a") + Asc("A")
    ElseIf KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("F") Then
    ElseIf KeyAscii = Asc("-") Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtColor_LostFocus(Index As Integer)
    '补全6位数
    If txtColor(Index) <> "-1" And Len(txtColor(Index)) < 6 Then
        txtColor(Index) = String(6 - Len(txtColor(Index)), "0") & txtColor(Index)
    End If
End Sub

Private Sub txtCreator_Change()
    Colors(listProg.ListIndex).Creator = txtCreator
End Sub
