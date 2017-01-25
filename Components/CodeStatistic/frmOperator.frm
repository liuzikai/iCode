VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "iCode 代码统计"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3975
   Icon            =   "frmOperator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   138
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.CheckBox chkProc 
      Caption         =   "统计成员（过程、属性、事件等）"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   420
      Value           =   1  'Checked
      Width           =   3555
   End
   Begin VB.CheckBox chkChar 
      Caption         =   "统计字符数"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   120
      Value           =   1  'Checked
      Width           =   3555
   End
   Begin MSComctlLib.ProgressBar pbTotal 
      Height          =   195
      Left            =   1020
      TabIndex        =   2
      Top             =   840
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "开始"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   1560
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar pbPart 
      Height          =   195
      Left            =   1020
      TabIndex        =   4
      Top             =   1200
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblPart 
      Caption         =   "分进度："
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblTotal 
      Caption         =   "总进度："
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frmOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Event Start()
Event Cancel()

Private Sub cmdButton_Click()
    If cmdButton.Caption = "开始" Then
        chkChar.Enabled = False
        chkProc.Enabled = False
        cmdButton.Caption = "取消"
        RaiseEvent Start
    Else
        Me.Hide
        RaiseEvent Cancel
    End If
End Sub

Public Sub iShow()
    pbTotal.Value = 0
    pbPart.Value = 0
    chkChar.Enabled = True
    chkProc.Enabled = True
    cmdButton.Caption = "开始"
    Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent Cancel
End Sub
