VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "iCode"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "iCode"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   176.25
   ScaleMode       =   2  'Point
   ScaleWidth      =   228
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCopyEmail 
      Caption         =   "复制邮箱地址"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   3060
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3060
      Width           =   855
   End
   Begin VB.TextBox txtMain 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2835
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmAbout.frx":0442
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopyEmail_Click()
    Clipboard.SetText "liuzikai@163.com"
    MsgBox "已复制到剪切板！", 0, "iCode"
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
