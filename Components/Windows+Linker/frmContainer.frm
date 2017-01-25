VERSION 5.00
Begin VB.Form frmContainer 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   116
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Container 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   0
      Top             =   720
      Width           =   1815
      Begin VB.CommandButton cmdiCodeLinker 
         Caption         =   "iCode 编译工具..."
         Height          =   345
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdiCodeLinker_Click()

End Sub
