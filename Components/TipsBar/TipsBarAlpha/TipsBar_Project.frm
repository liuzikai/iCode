VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   11760
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   1740
      Top             =   1650
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   2070
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4620
      TabIndex        =   1
      Top             =   1620
      Width           =   915
   End
   Begin iCode_TipsBar_Project.TipsBar TipsBar1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11115
      _extentx        =   12726
      _extenty        =   529
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   405
      Left            =   870
      TabIndex        =   2
      Top             =   2700
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    TipsBar1.Add "ObjectViewer", Rnd() * 2000
End Sub

Private Sub Form_Load()
    Me.Show
    TipsBar1.Add "Win32 Declare", 1
    TipsBar1.Add "DesignerWindow", 2
    TipsBar1.Add "ObjectViewer", 3
End Sub

Private Sub TipsBar1_TipClick(ByVal ID As Long)
    TipsBar1.Activate ID
End Sub

Private Sub TipsBar1_TipClose(ByVal ID As Long)
    TipsBar1.Remove ID
End Sub
