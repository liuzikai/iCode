VERSION 5.00
Begin VB.Form frmOperator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operator"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   2565
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   285
      Left            =   270
      TabIndex        =   9
      Top             =   1830
      Width           =   825
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   285
      Left            =   1620
      TabIndex        =   8
      Top             =   600
      Width           =   585
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   285
      Left            =   1650
      TabIndex        =   7
      Top             =   180
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ViewCB"
      Height          =   240
      Left            =   270
      TabIndex        =   6
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SortCode"
      Height          =   200
      Left            =   180
      TabIndex        =   5
      Top             =   420
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "¡ò"
      Height          =   315
      Left            =   510
      TabIndex        =   4
      Top             =   1020
      Width           =   315
   End
   Begin VB.CommandButton Command4 
      Caption         =   "¡ý"
      Height          =   315
      Left            =   510
      TabIndex        =   3
      Top             =   1380
      Width           =   315
   End
   Begin VB.CommandButton Command5 
      Caption         =   "¡ú"
      Height          =   315
      Left            =   870
      TabIndex        =   2
      Top             =   1020
      Width           =   315
   End
   Begin VB.CommandButton Command6 
      Caption         =   "¡û"
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   1020
      Width           =   315
   End
   Begin VB.CommandButton Command7 
      Caption         =   "¡ü"
      Height          =   315
      Left            =   510
      TabIndex        =   0
      Top             =   660
      Width           =   315
   End
End
Attribute VB_Name = "frmOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
    frmDebug.txtMain.Text = ""
    Dim i As Long, j As Long
    For i = 1 To VBIns.CommandBars.Count
        DBPrint VBIns.CommandBars.item(i).Name
        For j = 1 To VBIns.CommandBars.item(i).Controls.Count
            DBPrint "   " & VBIns.CommandBars.item(i).Controls(j).Caption & "(" & VBIns.CommandBars.item(i).Controls(j).id & ")"
        Next
    Next
    Clipboard.Clear
    Clipboard.SetText frmDebug.txtMain.Text
End Sub

Private Sub Command10_Click()
    iCode.AC.ACShow
End Sub

Private Sub Command2_Click()
    iCode.CodeSort.SortMouseKeyEvent
End Sub

Private Sub Command3_Click()
    iCode.AC.ACShow
End Sub

Private Sub Command4_Click()
    iCode.AC.KeyAction vbKeyDown
End Sub

Private Sub Command5_Click()
    iCode.AC.KeyAction vbKeyRight
End Sub

Private Sub Command6_Click()
    iCode.AC.KeyAction vbKeyLeft
End Sub

Private Sub Command7_Click()
    iCode.AC.KeyAction vbKeyUp
End Sub

Private Sub Form_Load()
    Me.Show
    SetAlwaysOnTop Me.hWnd, True
End Sub
