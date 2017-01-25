VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Debug"
   ClientHeight    =   4470
   ClientLeft      =   12780
   ClientTop       =   1335
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   210
   Begin VB.TextBox txtMain 
      Appearance      =   0  'Flat
      Height          =   4335
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   3030
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Show
    SetAlwaysOnTop Me.hWnd, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    VBIns.Quit
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtMain.Width = Me.ScaleWidth - txtMain.Left * 2
    txtMain.Height = Me.ScaleHeight - txtMain.Top * 2
End Sub

Private Sub txtMain_DblClick()
    txtMain.Text = ""
End Sub

Private Sub txtMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        frmOperator.Show
    End If
End Sub
