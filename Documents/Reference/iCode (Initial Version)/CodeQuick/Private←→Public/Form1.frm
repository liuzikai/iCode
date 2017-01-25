VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   8625
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   5535
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   360
      Width           =   7335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim s As String
    s = Text1.Text
    Do Until InStr(1, s, "Private") = 0
        s = Replace(s, "Private", "#3607#")
    Loop
    Do Until InStr(1, s, "Public") = 0
        s = Replace(s, "Public", "Private")
    Loop
    Do Until InStr(1, s, "#3607#") = 0
        s = Replace(s, "#3607#", "Public")
    Loop
    Text1.Text = s

End Sub

