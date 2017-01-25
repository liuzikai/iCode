VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResult 
   Caption         =   "统计结果"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   Icon            =   "frmResult.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   264
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制到剪切板"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5580
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   435
      Left            =   2880
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin MSComctlLib.TreeView tvResult 
      Height          =   5175
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   9128
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strResult As String

Private Sub WriteString(ByVal N As Node, ByVal Level As Long)
    strResult = strResult & Space(4 * Level) & N.Text & vbCrLf
    Dim M As Node
    Set M = N.Child
    Do Until M Is Nothing
        WriteString M, Level + 1
        Set M = M.Next
    Loop
End Sub

Private Sub cmdCopy_Click()
    strResult = ""
    WriteString tvResult.Nodes(1), 0
    Clipboard.Clear
    Clipboard.SetText strResult
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    
    cmdOK.Left = Me.ScaleWidth - tvResult.Left - cmdOK.Width
    cmdOK.Top = Me.ScaleHeight - cmdOK.Height - tvResult.Top
    
    cmdCopy.Top = Me.ScaleHeight - cmdCopy.Height - tvResult.Top
    
    tvResult.Width = Me.ScaleWidth - tvResult.Left * 2
    tvResult.Height = cmdOK.Top - tvResult.Top * 2
    
End Sub

Private Sub tvResult_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub
