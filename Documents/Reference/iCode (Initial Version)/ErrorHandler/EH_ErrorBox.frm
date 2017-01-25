VERSION 5.00
Begin VB.Form EH_ErrorBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   1770
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   2790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   186
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label lblMain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblMain"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   3
      Top             =   420
      Width           =   630
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "确定"
      ForeColor       =   &H00FF8080&
      Height          =   180
      Left            =   1770
      TabIndex        =   2
      Top             =   1500
      Width           =   360
   End
   Begin VB.Label lblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "帮助"
      ForeColor       =   &H00FF8080&
      Height          =   180
      Left            =   2280
      TabIndex        =   1
      Top             =   1500
      Width           =   360
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Microsoft Visual Basic"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   1980
   End
End
Attribute VB_Name = "EH_ErrorBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
'事件声明:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event PressOK()
Event PressHelp(ByVal sHelpFile As String, ByVal nHelpContext As Long)
Event GotFocus()


Private Const cButtonXDis As Single = 7
Private Const cButtonsDis As Single = 9
Private Const cButtonYDis As Single = 9

Private iHelpContext As Long
Private nFormBorderWidth As Long, nFormBorderHeight As Long

Private Sub Form_Load()
    nFormBorderWidth = Me.ScaleX(Me.Width, 1, 3) - Me.ScaleWidth
    nFormBorderHeight = Me.ScaleY(Me.Height, 1, 3) - Me.ScaleHeight
    SetLayered Me.hWnd, 200
    SetAlwaysOnTop Me.hWnd, True
End Sub

Private Sub lblHelp_Click()
    RaiseEvent PressHelp(VBIns.ActiveVBProject.HelpFile, iHelpContext)
End Sub

Private Sub lblHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.ForeColor = &HC00000
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblHelp.ForeColor <> &HFF0000 Then lblHelp.ForeColor = &HFF0000
End Sub

Private Sub lblHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.ForeColor = &HFF0000
End Sub

Private Sub lblOK_Click()
    DBPrint "lblOK_Click"
    RaiseEvent PressOK
End Sub

Private Sub lblOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblOK.ForeColor = &HC00000
End Sub

Private Sub lblOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblOK.ForeColor <> &HFF0000 Then lblOK.ForeColor = &HFF0000
End Sub

Private Sub lblOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblOK.ForeColor = &HFF0000
End Sub

Public Sub ShowError(sStr As String, nHelpContext As Long)
    iHelpContext = nHelpContext
    lblMain.Caption = sStr
    Form_Resize
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If lblOK.ForeColor <> &HFF8080 Then lblOK.ForeColor = &HFF8080
    If lblHelp.ForeColor <> &HFF8080 Then lblHelp.ForeColor = &HFF8080
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Form_ReadProperties(PropBag As PropertyBag)
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub Form_Resize()
    With Me
        .Cls
        
        .Width = 2700 + nFormBorderWidth
        lblHelp.Left = .ScaleWidth - cButtonXDis - lblHelp.Width
        lblOK.Left = lblHelp.Left - cButtonsDis - lblOK.Width
        
        lblOK.Top = lblMain.Top + lblMain.Height + cButtonYDis
        lblHelp.Top = lblOK.Top
        .Height = .ScaleY(lblOK.Top + lblOK.Height + cButtonYDis + nFormBorderHeight, 3, 1)
        
    End With
End Sub

Private Sub Form_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", Me.Enabled, True)
End Sub

Private Sub SetLayered(ByVal hWnd As Long, ByVal Alpha As Byte)
    SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes hWnd, 0, Alpha, 2
End Sub

