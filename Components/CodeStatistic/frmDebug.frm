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
      Top             =   90
      Width           =   3030
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Sub Form_Load()
    Me.Show
    SetAlwaysOnTop Me.hwnd, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    VBIns.Quit
    Cancel = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtMain.Width = Me.ScaleWidth - txtMain.Left * 2
    txtMain.Height = Me.ScaleHeight - txtMain.Top * 2
End Sub

Private Sub txtMain_Change()
    Clipboard.SetText txtMain
End Sub

Private Sub txtMain_DblClick()
    txtMain.Text = ""
End Sub

Private Sub txtMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        #If HaveOperator = True Then
            frmOperator.Show
        #End If
    End If
End Sub

Private Sub SetAlwaysOnTop(nHwnd As Long, bOnTop As Boolean)
    Dim rRect As RECT
    
    GetWindowRect nHwnd, rRect
    
    If bOnTop = True Then
        SetParent nHwnd, 0
        
        SetWindowPos nHwnd, HWND_TOPMOST, rRect.Left, rRect.Top, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Else
        Dim h
        
        h = FindWindow(vbNullString, "Program Manager")
        h = FindWindowEx(h, ByVal 0, "SHELLDLL_DefView", vbNullString)
        
        SetParent nHwnd, h
        
        SetWindowPos nHwnd, HWND_TOPMOST, rRect.Left, rRect.Top, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    End If
End Sub

