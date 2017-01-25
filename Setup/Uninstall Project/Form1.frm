VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "iCode"
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   StartUpPosition =   3  '窗口缺省
   Begin iCode_Uninstall.Button btnMain 
      Height          =   735
      Left            =   2730
      TabIndex        =   2
      Top             =   1230
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontSize        =   10.5
      Caption         =   "卸载"
      ChangeBackColor =   0   'False
   End
   Begin iCode_Uninstall.MinimizeButton MinimizeButton 
      Height          =   255
      Left            =   5340
      TabIndex        =   1
      Top             =   120
      Width           =   255
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin iCode_Uninstall.CloseButton CloseButton 
      Height          =   255
      Left            =   5820
      TabIndex        =   0
      Top             =   120
      Width           =   255
      _ExtentX        =   661
      _ExtentY        =   556
   End
   Begin VB.Image imgcc 
      Height          =   195
      Index           =   1
      Left            =   4920
      Picture         =   "Form1.frx":1856A
      Top             =   2280
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgcc 
      Height          =   195
      Index           =   0
      Left            =   4560
      Picture         =   "Form1.frx":187B6
      Top             =   2280
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblDeleteSetting 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "删除设置文件"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   660
      TabIndex        =   6
      Top             =   2640
      Width           =   1080
   End
   Begin VB.Image ccDeleteSetting 
      Height          =   195
      Left            =   300
      Picture         =   "Form1.frx":18A02
      Tag             =   "0"
      Top             =   2655
      Width           =   195
   End
   Begin VB.Image ccSp6 
      Height          =   195
      Left            =   300
      Picture         =   "Form1.frx":18C4E
      Top             =   3840
      Width           =   195
   End
   Begin VB.Image ccPath 
      Height          =   195
      Left            =   300
      Picture         =   "Form1.frx":18E9A
      Top             =   3360
      Width           =   195
   End
   Begin VB.Image imgMain 
      Height          =   1155
      Left            =   2520
      Picture         =   "Form1.frx":190E6
      Top             =   1020
      Width           =   1155
   End
   Begin VB.Label lblContact 
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎与我联系：liuzikai@163.com   ：)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   660
      TabIndex        =   5
      Top             =   3810
      Width           =   4875
   End
   Begin VB.Label lblThanks 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "感谢您的使用!"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   660
      TabIndex        =   4
      Top             =   3345
      Width           =   1140
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "iCode 卸载"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   4875
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oShadow As New aShadow

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Sub Uninstall()

    Dim p() As String
    p = Split(LoadResData(100, 6), ",")
    
    
    Dim i As Long, k As Long, f As String, s As String
    
    s = "@echo off" & vbCrLf & _
        "sleep 700" & vbCrLf
    
    If ccDeleteSetting.Tag = "1" Then
        f = Environ("AppData") & "\iCode\Settings.ini"
        If Dir(f) <> "" Then
            If DeleteFile(f) = False Then
                s = s & "del """ & f & """" & vbCrLf
            End If
        End If
    End If
    
    For i = CLng(p(1)) To CLng(p(0)) + 1 Step -1
        
        f = App.Path & "\" & LoadResData(100 + i, 6)
        If Dir(f) <> "" Then
            If LCase(Right(f, 3)) = "dll" Then
                Shell "regsvr32 /u /s """ & f & """"
            End If
            DoEvents
            If DeleteFile(f) = False Then
                s = s & "del """ & f & """" & vbCrLf
            End If
        End If
        
        k = k + 1
        btnMain.Caption = Int(k / CLng(p(1)) * 100) & "%"
        DoEvents
        
    Next
    
    For i = CLng(p(0)) To 1 Step -1
        
        f = App.Path & "\" & LoadResData(100 + i, 6)
        If Dir(f, vbDirectory) <> "" Then
            RmDir f
        End If
        
        k = k + 1
        btnMain.Caption = Int(k / CLng(p(1)) * 100) & "%"
        DoEvents
    Next
    
    RegDeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\iCode"

    s = s & "sleep 300" & vbCrLf & _
            "rd """ & App.Path & """" & vbCrLf & _
            "del """ & Environ("Temp") & "\iCode_DeleteFile.bat"
    
    Open Environ("Temp") & "\iCode_DeleteFile.bat" For Output As #1
    Print #1, s
    Close #1
    
    DoEvents
    
    btnMain.Caption = "完成"
    
End Sub

Private Sub LoadShadow()
    With oShadow
        If .Shadow(Me) Then
            .Depth = 6 '阴影宽度
            .Color = RGB(0, 0, 0) '阴影颜色
            .Transparency = 36 '阴影色深
        End If
    End With
End Sub

Private Sub btnMain_Click()
    Select Case btnMain.Caption
    Case "卸载"
        Uninstall
    Case "完成"
        Unload Me
    End Select
End Sub

Private Sub ccDeleteSetting_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ccDeleteSetting.Tag = 1 - ccDeleteSetting.Tag
    Set ccDeleteSetting.Picture = imgcc(ccDeleteSetting.Tag).Picture
End Sub

Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    If Right(App.Path, 5) <> "iCode" Then
        MsgBox "检测到当前目录不是默认安装目录，可能导致其他文件被误删除", vbCritical, "警告"
    End If
    
    LoadShadow
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CloseButton.Reset
    MinimizeButton.Reset
    btnMain.Reset
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If btnMain.Caption = "完成" Then Shell "cmd /c """ & Environ("Temp") & "\iCode_DeleteFile.bat"""
    Set oShadow = Nothing
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
    End If
End Sub
