VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   420
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin iCode_Setup.Button btnMain 
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
      Caption         =   "Hi"
      ChangeBackColor =   0   'False
   End
   Begin iCode_Setup.MinimizeButton MinimizeButton 
      Height          =   255
      Left            =   5340
      TabIndex        =   1
      Top             =   120
      Width           =   255
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin iCode_Setup.CloseButton CloseButton 
      Height          =   255
      Left            =   5820
      TabIndex        =   0
      Top             =   120
      Width           =   255
      _ExtentX        =   661
      _ExtentY        =   556
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "安装iCode即意味着您同意许可协议"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3180
      TabIndex        =   6
      Top             =   4260
      Width           =   2850
   End
   Begin VB.Image ccSp6 
      Height          =   195
      Left            =   300
      Picture         =   "Form1.frx":1856A
      Top             =   3840
      Width           =   195
   End
   Begin VB.Image ccPath 
      Height          =   195
      Left            =   300
      Picture         =   "Form1.frx":187B6
      Top             =   3120
      Width           =   195
   End
   Begin VB.Image imgMain 
      Height          =   1155
      Left            =   2520
      Picture         =   "Form1.frx":18A02
      Top             =   1020
      Width           =   1155
   End
   Begin VB.Image imgcc1 
      Height          =   195
      Left            =   5700
      Picture         =   "Form1.frx":1D00E
      Top             =   2100
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblSP6 
      BackStyle       =   0  'Transparent
      Caption         =   "需要部件：VB6 SP6（未找到）"
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
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Caption         =   "安装目录：VB6 安装目录（未能识别）"
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
      Height          =   675
      Left            =   660
      TabIndex        =   4
      Top             =   3105
      Width           =   5460
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "iCode 安装"
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

Private Const iVersion As String = "1.1"

Private oShadow As New aShadow

Private Path As String
Private bVB6 As Boolean, bSp6 As Boolean



Private Sub Install()

    Dim p() As String
    p = Split(LoadResData(100, 6), ",")
    
    If Dir(Path & "\iCode", vbDirectory) = "" Then MkDir Path & "\iCode"
    
    Dim i As Long, s As String
    For i = 1 To CLng(p(0))
        s = Path & "\iCode\" & LoadResData(100 + i, 6)
        If Dir(s, vbDirectory) = "" Then
            MkDir s
        End If
    Next
    
    Dim Data() As Byte
    For i = CLng(p(0)) + 1 To CLng(p(1))
    
        s = Path & "\iCode\" & LoadResData(100 + i, 6)
        Data = LoadResData(100 + i, "CUSTOM")
        Open s For Binary As #1
        Put #1, , Data
        Close #1
        
        btnMain.Caption = Int((i * 2 - 1) / (CLng(p(1)) * 2) * 100) & "%"
        DoEvents
        If LCase(Right(s, 3)) = "dll" Then
            Shell "regsvr32 /s """ & s & """"
            DoEvents
        End If
        btnMain.Caption = Int((i * 2) / (CLng(p(1)) * 2) * 100) & "%"
        DoEvents
        
    Next
    
    Dim hKey As Long
    If RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\iCode", hKey) = 0 Then
        Reg_SetValue_SZ hKey, "DisplayName", "iCode For VB6"
        Reg_SetValue_SZ hKey, "DisplayIcon", Path & "\iCode\icon.ico"
        Reg_SetValue_SZ hKey, "Publisher", "liuzikai"
        Reg_SetValue_SZ hKey, "UninstallString", Path & "\iCode\uninstall.exe"
        Reg_SetValue_SZ hKey, "URLInfoAbout", "liuzikai@163.com"
        Reg_SetValue_SZ hKey, "DisplayVersion", iVersion
        RegSetValueEx hKey, "NoModify", 0&, REG_DWORD, 1, 4
        RegSetValueEx hKey, "NoRepair", 0&, REG_DWORD, 1, 4
        RegCloseKey hKey
    Else
        MsgBox "写入注册表失败！需要卸载时请手动运行卸载软件", vbExclamation, "iCode Setup"
    End If
    
    
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

Private Sub SelectPath_VB6()
    CommonDialog1.Filter = "VB主程序(VB6.exe)|VB6.exe"
    CommonDialog1.FileName = Environ("ProgramFiles") & "\VB6.exe"
    CommonDialog1.ShowOpen
    Confirm_VB6 Left(CommonDialog1.FileName, InStrRev(CommonDialog1.FileName, "\"))
End Sub

Private Sub SelectPath_SP6()
    CommonDialog1.Filter = "SP6 组件(MSCOMCTL.OCX)|MSCOMCTL.OCX"
    CommonDialog1.FileName = Environ("WINDIR") & "\MSCOMCTL.OCX"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then Confirm_SP6
End Sub

Private Sub Confirm_VB6(ByVal s As String)
    If s <> "" Then
        Path = s
        lblPath.Caption = "安装目录：VB6 安装目录" & vbCrLf & "（" & Path & "）"
        Set ccPath.Picture = imgcc1.Picture
        bVB6 = True
    End If
End Sub

Private Sub Confirm_SP6()
    lblSP6.Caption = "需要部件：VB6 SP6（已识别）"
    Set ccSp6.Picture = imgcc1.Picture
    bSp6 = True
End Sub

Private Sub btnMain_Click()
    Select Case btnMain.Caption
    Case "退出"
        Unload Me
    Case "手动" & vbCrLf & "选择"
        If Not bVB6 Then
            SelectPath_VB6
        End If
        If Not bSp6 Then
            SelectPath_SP6
        End If
        If bVB6 And bSp6 Then
            btnMain.Caption = "安装"
        End If
    Case "安装"
        Install
    End Select
End Sub

Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Confirm_VB6 Reg_Read_SZ(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic\", "ProductDir")

    
    If Dir(Environ("WINDIR") & "\system32\MSCOMCTL.OCX") <> "" Or Dir(Environ("WINDIR") & "\sysWOW64\MSCOMCTL.OCX") <> "" Then
        Confirm_SP6
    End If
    
    If bVB6 And bSp6 Then
        btnMain.Caption = "安装"
    Else
        btnMain.Caption = "手动" & vbCrLf & "选择"
    End If
    
    LoadShadow
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CloseButton.Reset
    MinimizeButton.Reset
    btnMain.Reset
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oShadow = Nothing
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
    End If
End Sub

Private Sub lblPath_DblClick()
    SelectPath_VB6
End Sub

Private Sub lblSP6_Click()
    SelectPath_SP6
End Sub
