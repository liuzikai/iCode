VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLinkerTool 
   Caption         =   "iCode 编译工具"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   Icon            =   "frmLinkerTool.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   465
      Left            =   5190
      TabIndex        =   13
      Top             =   3840
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   465
      Left            =   3960
      TabIndex        =   12
      Top             =   3840
      Width           =   1065
   End
   Begin VB.Frame FrameAdvance 
      Caption         =   "高级"
      Height          =   2115
      Left            =   120
      TabIndex        =   5
      Top             =   1620
      Width           =   6135
      Begin VB.CommandButton Command1 
         Caption         =   "浏览..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   5070
         TabIndex        =   11
         Top             =   1620
         Width           =   885
      End
      Begin VB.TextBox txtManifest 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   510
         TabIndex        =   10
         Top             =   1620
         Width           =   4485
      End
      Begin VB.CheckBox chkManifest 
         Caption         =   "（高级）注入自定义 Manifest 文件"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1290
         Width           =   3915
      End
      Begin VB.CheckBox chkStyle 
         Caption         =   "加入系统风格文件（支持 Windows Common Control 随系统变化风格）"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   660
         Width           =   5895
      End
      Begin VB.CheckBox chkUAC 
         Caption         =   "默认申请UAC权限"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   2295
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   5190
         Top             =   1050
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Manifest文件（*.Manifest）|*.Manifest"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "建议代码中使用 InitCommonControlsEx 初始化控件"
         Height          =   180
         Left            =   510
         TabIndex        =   8
         Top             =   960
         Width           =   4140
      End
   End
   Begin VB.Frame FrameIcon 
      Caption         =   "图标"
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6105
      Begin VB.CommandButton cmdIconNext 
         Caption         =   "→"
         Height          =   255
         Left            =   5460
         TabIndex        =   17
         Top             =   1020
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdIconLast 
         Caption         =   "←"
         Height          =   255
         Left            =   5100
         TabIndex        =   16
         Top             =   1020
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox imgIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   5100
         ScaleHeight     =   45
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   43
         TabIndex        =   15
         Top             =   240
         Width           =   645
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4320
         Top             =   90
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "图标文件（*.ico;*.icon）|*.ico;*.icon"
      End
      Begin VB.CommandButton cmdOpenIcon 
         Caption         =   "浏览..."
         Height          =   345
         Left            =   3960
         TabIndex        =   3
         Top             =   585
         Width           =   885
      End
      Begin VB.CheckBox chkIcon 
         Caption         =   "替换生成文件图标（支持真彩色）"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.TextBox txtIcon 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   600
         Width           =   3705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "仅支持替换EXE文件图标（任务栏、标题栏图标不变）"
         Height          =   180
         Left            =   180
         TabIndex        =   4
         Top             =   1020
         Width           =   4230
      End
   End
   Begin VB.Label lblState 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   150
      TabIndex        =   14
      Top             =   4050
      Width           =   90
   End
End
Attribute VB_Name = "frmLinkerTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const iSection = "iCode_Linker"

Public CurProject As VBProject

Private Declare Function ExtractIcon Lib "shell32" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Type ICONINFO
        fIcon As Long
        xHotspot As Long
        yHotspot As Long
        hbmMask As Long
        hbmColor As Long
End Type

Private CurIconIndex As Long
Private IconCount As Long
Private hIcon(30) As Long

Private Sub LoadIcons(ByVal sIconFile As String)
    Do
        hIcon(IconCount + 1) = ExtractIcon(App.hInstance, sIconFile, IconCount)
        If hIcon(IconCount + 1) <> 0 Then
            IconCount = IconCount + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub UnloadIcons()
    Dim i As Long
    For i = 1 To IconCount
        DestroyIcon hIcon(i)
    Next
    IconCount = 0
End Sub

Private Sub PrintIcon(ByVal Index As Long)
    Dim Info As ICONINFO
    If GetIconInfo(hIcon(Index), Info) <> 0 Then
        Dim x As Long, y As Long
        x = imgIcon.Left + imgIcon.Width / 2
        y = imgIcon.Top + imgIcon.Height / 2
        imgIcon.Left = x - Me.ScaleX(Info.xHotspot, 3, 1)
        imgIcon.Top = y - Me.ScaleY(Info.yHotspot, 3, 1)
        imgIcon.Width = Me.ScaleX(Info.xHotspot, 3, 1) * 2
        imgIcon.Height = Me.ScaleY(Info.yHotspot, 3, 1) * 2
        imgIcon.Cls
        DrawIcon imgIcon.hdc, 0, 0, hIcon(Index)
    End If
End Sub
Private Sub chkIcon_Click()
    txtIcon.Enabled = chkIcon.Enabled
    cmdOpenIcon.Enabled = chkIcon.Enabled
    Label1.Enabled = chkIcon.Enabled
    imgIcon.Enabled = chkIcon.Enabled
End Sub

Private Sub chkManifest_Click()
    chkUAC.Enabled = chkManifest.Value - 1
    chkStyle.Enabled = chkManifest.Value - 1
    Label2.Enabled = chkManifest.Value - 1
    txtManifest.Enabled = chkManifest.Value
    Command1.Enabled = chkManifest.Value
    If chkManifest.Enabled Then
        chkUAC.Value = 0
        chkStyle.Value = 0
    End If
End Sub

Private Sub chkStyle_Click()
    chkManifest.Enabled = chkUAC.Value - 1 And chkStyle.Value - 1
    If chkStyle.Value = 0 Then chkManifest.Value = 0
End Sub

Private Sub chkUAC_Click()
    chkManifest.Enabled = chkUAC.Value - 1 And chkStyle.Value - 1
    If chkUAC.Value = 0 Then chkManifest.Value = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdIconLast_Click()
    CurIconIndex = CurIconIndex - 1
    PrintIcon CurIconIndex
    If CurIconIndex <= 1 Then cmdIconLast.Visible = False
    If IconCount > 1 Then cmdIconNext.Visible = True
End Sub

Private Sub cmdIconNext_Click()
    CurIconIndex = CurIconIndex + 1
    PrintIcon CurIconIndex
    If CurIconIndex >= IconCount Then cmdIconNext.Visible = False
    cmdIconLast.Visible = True
End Sub

Private Sub cmdOK_Click()
    If chkIcon.Value = 1 Then LinkerData.pIcon = txtIcon.Text
    If chkManifest.Value = 1 Then
        LinkerData.pManifest = txtManifest
    Else
        If chkUAC.Value = 1 And chkStyle.Value = 1 Then
            LinkerData.pManifest = LinkerPath & "UAC_Style.manifest"
        ElseIf chkUAC.Value = 1 Then
            LinkerData.pManifest = LinkerPath & "UAC.manifest"
        ElseIf chkStyle.Value = 1 Then
            LinkerData.pManifest = LinkerPath & "Style.manifest"
        End If
    End If
    
    If Not (CurProject Is Nothing) Then
        
        CurProject.WriteProperty iSection, "Setting", chkIcon.Value & "," & _
                                                      txtIcon.Text & "," & _
                                                      chkUAC.Value & "," & _
                                                      chkStyle.Value & "," & _
                                                      chkManifest.Value & "," & _
                                                      txtManifest.Text
    End If
    
    Unload Me
End Sub

Private Sub SetIconImage(ByVal sIconFile As String)
    UnloadIcons
    LoadIcons sIconFile
    CurIconIndex = 1
    PrintIcon CurIconIndex
    cmdIconNext.Visible = (IconCount > 1)
    cmdIconLast.Visible = False
End Sub

Private Sub cmdOpenIcon_Click()
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        txtIcon.Text = CommonDialog1.FileName
        SetIconImage CommonDialog1.FileName
    End If
End Sub

Private Sub Command1_Click()
    CommonDialog2.ShowOpen
    If CommonDialog2.FileName <> "" Then
        txtManifest.Text = CommonDialog2.FileName
    End If
End Sub

Private Sub Form_Load()
    
    On Error Resume Next '避免没有记录项时ReadProperty引发错误
    
    If Not (CurProject Is Nothing) Then
        
        Dim Setting As String
        Setting = CurProject.ReadProperty(iSection, "Setting")
        
        If Setting <> "" Then
        
        Dim pS() As String
        pS = Split(Setting, ",")
        
        chkIcon.Value = CLng(pS(0))
        txtIcon.Text = pS(1): SetIconImage txtIcon.Text
        chkUAC.Value = CLng(pS(2))
        chkStyle.Value = CLng(pS(3))
        chkManifest.Value = CLng(pS(4))
        txtManifest.Text = pS(5)
                                                      
        End If
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadIcons
End Sub

Private Sub Label2_Click()
    chkStyle.Value = 1 - chkStyle.Value
End Sub
