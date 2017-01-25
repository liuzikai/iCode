VERSION 5.00
Begin VB.Form frmAddProperty 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   327
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txt_VarName 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   840
      TabIndex        =   19
      Top             =   3780
      Width           =   3195
   End
   Begin VB.CheckBox chk_AutoVar 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "生成存储变量"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   180
      TabIndex        =   18
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txt_Params 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   840
      TabIndex        =   14
      Top             =   2940
      Width           =   3195
   End
   Begin VB.TextBox txt_Type 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   840
      TabIndex        =   13
      Top             =   2580
      Width           =   3195
   End
   Begin VB.TextBox txt_Name 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   840
      TabIndex        =   12
      Top             =   2220
      Width           =   3195
   End
   Begin VB.CheckBox chk_Set 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Set"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   180
      TabIndex        =   11
      Top             =   1320
      Width           =   675
   End
   Begin VB.CheckBox Chk_Let 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Let"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   180
      TabIndex        =   9
      Top             =   960
      Value           =   1  'Checked
      Width           =   675
   End
   Begin VB.CheckBox chk_Get 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Get"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Top             =   540
      Value           =   1  'Checked
      Width           =   675
   End
   Begin iCode_IDEEnhancer.SelectBox Scope_Get 
      Height          =   315
      Left            =   900
      TabIndex        =   6
      Top             =   540
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      BackColor       =   4210752
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Col_Next 
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1620
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   4260
      Width           =   1155
      Begin VB.Label lbl_Next 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "确认并继续"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.PictureBox Con_Minimize 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   3
      Top             =   120
      Width           =   255
      Begin VB.Line Line_Minimize 
         BorderColor     =   &H80000010&
         BorderWidth     =   3
         X1              =   4
         X2              =   12
         Y1              =   8
         Y2              =   8
      End
   End
   Begin VB.PictureBox Con_Close 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3810
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   2
      Top             =   120
      Width           =   255
      Begin VB.Line Line_Close 
         BorderColor     =   &H80000010&
         BorderWidth     =   3
         Index           =   1
         X1              =   4
         X2              =   12
         Y1              =   12
         Y2              =   4
      End
      Begin VB.Line Line_Close 
         BorderColor     =   &H80000010&
         BorderWidth     =   3
         Index           =   0
         X1              =   4
         X2              =   12
         Y1              =   4
         Y2              =   12
      End
   End
   Begin VB.PictureBox Col_OK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   2880
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   4260
      Width           =   1155
      Begin VB.Label lbl_OK 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "确定"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   420
         TabIndex        =   1
         Top             =   120
         Width           =   360
      End
   End
   Begin iCode_IDEEnhancer.SelectBox Scope_Let 
      Height          =   315
      Left            =   900
      TabIndex        =   7
      Top             =   960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      BackColor       =   4210752
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin iCode_IDEEnhancer.SelectBox Scope_Set 
      Height          =   315
      Left            =   900
      TabIndex        =   10
      Top             =   1320
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      BackColor       =   4210752
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin iCode_IDEEnhancer.SelectBox DeliverWay 
      Height          =   315
      Left            =   1980
      TabIndex        =   22
      Top             =   1680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      BackColor       =   4210752
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image img_Icon 
      Height          =   210
      Left            =   300
      Picture         =   "frmAddProperty.frx":0000
      Top             =   4440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "添加属性"
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
      TabIndex        =   21
      Top             =   120
      Width           =   3075
   End
   Begin VB.Label lbl_VarName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "变量："
      ForeColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   180
      TabIndex        =   20
      Top             =   3825
      Width           =   540
   End
   Begin VB.Label lbl_Params 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数："
      ForeColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   180
      TabIndex        =   17
      Top             =   2985
      Width           =   540
   End
   Begin VB.Label lbl_Type 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "类型："
      ForeColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   180
      TabIndex        =   16
      Top             =   2625
      Width           =   540
   End
   Begin VB.Label lbl_Name 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称："
      ForeColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   180
      TabIndex        =   15
      Top             =   2265
      Width           =   540
   End
End
Attribute VB_Name = "frmAddProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Const Color_Frame_Normal = &H80000010
Private Const Color_Frame_Selected = &H80000015

Private Const Color_Button_Normal = &H80000015
Private Const Color_Text_Normal = &H8000000F
Private Const Color_Button_Selected = &H80000010
Private Const Color_Text_Selected = vbBlack
Private Const Color_Button_Clicked = &H6C6C6C

Private Frame_MouseIn As Boolean
Private ButtonOK_MouseIn As Boolean
Private ButtonCopy_MouseIn As Boolean

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)

Private Sub SetTransparent(ByVal hwnd As Long, ByVal Alpha As Byte)
    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, 0, Alpha, 2
End Sub

Public Property Let iCaption(ByVal value As String)
    Label1.Caption = value
End Property

Private Sub Clear()
    chk_Get.value = 1
    chk_Get_Click
    Chk_Let.value = 1
    Chk_Let_Click
    chk_Set.value = 0
    chk_Set_Click
    txt_Name = ""
    txt_Type = ""
    txt_Params = ""
    chk_AutoVar.value = 0
    chk_AutoVar_Click
    txt_VarName = ""
    Scope_Get.SelectItem 1
    Scope_Let.SelectItem 1
    Scope_Set.SelectItem 1
    DeliverWay.SelectItem 1
End Sub

Private Sub chk_AutoVar_Click()
    txt_VarName.Enabled = chk_AutoVar.value
End Sub

Private Sub chk_Get_Click()
    Scope_Get.Enabled = chk_Get.value
End Sub

Private Sub Chk_Let_Click()
    Scope_Let.Enabled = Chk_Let.value
End Sub

Private Sub chk_Set_Click()
    Scope_Set.Enabled = chk_Set.value
End Sub

Private Sub Con_Close_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Line_Close(0).BorderColor = Color_Frame_Selected
    Line_Close(1).BorderColor = Color_Frame_Selected
    
    Line_Minimize.BorderColor = Color_Frame_Normal
    
    Frame_MouseIn = True
End Sub

Private Sub Con_Close_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    iHide
End Sub

Private Sub Con_Minimize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Line_Minimize.BorderColor = Color_Frame_Selected
    
    Line_Close(0).BorderColor = Color_Frame_Normal
    Line_Close(1).BorderColor = Color_Frame_Normal
    
    Frame_MouseIn = True
End Sub

Private Sub Con_Minimize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.WindowState = 1
    Line_Minimize.BorderColor = Color_Frame_Normal
End Sub

Private Sub Drop(ByVal Button As Integer)
    If Button = 1 Then
        ReleaseCapture
        Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
    End If
End Sub

Private Sub Form_Load()
    SetTransparent Me.hwnd, 245
    
    Me.Left = Screen.Width * (1 - 0.618) - Me.Width / 2
    Me.Top = Screen.Height * (1 - 0.618) - Me.Height / 2
    
    Me.Show
    
    Scope_Get.AddItem "Private"
    Scope_Get.AddItem "Public"
    Scope_Get.AddItem "Friend"
    
    Scope_Let.AddItem "Private"
    Scope_Let.AddItem "Public"
    Scope_Let.AddItem "Friend"
    
    Scope_Set.AddItem "Private"
    Scope_Set.AddItem "Public"
    Scope_Set.AddItem "Friend"
    
    DeliverWay.AddItem "Byval"
    DeliverWay.AddItem "Byref"
    
    Clear
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Drop Button
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Frame_MouseIn Then
        Line_Close(0).BorderColor = Color_Frame_Normal
        Line_Close(1).BorderColor = Color_Frame_Normal
        Line_Minimize.BorderColor = Color_Frame_Normal
        Frame_MouseIn = False
    End If
    
    If ButtonOK_MouseIn Then
        Col_OK.BackColor = Color_Button_Normal
        lbl_OK.ForeColor = Color_Text_Normal
        ButtonOK_MouseIn = False
    End If
    If ButtonCopy_MouseIn Then
        Col_Next.BackColor = Color_Button_Normal
        lbl_Next.ForeColor = Color_Text_Normal
        ButtonCopy_MouseIn = False
    End If
End Sub

Private Sub iHide()
    Call Form_MouseMove(0, 0, 0, 0)
    Me.Hide
End Sub

Private Sub Col_Next_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Col_Next.BackColor = Color_Button_Clicked
End Sub

Private Sub Col_Next_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not ButtonCopy_MouseIn Then
        Col_Next.BackColor = Color_Button_Selected
        lbl_Next.ForeColor = Color_Text_Selected
        ButtonCopy_MouseIn = True
    End If
End Sub

Private Sub Col_Next_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCode
    Clear
End Sub

Private Sub Col_OK_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Col_OK.BackColor = Color_Button_Clicked
End Sub

Private Sub Col_OK_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not ButtonOK_MouseIn Then
        Col_OK.BackColor = Color_Button_Selected
        lbl_OK.ForeColor = Color_Text_Selected
        ButtonOK_MouseIn = True
    End If
End Sub

Private Sub SetCode()
    Dim S As String
    
    If chk_AutoVar.value = 1 Then S = "Private " & txt_VarName & " As " & txt_Type & vbCrLf
    
    If chk_Get.value = 1 Then
        S = S & vbCrLf & Scope_Get.Selection & " Property Get " & txt_Name & "(" & txt_Params & ") As " & txt_Type & vbCrLf
        If chk_AutoVar.value = 1 Then S = S & "    " & txt_Name & " = " & txt_VarName
        S = S & vbCrLf & "End Property" & vbCrLf
    End If
    
    If Chk_Let.value = 1 Then
        S = S & vbCrLf & Scope_Let.Selection & " Property Let " & txt_Name & "("
        If txt_Params <> "" Then S = S & txt_Params & ", "
        S = S & DeliverWay.Selection & " Value As " & txt_Type & ")" & vbCrLf
        If chk_AutoVar.value = 1 Then S = S & "    " & txt_VarName & " = Value"
        S = S & vbCrLf & "End Property" & vbCrLf
    End If
    
    If chk_Set.value = 1 Then
        S = S & vbCrLf & Scope_Set.Selection & " Property Set " & txt_Name & "("
        If txt_Params <> "" Then S = S & txt_Params & ", "
        S = S & DeliverWay.Selection & " Value As " & txt_Type & ")" & vbCrLf
        If chk_AutoVar.value = 1 Then S = S & "    Set " & txt_VarName & " = Value"
        S = S & vbCrLf & "End Property" & vbCrLf
    End If
    
    CodeOpe.AddCodeToDeclaration S
    'MsgBox S
End Sub

Private Sub Col_OK_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCode
    iHide
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Drop Button
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Form_MouseMove(Button, Shift, x, y)
End Sub

Private Sub lbl_Next_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Col_Next_MouseDown(Button, Shift, x, y)
End Sub

Private Sub lbl_Next_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Col_Next_MouseMove(Button, Shift, x, y)
End Sub

Private Sub lbl_Next_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Col_Next_MouseUp(Button, Shift, x, y)
End Sub

Private Sub lbl_OK_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Col_OK_MouseDown(Button, Shift, x, y)
End Sub

Private Sub lbl_OK_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Col_OK_MouseMove(Button, Shift, x, y)
End Sub

Private Sub lbl_OK_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Col_OK_MouseUp(Button, Shift, x, y)
End Sub

