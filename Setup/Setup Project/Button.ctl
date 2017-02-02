VERSION 5.00
Begin VB.UserControl Button 
   BackColor       =   &H00404040&
   CanGetFocus     =   0   'False
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Button"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   300
      Width           =   735
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Const UI_R As Long = 4

Private Const Color_Button_Normal = &H80000015
Private Const Color_Button_Selected = &H80000010
Private Const Color_Button_Clicked = &H6C6C6C

Private Const Color_Text_Normal = &H8000000F
Private Const Color_Text_Selected = vbBlack

Private bMouseIn As Boolean
'缺省属性值:
Const m_def_ChangeBackColor = True
'属性变量:
Dim m_ChangeBackColor As Boolean



Public Sub Reset()
    If bMouseIn Then
        If m_ChangeBackColor Then UserControl.BackColor = Color_Button_Normal
        lblCaption.ForeColor = Color_Text_Normal
        bMouseIn = False
    End If
End Sub

Private Sub lblCaption_Click()
    Call UserControl_Click
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, 0, 0)
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, 0, 0)
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, 0, 0)
End Sub

'
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    If m_ChangeBackColor Then PropertyChanged "Enabled"
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If m_ChangeBackColor Then
        UserControl.BackColor = Color_Button_Clicked
    Else
        lblCaption.ForeColor = Color_Text_Selected
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Not bMouseIn Then
        If m_ChangeBackColor Then
            UserControl.BackColor = Color_Button_Selected
            lblCaption.ForeColor = Color_Text_Selected
        Else
            lblCaption.ForeColor = Color_Button_Clicked
        End If
        bMouseIn = True
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If m_ChangeBackColor Then
        UserControl.BackColor = Color_Button_Selected
    Else
        lblCaption.ForeColor = Color_Button_Clicked
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.FontBold = PropBag.ReadProperty("FontBold", 0)
    lblCaption.FontSize = PropBag.ReadProperty("FontSize", 0)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "Button")
    m_ChangeBackColor = PropBag.ReadProperty("ChangeBackColor", m_def_ChangeBackColor)
End Sub

Private Sub UserControl_Resize()
    lblCaption.Left = (UserControl.ScaleWidth - lblCaption.Width) / 2
    lblCaption.Top = (UserControl.ScaleHeight - lblCaption.Height) / 2
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", lblCaption.FontBold, 0)
    Call PropBag.WriteProperty("FontSize", lblCaption.FontSize, 0)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Button")
    Call PropBag.WriteProperty("ChangeBackColor", m_ChangeBackColor, m_def_ChangeBackColor)
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "返回/设置粗体字样式。"
    FontBold = lblCaption.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    lblCaption.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "指定给定层的每一行出现的字体大小(以磅为单位)。"
    FontSize = lblCaption.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    lblCaption.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "返回/设置对象的标题栏中或图标下面的文本。"
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
    DoEvents
    UserControl_Resize
End Property
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get ChangeBackColor() As Boolean
    ChangeBackColor = m_ChangeBackColor
End Property

Public Property Let ChangeBackColor(ByVal New_ChangeBackColor As Boolean)
    m_ChangeBackColor = New_ChangeBackColor
    PropertyChanged "ChangeBackColor"
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_ChangeBackColor = m_def_ChangeBackColor
End Sub

