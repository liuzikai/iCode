VERSION 5.00
Begin VB.UserControl SelectBox 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "SelectBox.ctx":0000
   Begin VB.Timer mTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3630
      Top             =   1650
   End
End
Attribute VB_Name = "SelectBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum SB_Style
    SB_sTransverse = 0
    SB_sLongitudinal = 1
End Enum

Public Enum SB_SelectMode
    SB_mClick = 0
    SB_mMouse = 1
End Enum

Private m_Items() As String
Public SelectIndex As Long


Private Const RectBorderColor = &HFF0000C8
Private Const RectFillColor = &H400000FF
Private Const RectDis = 1
Private Const StringColor = &HFF000000

Private Const LineXPC = 0.15
Private Const LineYPC = 0.15

Private Const StringSize = 16

Private g As Long, RPen As Long, RBrush As Long, LPen As Long
Private FontFam As Long, StrFormat As Long, curFont As Long, SBrush As Long


Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function lStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Const V1PerV0 = 0.01
Private Const tTotal = 400

Private V0 As Single, a As Single, t As Single

Private mStart As Single, mAim As Single, mO As Single, mRectW As Single, mRectH As Single, mAimIndex As Long

'缺省属性值:
Const m_def_Style = 0
Const m_def_SelectMode = 0
Const m_def_LineColor = &HFF5E5E5E

'属性变量:
Dim m_Style As SB_Style
Dim m_SelectMode As SB_SelectMode
Dim m_LineColor As Long

'事件声明:
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Hide()
Event Show()

Event ItemSelect(Index As Long, item As String)
Event ItemUnSelect(Index As Long, item As String)
Event ItemMouseMove(Index As Long, item As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event ItemMouseDown(Index As Long, item As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event ItemMouseUp(Index As Long, item As String, Button As Integer, Shift As Integer, X As Single, Y As Single)



Public Sub PrintSelection()
    Dim nStep As Double
    
    If SelectIndex = 0 Then Exit Sub
    
    If Me.Style = SB_sTransverse Then
        
        nStep = (UserControl.ScaleWidth - (Me.ItemsTotal - 1)) / Me.ItemsTotal
        
        DrawRect (nStep + 1) * (SelectIndex - 1) + RectDis, _
        RectDis, _
        nStep - RectDis * 2 - 1, _
        UserControl.ScaleHeight - RectDis * 2 - 1
        
    ElseIf Me.Style = SB_sLongitudinal Then
        
        nStep = (UserControl.ScaleHeight - (Me.ItemsTotal - 1)) / Me.ItemsTotal
        
        DrawRect RectDis, _
        (nStep + 1) * (SelectIndex - 1) + RectDis, _
        UserControl.ScaleWidth - RectDis * 2 - 1, _
        nStep - RectDis * 2 - 1
    End If
    
End Sub

Private Sub mStartMoving(ByVal nAim As Long)
    Dim mDis As Single
    
    Dim nStep As Double
    
    If Me.Style = SB_sTransverse Then
        nStep = (UserControl.ScaleWidth - (Me.ItemsTotal - 1)) / Me.ItemsTotal
        
        mRectW = nStep - RectDis * 2 - 1
        mRectH = UserControl.ScaleHeight - RectDis * 2 - 1
    ElseIf Me.Style = SB_sLongitudinal Then
        nStep = (UserControl.ScaleHeight - (Me.ItemsTotal - 1)) / Me.ItemsTotal
        
        mRectW = UserControl.ScaleWidth - RectDis * 2 - 1
        mRectH = nStep - RectDis * 2 - 1
    End If
    
    If SelectIndex = -1 Then
        mStart = mStart + (V0 * t + (a * t * t) / 2)
    ElseIf SelectIndex = 0 Then
        If Me.Style = SB_sTransverse Then
            mStart = -mRectW
        ElseIf Me.Style = SB_sLongitudinal Then
            mStart = -mRectH
        End If
    Else
        
        If Me.Style = SB_sTransverse Then
            mStart = (nStep + 1) * (SelectIndex - 1) + RectDis
        ElseIf Me.Style = SB_sLongitudinal Then
            mStart = (nStep + 1) * (SelectIndex - 1) + RectDis
        End If
        
    End If
    
    If Me.Style = SB_sTransverse Then
        mAim = (nStep + 1) * (nAim - 1) + RectDis
        mO = RectDis
    ElseIf Me.Style = SB_sLongitudinal Then
        mAim = (nStep + 1) * (nAim - 1) + RectDis
        mO = RectDis
    End If
    
    mDis = mAim - mStart
    V0 = 2 * mDis / (tTotal / mTimer.Interval) / (1 + V1PerV0)
    a = (V1PerV0 - 1) * V0 / (tTotal / mTimer.Interval)
    
    t = 0
    
    mAimIndex = nAim
    
    SelectIndex = -1
    
    mTimer.Enabled = True
End Sub

Public Sub Init()
    
    GDIP.GdipCreateFromHDC UserControl.hDC, g
    
    GDIP.GdipCreatePen1 RectBorderColor, 1, UnitPixel, RPen
    GDIP.GdipCreateSolidFill RectFillColor, RBrush
    
    GDIP.GdipCreatePen1 LineColor, 1, UnitPixel, LPen
    
    GdipCreateFontFamilyFromName StrPtr("宋体"), 0, FontFam
    GdipCreateStringFormat 0, 0, StrFormat
    GdipCreateSolidFill &HFF000000, SBrush
    GdipSetStringFormatAlign StrFormat, StringAlignmentNear
    GdipCreateFont FontFam, 12, FontStyle.FontStyleRegular, UnitPixel, curFont
    
End Sub

Private Sub Terminate()
    
    GdipDeleteFontFamily FontFam
    GdipDeleteStringFormat StrFormat
    GdipDeleteFont curFont
    GdipDeleteBrush SBrush
    
    
    GDIP.GdipDeletePen LPen
    
    GDIP.GdipDeleteBrush RBrush
    GDIP.GdipDeletePen RPen
    
    GDIP.GdipDeleteGraphics g
    
End Sub

Private Sub DrawRect(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
    GDIP.GdipFillRectangle g, RBrush, X, Y, Width, Height
    GDIP.GdipDrawRectangle g, RPen, X, Y, Width, Height
End Sub

Private Sub DrawText(ByVal Str As String, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
    Dim RcLayOut As RECTF
    
    RcLayOut.Left = X + 2
    RcLayOut.Top = Y + 2
    RcLayOut.Right = Width - 2 * 2
    RcLayOut.Bottom = Height - 2 * 2
    
    GDIP.GdipMeasureString g, StrPtr(Str), lStrLen(Str), curFont, RcLayOut, FontStyle.FontStyleRegular, RcLayOut, 0, 0
    
    RcLayOut.Left = X + (Width - RcLayOut.Right) / 2
    RcLayOut.Top = Y + (Height - RcLayOut.Bottom) / 2
    
    GdipDrawString g, StrPtr(Str), -1, curFont, RcLayOut, StrFormat, SBrush
End Sub

Private Sub DrawLine(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single)
    GdipDrawLine g, LPen, X1, Y1, X2, Y2
End Sub

Public Sub PrintBackground()
    If Me.ItemsTotal = 0 Then Exit Sub
    
    Dim i As Long
    Dim nStep As Single
    
    If Me.Style = SB_sTransverse Then
        
        nStep = (UserControl.ScaleWidth - (Me.ItemsTotal - 1)) / Me.ItemsTotal
        
        For i = 1 To Me.ItemsTotal
            
            DrawText Items(i), (nStep + 1) * (i - 1), 0, nStep, UserControl.ScaleHeight
            
            If i <> Me.ItemsTotal Then
                DrawLine (nStep + 1) * i - 1, UserControl.ScaleHeight * LineYPC, (nStep + 1) * i - 1, UserControl.ScaleHeight * (1 - LineYPC) - 0.2
            End If
        Next
        
    ElseIf Me.Style = SB_sLongitudinal Then
        
        nStep = (UserControl.ScaleHeight - (Me.ItemsTotal - 1)) / Me.ItemsTotal
        
        For i = 1 To Me.ItemsTotal
            
            DrawText Items(i), 0, (nStep + 1) * (i - 1), UserControl.ScaleWidth, nStep
            
            If i <> Me.ItemsTotal Then
                DrawLine UserControl.ScaleWidth * LineXPC, (nStep + 1) * i - 1, UserControl.ScaleWidth * (1 - LineXPC) - 0.2, (nStep + 1) * i - 1
            End If
        Next
        
        
    End If
End Sub

Public Sub Draw()
    Me.Cls
    
    PrintBackground
    PrintSelection
    
    UserControl.Refresh
End Sub


Public Function HitTest(ByVal X, ByVal Y) As Long
    If Me.ItemsTotal = 0 Then Exit Function
    
    If Me.Style = SB_sTransverse Then
        HitTest = (X \ (CDbl(UserControl.ScaleWidth) / Me.ItemsTotal)) + 1
    ElseIf Me.Style = SB_sLongitudinal Then
        HitTest = (Y \ (CDbl(UserControl.ScaleHeight) / Me.ItemsTotal)) + 1
    End If
    
    If HitTest > Me.ItemsTotal Then HitTest = Me.ItemsTotal
End Function

Public Property Get Items(ByVal n) As String
    If n = -1 Then Exit Property
    Items = m_Items(n)
End Property

Public Property Let Items(ByVal n, ByVal NewValue As String)
    If n = -1 Then Exit Property
    m_Items(n) = NewValue
End Property

Public Function AddItem(ByVal Caption As String) As String
    ReDim Preserve m_Items(ItemsTotal + 1)
    Items(ItemsTotal) = Caption
End Function

Public Property Get Selection() As String
    Selection = Items(SelectIndex)
End Property

Public Function SelectItem(ByVal n) As Boolean
    Dim i As Long
    
    If IsNumeric(n) And n <= Me.ItemsTotal Then
        i = n
    Else
        
        For i = 1 To Me.ItemsTotal
            If Items(i) = n Then Exit For
        Next
        
        If i > Me.ItemsTotal Then Exit Function
        
    End If
    
    SelectionChange i
    
    SelectItem = True
End Function

Public Property Get ItemsTotal() As Long
    ItemsTotal = UBound(m_Items)
End Property

Private Sub SelectionChange(ByVal i As Long)
    If i = SelectIndex Then
        
    ElseIf SelectIndex = -1 Then
        If mAimIndex = i Then
            
        Else
            RaiseEvent ItemUnSelect(mAimIndex, Items(mAimIndex))
            mStartMoving i
        End If
    Else
        If SelectIndex <> 0 Then RaiseEvent ItemUnSelect(SelectIndex, Items(SelectIndex))
        mStartMoving i
    End If
    
End Sub


Private Sub mTimer_Timer()
    
    Dim mP As Single
    
    mP = mStart + (V0 * t + (a * t * t) / 2)
    
    t = t + 1
    
    If t >= (tTotal / mTimer.Interval) Then
        mP = mAim
        SelectIndex = mAimIndex
        RaiseEvent ItemSelect(SelectIndex, Items(SelectIndex))
        mTimer.Enabled = False
    End If
    
    Me.Cls
    Me.PrintBackground
    
    If Me.Style = SB_sTransverse Then
        DrawRect mP, mO, mRectW, mRectH
    ElseIf Me.Style = SB_sLongitudinal Then
        DrawRect mO, mP, mRectW, mRectH
    End If
    
    UserControl.Refresh
End Sub

Private Sub UserControl_Initialize()
    ReDim m_Items(0)
End Sub


Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property



Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BackStyle() As Integer
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property



Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
    Dim i As Long
    
    i = HitTest(X, Y)
    
    RaiseEvent ItemMouseDown(i, Items(i), Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    Dim i As Long
    
    i = HitTest(X, Y)
    
    
    If Me.SelectMode = SB_mMouse Then
        SelectionChange i
    End If
    
    RaiseEvent ItemMouseMove(i, Items(i), Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
    Dim i As Long
    
    i = HitTest(X, Y)
    
    RaiseEvent ItemMouseUp(i, Items(i), Button, Shift, X, Y)
    
    If i = 0 Then Exit Sub
    
    If Me.SelectMode = SB_mClick Then
        SelectionChange i
    End If
    
End Sub


Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub

Public Sub Cls()
    GDIP.GdipGraphicsClear g, RGBToARGB(OLEColorToRGB(Me.BackColor))
    UserControl.Refresh
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_Style = m_def_Style
    m_SelectMode = m_def_SelectMode
    m_LineColor = m_def_LineColor
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_SelectMode = PropBag.ReadProperty("SelectMode", m_def_SelectMode)
    m_LineColor = PropBag.ReadProperty("LineColor", m_def_LineColor)
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        RaiseEvent Show
        Init
    End If
End Sub

Private Sub UserControl_Terminate()
    Terminate
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("SelectMode", m_SelectMode, m_def_SelectMode)
    Call PropBag.WriteProperty("LineColor", m_LineColor, m_def_LineColor)
End Sub

Public Property Get Style() As SB_Style
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As SB_Style)
    m_Style = New_Style
    PropertyChanged "Style"
End Property

Public Property Get SelectMode() As SB_SelectMode
    SelectMode = m_SelectMode
End Property

Public Property Let SelectMode(ByVal New_SelectMode As SB_SelectMode)
    m_SelectMode = New_SelectMode
    PropertyChanged "SelectMode"
End Property

Public Property Get LineColor() As String
    LineColor = "&H" & Hex(m_LineColor)
End Property

Public Property Let LineColor(ByVal New_LineColor As String)
    m_LineColor = New_LineColor
    PropertyChanged "LineColor"
End Property

