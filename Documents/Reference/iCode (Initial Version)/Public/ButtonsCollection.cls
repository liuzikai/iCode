VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtonsCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'此类销毁时会自动清除已添加的按钮

Event ButtonClick(ByVal CommandBarControl As CommandBarControl)

Private m_TargetBar As CommandBar
Private m_TargetControl As CommandBarControl

Private m_Items() As BC_Button


Public Property Get Count() As Long
    Count = UBound(m_Items)
End Property

Private Function FindItemByCaption(ByVal Caption As String) As Long
    Dim i As Long
    For i = 1 To Me.Count
        If m_Items(i).Button.Caption = Caption Then Exit For
    Next
    If i > Me.Count Then i = 0
    FindItemByCaption = i
End Function

Friend Property Get Items(ByVal n) As BC_Button
    If Not (IsNumeric(n) And n <= Me.Count) Then n = FindItemByCaption(n)
    If n <> 0 Then Set Items = m_Items(n)
End Property

Friend Property Set Items(ByVal n, ByVal Value As BC_Button)
    If Not (IsNumeric(n) And n <= Me.Count) Then n = FindItemByCaption(n)
    If n <> 0 Then Set m_Items(n) = Value
End Property

Public Function Add(ByVal ButtonType As MsoControlType, ByVal Caption As String, _
    Optional Style As MsoButtonStyle = msoButtonCaption, Optional ByVal BeginGroup As Boolean = False, _
    Optional ByVal Tag, Optional ByRef Index As Long) As CommandBarControl
    ReDim Preserve m_Items(Me.Count + 1)
    Set m_Items(Me.Count) = New BC_Button
    Set m_Items(Me.Count).Parent = Me
    'If m_TargetControl Is Nothing Then
    m_Items(Me.Count).Init Me.Target, ButtonType, Caption, Style, BeginGroup, Tag
    'Else
    'm_Items(Me.Count).Init m_TargetControl, ButtonType, Caption, Style, BeginGroup, Tag
    'End If
    Set Add = m_Items(Me.Count).Button
    Index = Me.Count
End Function

Friend Sub CallButtonEvent(ByVal CommandBarControl As Object)
    RaiseEvent ButtonClick(CommandBarControl)
End Sub

Public Function AddTargetBar(ByVal Name As String, Optional Visable As Boolean = False, Optional ByVal Protection As MsoBarProtection = 0) As CommandBar
    Set m_TargetBar = VBIns.CommandBars.Add(Name)
    m_TargetBar.Visible = Visable
    m_TargetBar.Protection = Protection
    Set AddTargetBar = m_TargetBar
End Function

Public Property Get Target() As Object
    If m_TargetControl Is Nothing Then
        Set Target = m_TargetBar
    Else
        Set Target = m_TargetControl
    End If
End Property

Public Property Set Target(BarOrControl As Object)
    If TypeName(BarOrControl) = "CommandBar" Then
        Set m_TargetBar = BarOrControl
        Set m_TargetControl = Nothing '避免目标仍保留在TargetControl上
    ElseIf TypeName(BarOrControl) = "CommandBarControl" Then
        Set m_TargetControl = BarOrControl
    End If
End Property

Public Sub AddTargetPopup(ByVal Caption As String, Optional ByVal BeginGroup As Boolean = False)
    Set m_TargetControl = Me.Add(msoControlPopup, Caption, , BeginGroup)
End Sub

Public Sub SetTargetBackToBar()
    Set m_TargetControl = Nothing
End Sub

Private Sub Class_Initialize()
    ReDim m_Items(0)
End Sub

Private Sub Class_Terminate()
    If Me.Count > 0 Then
        Dim i As Long
        For i = 1 To Me.Count
            Me.Items(i).Button.Delete
        Next
    End If
End Sub
