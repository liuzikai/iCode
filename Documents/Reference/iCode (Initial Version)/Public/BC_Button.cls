VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BC_Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Parent As ButtonsCollection
Private m_Button As CommandBarControl
Private WithEvents ButtonHandler As CommandBarEvents
Attribute ButtonHandler.VB_VarHelpID = -1

Friend Property Get Parent() As ButtonsCollection
    Set Parent = m_Parent
End Property

Friend Property Set Parent(ByVal Value As ButtonsCollection)
    Set m_Parent = Value
End Property

Friend Property Get Button() As CommandBarControl
    Set Button = m_Button
End Property

Friend Property Set Button(ByVal Value As CommandBarControl)
    Set m_Button = Value
End Property

Public Sub Init(ByVal CommandBar As Object, ByVal ButtonType As MsoControlType, ByVal Caption As String, Optional Style As MsoButtonStyle = msoButtonCaption, Optional ByVal BeginGroup As Boolean = False, Optional ByVal Tag)
    Set Button = CommandBar.Controls.Add(ButtonType)
    Button.Caption = Caption
    If ButtonType = msoControlButton Then Button.Style = Style
    Button.BeginGroup = BeginGroup
    If Not IsMissing(Tag) Then Button.Tag = Tag
    Set ButtonHandler = VBIns.Events.CommandBarEvents(Button)
End Sub

Private Sub ButtonHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Parent.CallButtonEvent CommandBarControl
End Sub
