VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtonCollection_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Button As CommandBarButton
Private WithEvents m_Button_Event As CommandBarEvents
Attribute m_Button_Event.VB_VarHelpID = -1

Public Parent As ButtonCollection

Public Property Set Button(ByVal Value As CommandBarControl)
    Set m_Button = Value
    Set m_Button_Event = VBIns.Events.CommandBarEvents(m_Button)
End Property

Public Property Get Button() As CommandBarControl
    Set Button = m_Button
End Property

Private Sub m_Button_Event_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call Parent.CallBack(CommandBarControl)
End Sub
