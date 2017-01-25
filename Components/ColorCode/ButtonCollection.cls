VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtonCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Click(ByVal Button As CommandBarControl)

Private m_Item(100) As ButtonCollection_Item
Private m_Count As Long

Public Property Get Count() As Long
    Count = m_Count
End Property

Private Property Let Count(ByVal Value As Long)
    m_Count = Value
End Property

Public Property Get Item(ByVal Index As Long) As CommandBarControl
    Set Item = m_Item(Index).Button
End Property

Public Sub Clear()
    Dim i As Long
    For i = 1 To m_Count
        Set m_Item(i) = Nothing
    Next
    m_Count = 0
End Sub

Public Function Push(ByVal Button As CommandBarControl) As Long
    Count = Count + 1
    Set m_Item(Count) = New ButtonCollection_Item
    Set m_Item(Count).Parent = Me
    Set m_Item(Count).Button = Button
    Push = m_Count
End Function

Public Sub CallBack(ByVal Button As CommandBarButton)
    RaiseEvent Click(Button)
End Sub
