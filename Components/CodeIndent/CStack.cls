VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_colStack As Collection

Public Sub Clear()
    Do Until m_colStack.count <= 0
        m_colStack.Remove 1
    Loop
End Sub

Public Function Push(v As Variant)
    On Error Resume Next
    m_colStack.Add v
End Function

Public Sub Pop()
    On Error Resume Next
    m_colStack.Remove m_colStack.count
End Sub

Public Property Get count() As Long
    On Error Resume Next
    count = m_colStack.count
End Property

Public Property Get Top() As Variant
    On Error Resume Next
    Top = m_colStack.item(m_colStack.count)
End Property

Private Sub Class_Initialize()
    Set m_colStack = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colStack = Nothing
End Sub

