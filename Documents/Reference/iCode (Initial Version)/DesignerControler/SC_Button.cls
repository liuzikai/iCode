VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SC_Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum SC_ButtonType
    SC_BT_Copy
    SC_BT_Cut
    SC_BT_Paste
End Enum

Private ObjBtn As CommandBarButton
Private WithEvents BtnEventHandler As CommandBarEvents
Attribute BtnEventHandler.VB_VarHelpID = -1

Public Sub Load(ByVal ButtonType As SC_ButtonType, ByVal Bar As String, ByVal Before As Long)
    On Error Resume Next
    
    Set ObjBtn = VBIns.CommandBars(Bar).Controls.Add(msoControlButton, , , Before)
    Select Case ButtonType
    Case SC_BT_Copy
        ObjBtn.Caption = "复制 控件+代码(&C)"
    Case SC_BT_Cut
        ObjBtn.Caption = "剪切 控件+代码(&C)"
    Case SC_BT_Paste
        ObjBtn.Caption = "粘贴 控件+代码(&C)"
    End Select
    Set BtnEventHandler = VBIns.Events.CommandBarEvents(ObjBtn)
    ObjBtn.Tag = ButtonType
End Sub

Public Property Get Enabled() As Boolean
    Enabled = ObjBtn.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    ObjBtn.Enabled = NewValue
End Property

Public Property Get Visible() As Boolean
    Visible = ObjBtn.Visible
End Property

Public Property Let Visible(ByVal NewValue As Boolean)
    ObjBtn.Visible = NewValue
End Property

Public Property Get ButtonType() As SC_ButtonType
    ButtonType = ObjBtn.Tag
End Property

Public Sub Excute()
    ObjBtn.Execute
End Sub

Private Sub BtnEventHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    'iDesigner.SC_ButtonEvent CommandBarControl.Tag
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    ObjBtn.Delete
    Set ObjBtn = Nothing
    Set BtnEventHandler = Nothing
End Sub
