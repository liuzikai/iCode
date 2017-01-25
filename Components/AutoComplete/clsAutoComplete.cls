VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private udAPI As udAPI
Private wndAPI As Window

Private iCodeMenu As CommandBarPopup
Private WithEvents iCodeMenuEvent As CommandBarEvents
Attribute iCodeMenuEvent.VB_VarHelpID = -1

Private btnAPI As CommandBarButton
Private WithEvents btnAPIEvent As CommandBarEvents
Attribute btnAPIEvent.VB_VarHelpID = -1

Public Sub Initialize(ByVal VBIDE As VBE, ByVal DBForm As Object, ByVal AddInInst As Object, ByVal iMenu As Object, _
                      ByVal guid As String)
    
    Set VBIns = VBIDE
    Set AddInIns = AddInInst
    Let hVBIDE = VBIns.MainWindow.hWnd
    Set CodeOpe = New CodeOperator
    Set iCodeMenu = iMenu
    Set iCodeMenuEvent = VBIns.Events.CommandBarEvents(iCodeMenu)
    
    Set DebugForm = DBForm
    
    Set wndAPI = VBIns.Windows.CreateToolWindow(AddInIns, "iCode_AutoComplete.udAPI", "iCode - API", guid, udAPI)
    'wndAPI.Visible = True
    
    Set btnAPI = iCodeMenu.Controls.Add(msoControlButton)
    btnAPI.Visible = True
    btnAPI.BeginGroup = True
    Clipboard.SetData LoadResPicture(101, 0)
    btnAPI.PasteFace
    Set btnAPIEvent = VBIns.Events.CommandBarEvents(btnAPI)
    
End Sub

Private Sub btnAPIEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    wndAPI.Visible = Not wndAPI.Visible
End Sub

Private Sub Class_Terminate()
    Set udAPI = Nothing
    Set wndAPI = Nothing
End Sub


Private Sub iCodeMenuEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If wndAPI.Visible Then
        btnAPI.Caption = "Òþ²ØAPI´°¿Ú"
    Else
        btnAPI.Caption = "ÏÔÊ¾API´°¿Ú"
    End If
End Sub
