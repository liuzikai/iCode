VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private frmContainer As frmContainer
Private WithEvents cmdLinkerTool As CommandButton
Attribute cmdLinkerTool.VB_VarHelpID = -1
Public hContainer As Long

Public CurProject As VBProject

Public Function Msg_CompileWindowShow(ByVal hWnd As Long) As Boolean
    Msg_CompileWindowShow = (SetParent(hContainer, hWnd) <> 0)
End Function

'父窗口WM_DESTORY时按钮仍未被销毁，可回收
Public Sub Msg_CompileWindowDestory()
    Call SetParent(hContainer, frmContainer.hWnd)
End Sub

Private Sub Class_Initialize()
    LinkerPath = App.Path & "\iCode Linker\"
    Set frmContainer = New frmContainer
    frmContainer.Visible = False
    Set cmdLinkerTool = frmContainer.cmdiCodeLinker
    hContainer = frmContainer.Container.hWnd
End Sub

Private Sub Class_Terminate()
    Set cmdLinkerTool = Nothing
    Set frmContainer = Nothing
End Sub

Private Sub cmdLinkerTool_Click()
    Set frmLinkerTool.CurProject = CurProject
    If frmLinkerTool.Visible = False Then
        SetOnTop frmLinkerTool.hWnd, True
        frmLinkerTool.Show 1
    End If
End Sub

'资源替换等操作放在clsWindowsHandler中
