VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDesignerControler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents EProject As VBProjectsEvents
Attribute EProject.VB_VarHelpID = -1
Private WithEvents ECom As VBComponentsEvents
Attribute ECom.VB_VarHelpID = -1
Private WithEvents EControls As VBControlsEvents
Attribute EControls.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set EProject = VBIns.Events.VBProjectsEvents
End Sub

Private Sub Class_Terminate()
    Set EProject = Nothing
    Set ECom = Nothing
    Set EControls = Nothing
End Sub

Private Sub ECom_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
    If VBComponent.type = vbext_ct_ActiveXDesigner Or _
        VBComponent.type = vbext_ct_DocObject Or _
        VBComponent.type = vbext_ct_MSForm Or _
        VBComponent.type = vbext_ct_PropPage Or _
        VBComponent.type = vbext_ct_RelatedDocument Or _
        VBComponent.type = vbext_ct_UserControl Or _
        VBComponent.type = vbext_ct_VBForm Or _
        VBComponent.type = vbext_ct_VBMDIForm Then
        
        Set EControls = VBIns.Events.VBControlsEvents(VBIns.ActiveVBProject, VBComponent.Designer)
        
    End If
End Sub

Private Sub ECom_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
    If MsgBox("是否更新代码 (" & """" & OldName & """" & " -> " & """" & VBComponent.Name & """" & ") ？", vbInformation Or vbYesNo, "iCode") = vbYes Then
        CodeOpe.ReplaceAll OldName, VBComponent.Name
    End If
End Sub

Private Sub EControls_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
    If VBControl.Properties("Index").Value = -1 Then
        If MsgBox("是否更新代码 (" & """" & OldName & """" & " -> " & """" & VBControl.Properties("Name") & """" & ") ？", vbInformation Or vbYesNo, "iCode") = vbYes Then
            CodeOpe.ReplaceAll OldName, VBControl.Properties("Name")
        End If
    Else
        If MsgBox("是否更新代码 (" & """" & OldName & "(" & OldIndex & ")" & """" & " -> " & """" & VBControl.Properties("Name") & "(" & VBControl.Properties("Index").Value & ")" & """" & ") ？", vbInformation Or vbYesNo, "iCode") = vbYes Then
            CodeOpe.ReplaceAll OldName & "(" & OldIndex & ")", VBControl.Properties("Name") & "(" & VBControl.Properties("Index").Value & ")"
        End If
    End If
End Sub

Private Sub EProject_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    Set ECom = VBIns.Events.VBComponentsEvents(VBProject)
End Sub

Private Sub EProject_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    Set ECom = VBIns.Events.VBComponentsEvents(VBProject)
End Sub

Private Sub EProject_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
    If MsgBox("是否更新代码 (" & """" & OldName & """" & " -> " & """" & VBProject.Name & """" & ") ？", vbInformation Or vbYesNo, "iCode") = vbYes Then
        CodeOpe.ReplaceAll OldName, VBProject.Name
    End If
End Sub
