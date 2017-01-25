VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11715
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   16230
   _ExtentX        =   28628
   _ExtentY        =   20664
   _Version        =   393216
   Description     =   "Add-In Project Template"
   DisplayName     =   "My Add-In"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private VBIns As VBE
Private E2 As Events2
Private WithEvents BuildEvent As VBBuildEvents
Attribute BuildEvent.VB_VarHelpID = -1
Private WithEvents FCE As FileControlEvents
Attribute FCE.VB_VarHelpID = -1
Private WithEvents VCE As VBComponentsEvents
Attribute VCE.VB_VarHelpID = -1
Private WithEvents VBCE As VBControlsEvents
Attribute VBCE.VB_VarHelpID = -1

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    Set VBIns = Application
    Set E2 = VBIns.Events
    Set BuildEvent = E2.VBBuildEvents
    Set FCE = E2.FileControlEvents(Nothing)
    Set VCE = E2.VBComponentsEvents(Nothing)
End Sub

Private Sub BuildEvent_BeginCompile(ByVal VBProject As VBIDE.VBProject)
    MsgBox 1
End Sub

Private Sub BuildEvent_EnterDesignMode()
    MsgBox 2
End Sub

Private Sub BuildEvent_EnterRunMode()
    MsgBox 3
End Sub

Private Sub FCE_AfterWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal Result As Integer)
     MsgBox "AfterWriteFile " & FileName
End Sub

Private Sub FCE_BeforeLoadFile(ByVal VBProject As VBIDE.VBProject, FileNames() As String)
    MsgBox "BeforeLoadFile"
End Sub

Private Sub FCE_DoGetNewFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, NewName As String, ByVal OldName As String, CancelDefault As Boolean)
    MsgBox "DoGetNewFileName " & FileType
End Sub

Private Sub FCE_RequestChangeFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal NewName As String, ByVal OldName As String, Cancel As Boolean)
    MsgBox "RequestChangeFileName"
End Sub

Private Sub FCE_RequestWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileName As String, Cancel As Boolean)
    MsgBox "RequestWriteFile " & FileName
End Sub

Private Sub VBCE_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
    MsgBox VBControl.Properties(1)
End Sub

Private Sub VCE_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
    Set VBCE = VBIns.Events.VBControlsEvents(VBComponent.Collection.Parent, VBComponent.Designer)
End Sub

Private Sub VCE_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    MsgBox "VBComponent " & VBComponent.Name & " Type = " & VBComponent.Type
End Sub
