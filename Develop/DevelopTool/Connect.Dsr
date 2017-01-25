VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11715
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   16230
   _ExtentX        =   28628
   _ExtentY        =   20664
   _Version        =   393216
   Description     =   "iCode - DevelopTool"
   DisplayName     =   "iCode - DevelopTool"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
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
Private Button As CommandBarButton
Private WithEvents ButtonEvent As CommandBarEvents
Attribute ButtonEvent.VB_VarHelpID = -1

Private Button2 As CommandBarButton
Private WithEvents Button2Event As CommandBarEvents
Attribute Button2Event.VB_VarHelpID = -1

Private TargetFile As String

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    Set VBIns = Application
    
    Dim Common As CommandBar
    Set Common = VBIns.CommandBars("±ê×¼")
    
    If Not (Common Is Nothing) Then
        Set Button = Common.Controls.Add(MsoControlType.msoControlButton, , , , True)
        If Not (Button Is Nothing) Then
            Button.Caption = "Compile"
            Button.Style = msoButtonCaption
            Set ButtonEvent = VBIns.Events.CommandBarEvents(Button)
        End If
        Set Button2 = Common.Controls.Add(MsoControlType.msoControlButton, , , , True)
        If Not (Button2 Is Nothing) Then
            Button2.Caption = "Start VB"
            Button2.Style = msoButtonCaption
            Set Button2Event = VBIns.Events.CommandBarEvents(Button2)
        End If
    End If
    
End Sub

Private Sub Button2Event_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    DoEvents
    Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE", vbMaximizedFocus
End Sub

Private Sub ButtonEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error GoTo iErr
    If Not (VBIns.ActiveVBProject Is Nothing) Then
        VBIns.ActiveVBProject.MakeCompiledFile
    End If
    
    Exit Sub
iErr:
    MsgBox "´íÎó£¡Çë¼ì²é´úÂë´íÎó£¡", vbExclamation, "Develop Tool"
    
End Sub
