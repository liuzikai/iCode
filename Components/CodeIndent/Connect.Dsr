VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13830
   _ExtentX        =   24395
   _ExtentY        =   17965
   _Version        =   393216
   Description     =   "iCode - CodeIndent"
   DisplayName     =   "iCode_CodeIndent"
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

Dim CodeIndent As New clsCodeIndent
Dim BC As ButtonsCollection

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    Set VBIns = Application
    
    #If 1 Then
        Set DebugForm = New frmDebug
        DebugForm.Show
    #End If
    
    Set BC = New ButtonsCollection
    BC.SetTarget VBIns.CommandBars("菜单条")
    
    Dim iMenu As CommandBarPopup
    Set iMenu = BC.Create(msoControlPopup, "iCode")
    
    CodeIndent.Initialize Application, DebugForm, AddInInst, _
                          iMenu, VBIns.CommandBars("菜单条"), _
                          GetGUID
    CodeIndent.Var_SpacePerLevel = 4
    CodeIndent.Var_MulitLines_IndentSpaceCount = 0
    CodeIndent.Var_QuickButtonMode = 2
    CodeIndent.Var_ResultWindow_AutoHide = True
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    Set CodeIndent = Nothing
    Set BC = Nothing
End Sub

