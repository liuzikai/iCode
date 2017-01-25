VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10200
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   13830
   _ExtentX        =   24395
   _ExtentY        =   17992
   _Version        =   393216
   Description     =   "iCode - Statistic"
   DisplayName     =   "iCode_Statistic"
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

Dim clsCodeStatistic As clsCodeStatistic


Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    Set VBIns = Application
    Let hVBIDE = VBIns.MainWindow.hwnd
    
'    Set DebugForm = New frmDebug
'    DebugForm.Show
    
    Dim Menu As CommandBarPopup
    Set Menu = Application.CommandBars("²Ëµ¥Ìõ").Controls.Add(MsoControlType.msoControlPopup, , , , True)
    Menu.Caption = "iCode"
    
    Set clsCodeStatistic = New clsCodeStatistic
    
    clsCodeStatistic.Initialize Application, DebugForm, Menu
    
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    Set clsCodeStatistic = Nothing
End Sub
