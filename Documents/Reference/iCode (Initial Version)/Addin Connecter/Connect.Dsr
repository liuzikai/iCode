VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9915
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   13995
   _ExtentX        =   24686
   _ExtentY        =   17489
   _Version        =   393216
   Description     =   "iCode - 全新的 VB 6.0 增强插件"
   DisplayName     =   "iCode"
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

Private Const lCBTHook As Boolean = False
Private Const lMsgHook As Boolean = False
Private Const lErrorHandle As Boolean = False
Private Const lTipsBar As Boolean = False
Private Const lProjectWindow As Boolean = False
Private Const lClipBoardControler As Boolean = True
Private Const lDesignerControler As Boolean = False



Private Const TipsBarMode As Long = tbNormal

Dim iClipBoard As New clsClipBoardControler

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    DBInit
    
    GDIP.InitGDIPlus
    
    LoadPublicValues Application
    
    
    If lClipBoardControler Then iClipBoard.GetClipBoard
    
    If lProjectWindow Then
        Set WndProject = VBIns.Windows.CreateToolWindow(AddInInst, "iCode_Project.udProject", "Project", GetGUID, iProject)
        WndProject.Visible = True
        
    End If
    
    If lTipsBar Then
        iTipsBar.Mode = TipsBarMode
        iTipsBar.Init
    End If
    
    
    If CBool(App.LogMode) = True Then
        If lCBTHook Then CBTHook.SetCBTHook
        If lMsgHook Then MsgHook.SetMsgHooks
        
    End If
    
    
    If lDesignerControler Then Set iDesigner = New clsDesignerControler
    
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    If lTipsBar Then iTipsBar.LoadWindows
    
    If CBool(App.LogMode) = True Then
        If lErrorHandle Then iCode.CWE.EHHookMessageBox
    End If
    
    
    If lProjectWindow Then
        iProject.Init
    End If
    
    If lClipBoardControler Then
        iClipBoard.SetClipBoard
        Set iClipBoard = Nothing
    End If
    
    Set iCode = New clsCodeControler
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    If CBool(App.LogMode) = True Then
        
        If lErrorHandle Then iCode.CWE.EHUnHookMessageBox
        If lCBTHook Then CBTHook.UnSetCBTHook
        If lMsgHook Then MsgHook.UnSetMsgHooks
        If lTipsBar Then iTipsBar.UnLoad
        
    End If
    
    If lDesignerControler Then Set iDesigner = Nothing
    
    GDIP.TerminateGDIPlus
    
    DBUnLoad
End Sub


