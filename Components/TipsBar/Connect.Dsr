VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13830
   _ExtentX        =   24395
   _ExtentY        =   17965
   _Version        =   393216
   Description     =   "iCode - TipsBar"
   DisplayName     =   "iCode_TipsBar"
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

Dim WithEvents RH As iRemoteHook
Attribute RH.VB_VarHelpID = -1

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    
    Set VBIns = Application
    Let hVBIDE = VBIns.MainWindow.hWnd
    
    Set TBH = New clsTipsBarHandler
    
    #If 1 Then
        Set DebugForm = New frmDebug
        DebugForm.Show
    #End If
        
    If IsIDEMode Then
        Set RH = New iRemoteHook
        RH.SetMode CALLWNDPROCHOOK
        RH.RegisterMessage WM_SIZE
        RH.RegisterMessage WM_MDIACTIVATE
        RH.RegisterMessage WM_MDIDESTROY
        'RH.RegisterMessage WM_SHOWWINDOW
        DBPrint "RemoteHook = " & RH.Inject(RH.GetThreadIDByhWnd(hVBIDE))
    Else
        MsgHook.SetMsgHooks
        'WndHook.SetWndHook hVBIDE
    End If
    
    TBH.Initialize Application, DebugForm, ConnectMode
    TBH.TipsBarAvliable = True
    
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    
    If IsIDEMode Then
        RH.UnInject
        Set RH = Nothing
    Else
        MsgHook.UnSetMsgHooks
        'WndHook.UnSetWndHook
        'DBPrint "Exit: UnHook Succeed!"
    End If
    
    DoEvents
    
    'DBPrint "Exit: Before UnSet!"
    
    Set TBH = Nothing
    
    'DBPrint "Exit Succeed!"
    
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    IDELoadUp = True
End Sub

Private Sub RH_GetCallWndMessage(Result As Long, ByVal hWnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)
    MsgHook.iMsgProc hWnd, Message, wParam, lParam
End Sub
