VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10200
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   13830
   _ExtentX        =   24395
   _ExtentY        =   17992
   _Version        =   393216
   Description     =   "iCode - Windows"
   DisplayName     =   "iCode - Windows"
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
    
    #If 1 Then
        Set DebugForm = New frmDebug
        DebugForm.Show
    #End If
    
    Set WND = New clsWindowsHandler
    
    WND.Initialize Application, DebugForm, False

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    
    If App.LogMode = 0 Then
        RH.UnInject
        Set RH = Nothing
    Else
        UnSetMsgHooks
    End If
    
    Set WND = Nothing
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    
    If App.LogMode = 0 Then
        Set RH = New iRemoteHook
        RH.SetMode CALLWNDPROCHOOK
        RH.RegisterMessage &H128
        RH.RegisterMessage WM_DESTROY
        DBPrint "RemoteHook = " & RH.Inject(RH.GetThreadIDByhWnd(hVBIDE))
    Else
        SetMsgHooks
    End If
    
End Sub

Private Sub RH_GetCallWndMessage(Result As Long, ByVal hWnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)
    MsgHook.iMsgProc hWnd, Message, wParam, lParam
End Sub
