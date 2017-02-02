VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13830
   _ExtentX        =   24395
   _ExtentY        =   17992
   _Version        =   393216
   Description     =   "iCode"
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

Dim WithEvents RH As iRemoteHook
Attribute RH.VB_VarHelpID = -1

Dim mTipsBarHandler As clsTipsBarHandler '需要在本模块中加载再赋值给Public的TipsBarHandler，不然会出现迷之错误

'===ClipBoard Preserve===
Dim CB_Text As String
Dim CB_Data As IPictureDisp

Dim CB_bText As Boolean, CB_bData As Boolean

'=== Public ===

Private iCodeMenu As CommandBarPopup

Private iToolBar As CommandBar
Private Bar_InfoExist As Boolean
Private Bar_Left As Long
Private Bar_Position As MsoBarPosition
Private Bar_RowIndex As Long
Private Bar_Visible As Boolean

Private btnSetting As CommandBarButton
Private WithEvents btnSetting_Events As CommandBarEvents
Attribute btnSetting_Events.VB_VarHelpID = -1

Private btnHelp As CommandBarButton
Private WithEvents btnHelp_Events As CommandBarEvents
Attribute btnHelp_Events.VB_VarHelpID = -1

Private btnAbout As CommandBarButton
Private WithEvents btnAbout_Events As CommandBarEvents
Attribute btnAbout_Events.VB_VarHelpID = -1


Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    '===ClipBoard Preserve===
    GetClipBoard
    
    '=== Public ===
    Set VBIns = Application
    Set AddInIns = AddInInst
    Let hVBIDE = VBIns.MainWindow.hWnd
    
    '=== DebugForm ===
    #If 0 Then
        Set DebugForm = New frmDebug
        DebugForm.Show
    #End If
    
    '=== Public CommandBar ===
    CB_iCodeMenu_Load
    CB_iToolBar_Load
    CB_SettingButton_Load
    
    '=== Component ===
    Windows_Linker_Load
    IDEEnhancer_Load
    CodeStatistic_Load
    TipsBarHandler_Load ConnectMode
    CodeIndent_Load
    AutoComplete_Load
    ColorCode_Load
    
    CB_HelpButton_Load
    CB_AboutButton_Load '"关于"按钮要出现在最下方
    
    '=== Hook ===
    If App.LogMode = 0 Then
        Set RH = New iRemoteHook
        RH.SetMode CALLWNDPROCHOOK
        'Windows+Linker
        RH.RegisterMessage &H128
        RH.RegisterMessage WM_DESTROY
        'TipsBarHandler
        RH.RegisterMessage WM_MDIACTIVATE
        RH.RegisterMessage WM_MDIDESTROY
        RH.RegisterMessage WM_SIZE
        DBPrint "RemoteHook = " & RH.Inject(RH.GetThreadIDByhWnd(hVBIDE))
    Else
        SetMsgHooks
    End If
    
End Sub

'===== Public CommandBar =====
Public Sub CB_iCodeMenu_Load()
    Set iCodeMenu = VBIns.CommandBars("菜单条").Controls.Add(msoControlPopup, , , VBIns.CommandBars("菜单条").Controls("帮助(&H)").Index, True)
    iCodeMenu.Caption = "iCode"
End Sub

Public Sub CB_SettingButton_Load()
    Set btnSetting = iCodeMenu.Controls.Add(MsoControlType.msoControlButton)
    btnSetting.Caption = "设置(&S)"
    Clipboard.SetData LoadResPicture(101, 0)
    btnSetting.PasteFace
    btnSetting.Style = msoButtonAutomatic
    Set btnSetting_Events = VBIns.Events.CommandBarEvents(btnSetting)
End Sub

Public Sub CB_HelpButton_Load()
    Set btnHelp = iCodeMenu.Controls.Add(MsoControlType.msoControlButton)
    btnHelp.Caption = "帮助(&H)"
    btnHelp.Style = msoButtonCaption
    btnHelp.BeginGroup = True
    Set btnHelp_Events = VBIns.Events.CommandBarEvents(btnHelp)
End Sub

Public Sub CB_AboutButton_Load()
    Set btnAbout = iCodeMenu.Controls.Add(MsoControlType.msoControlButton)
    btnAbout.Caption = "关于(&A)"
    btnAbout.Style = msoButtonCaption
    'btnAbout.BeginGroup = True
    Set btnAbout_Events = VBIns.Events.CommandBarEvents(btnAbout)
End Sub

Public Sub CB_iToolBar_Load()
    Set iToolBar = VBIns.CommandBars.Add("iCode Tools")

    Bar_Left = Settings_Get("IDEEnhancer", "Bar_Left", -1)
    If Bar_Left <> -1 Then
        Bar_InfoExist = True
        Bar_Position = Settings_Get("IDEEnhancer", "Bar_Position", 0)
        Bar_RowIndex = Settings_Get("IDEEnhancer", "Bar_RowIndex", 0)
        Bar_Visible = CBool(Settings_Get("IDEEnhancer", "Bar_Visible", True))
    End If

    If Bar_InfoExist Then
        With iToolBar
            .Position = Bar_Position
            .Left = Bar_Left
            .RowIndex = Bar_RowIndex
            .Visible = Bar_Visible
        End With
    End If
    
    iToolBar.Visible = True
End Sub

Public Sub CB_iToolBar_UnLoad()
    With iToolBar
        Bar_Left = .Left
        Bar_Position = .Position
        Bar_RowIndex = .RowIndex
        Bar_Visible = .Visible
    End With
            
    Settings_Write "IDEEnhancer", "Bar_Left", Bar_Left
    Settings_Write "IDEEnhancer", "Bar_Position", Bar_Position
    Settings_Write "IDEEnhancer", "Bar_RowIndex", Bar_RowIndex
    Settings_Write "IDEEnhancer", "Bar_Visible", Bar_Visible
End Sub



'===== ClipBoard Preserve =====
Public Sub GetClipBoard()
    With Clipboard
        If .GetFormat(vbCFText) = True Then CB_Text = .GetText: CB_bText = True
        If .GetFormat(vbCFLink) Or _
           .GetFormat(vbCFBitmap) Or _
           .GetFormat(vbCFMetafile) Or _
           .GetFormat(vbCFDIB) Or _
           .GetFormat(vbCFPalette) Then _
           Set CB_Data = .GetData: CB_bData = True
    End With
End Sub

Public Sub SetClipBoard()
    If CB_bText = True Then Clipboard.SetText CB_Text
    If CB_bData = True Then Clipboard.SetData CB_Data
End Sub



'===== Component:ColorCode =====

Private Sub ColorCode_Load()
    Set ColorCode = New clsColorCode
    ColorCode.Initialize VBIns, DebugForm, iCodeMenu
End Sub

Private Sub ColorCode_UnLoad()
    Set ColorCode = Nothing
End Sub



'===== Component:AutoComplete =====

Private Sub AutoComplete_Load()
    Set AutoComplete = New clsAutoComplete
    Dim guid As String
    guid = Settings_Get("AutoComplete", "GUID", "")
    If guid = "" Then
        guid = GetGUID
        Settings_Write "AutoComplete", "GUID", guid
    End If
    AutoComplete.Initialize VBIns, DebugForm, AddInIns, iCodeMenu, guid
End Sub

Private Sub AutoComplete_UnLoad()
    Set AutoComplete = Nothing
End Sub



'===== Component:CodeIndent =====

Private Sub CodeIndent_Load()
    Set CodeIndent = New clsCodeIndent
    
    Dim guid As String
    guid = Settings_Get("CodeIndent", "GUID", "")
    If guid = "" Then
        guid = GetGUID
        Settings_Write "CodeIndent", "GUID", guid
    End If
    CodeIndent.Initialize VBIns, DebugForm, AddInIns, _
                          iCodeMenu, iToolBar, _
                          guid
    CodeIndent.Var_SpacePerLevel = CLng(Settings_Get("CodeIndent", "SpacePerLevel", "4"))
    CodeIndent.Var_MulitLines_IndentSpaceCount = CLng(Settings_Get("CodeIndent", "MulitLines_IndentSpaceCount", "0"))
    CodeIndent.Var_QuickButtonMode = CLng(Settings_Get("CodeIndent", "MulitLines_QuickButtonMode", "1"))
    CodeIndent.Var_ResultWindow_AutoHide = True
End Sub

Private Sub CodeIndent_UnLoad()
    Set CodeIndent = Nothing
End Sub

'===== Component:TipsBarHandler =====

Private Sub TipsBarHandler_Load(ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode)
    Set mTipsBarHandler = New clsTipsBarHandler
    mTipsBarHandler.Initialize VBIns, DebugForm, ConnectMode
    mTipsBarHandler.TipsBarAvliable = CBool(Settings_Get("TipsBarHandler", "Enabled", "True"))
    Set TipsBarHandler = mTipsBarHandler
End Sub

Private Sub TipsBarHandler_UnLoad()
    Set TipsBarHandler = Nothing
    Set mTipsBarHandler = Nothing
End Sub



'===== Component:Windows_Linker =====

Private Sub Windows_Linker_Load()
    
    Dim Param As Boolean
    Param = CBool(Settings_Get("Windows_Linker", "ResizeWindow", "False"))
    
    Set Windows_Linker = New clsWindowsHandler
    Windows_Linker.Initialize VBIns, DebugForm, Param
    
End Sub

Private Sub Windows_Linker_UnLoad()
    Set Windows_Linker = Nothing
End Sub


'===== Component:IDEEnhancer =====

Private Sub IDEEnhancer_Load()
    
    Set IDEEnhancer = New clsIDEEnhancer
    
    With IDEEnhancer
            
        .m_ChangeScope_Button_Visible = CBool(Settings_Get("IDEEnhancer", "ChangeScope_Button_Visible", True))
        .m_ChangeScope_Button_Style = Settings_Get("IDEEnhancer", "ChangeScope_Button_Style", MsoButtonStyle.msoButtonIcon)
        
        .m_Compile_Button_Visible = CBool(Settings_Get("IDEEnhancer", "Compile_Button_Visible", True))
        
        .m_ToCommon_Button_Visible = CBool(Settings_Get("IDEEnhancer", "ToCommon_Button_Visible", True))
        .m_ToCommon_Button_Style = Settings_Get("IDEEnhancer", "ToCommon_Button_Style", MsoButtonStyle.msoButtonIcon)
        
        .m_MakeExeButton_Enabled = CBool(Settings_Get("IDEEnhancer", "MakeExeButton_Enabled", False))
        .m_AddFile_Buttons_Enabled = CBool(Settings_Get("IDEEnhancer", "AddFile_Buttons_Enabled", False))
        
        .Initialize VBIns, DebugForm, iToolBar
    
    End With
    
End Sub

Private Sub IDEEnhancer_UnLoad()

    With IDEEnhancer
    
        .Msg_ToExit '获取按钮信息
        
        Settings_Write "IDEEnhancer", "ChangeScope_Button_Visible", CStr(.m_ChangeScope_Button_Visible)
        Settings_Write "IDEEnhancer", "ChangeScope_Button_Style", .m_ChangeScope_Button_Style
        Settings_Write "IDEEnhancer", "Compile_Button_Visible", CStr(.m_Compile_Button_Visible)
        Settings_Write "IDEEnhancer", "ToCommon_Button_Visible", CStr(.m_ToCommon_Button_Visible)
        Settings_Write "IDEEnhancer", "ToCommon_Button_Style", .m_ToCommon_Button_Style
        Settings_Write "IDEEnhancer", "MakeExeButton_Enabled", .m_MakeExeButton_Enabled
        Settings_Write "IDEEnhancer", "AddFile_Buttons_Enabled", .m_AddFile_Buttons_Enabled
        
    End With

    Set IDEEnhancer = Nothing
    
End Sub

'===== Component:CodeStatistic =====

Private Sub CodeStatistic_Load()
    Set CodeStatistic = New clsCodeStatistic
    CodeStatistic.Initialize VBIns, DebugForm, iCodeMenu
    CodeStatistic.MenuButton_BeginGroup = True
End Sub

Private Sub CodeStatistic_UnLoad()
    Set CodeStatistic = Nothing
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    
    '=== Public CommandBar ===
    CB_iToolBar_UnLoad
    
    '=== Hook ===
    If App.LogMode = 0 Then
        RH.UnInject
        Set RH = Nothing
    Else
        UnSetMsgHooks
    End If
    
    '=== Component ===
    Windows_Linker_UnLoad
    IDEEnhancer_UnLoad
    CodeStatistic_UnLoad
    TipsBarHandler_UnLoad
    CodeIndent_UnLoad
    AutoComplete_UnLoad
    ColorCode_UnLoad
    
    '=== Public ===
    Set iCodeMenu = Nothing
    
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    '===ClipBoard Preserve===
    SetClipBoard
End Sub

Private Sub btnHelp_Events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Shell "hh.exe " & App.Path & "\Help.chm", vbNormalFocus
End Sub

Private Sub btnAbout_Events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    frmAbout.Show
End Sub

Private Sub btnSetting_Events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    frmSetting.Show
End Sub

Private Sub RH_GetCallWndMessage(Result As Long, ByVal hWnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)
    MsgHook.iCallWndProc hWnd, Message, wParam, lParam
End Sub
