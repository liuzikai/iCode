VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindowsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private WithEvents t As TasksCollection
Attribute t.VB_VarHelpID = -1

Private Const iSection = "iCode_Windows"

Private Const WS_EX_ACCEPTFILES = &H10&
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)

Private clsLinker As clsLinker
Private TargetEXEFileName As String

Private WithEvents FileControlEvents As FileControlEvents
Attribute FileControlEvents.VB_VarHelpID = -1

Private EnhanceWindow As Boolean

Public Property Get Mode() As Boolean
    Mode = EnhanceWindow
End Property

Public Property Let Mode(ByVal Value As Boolean)
    EnhanceWindow = Value
    t.Clear
    Load_Basic_Tasks
End Property

Private Sub Load_Basic_Tasks()

    t.Add hVBIDE, "#32770", "生成工程", 263, 136, 630, 467, True
    
    If EnhanceWindow Then
    
        t.Add hVBIDE, "#32770", "文件另存为", 263, 136, 630, 467, True
        t.Add hVBIDE, "#32770", "工程另存为", 263, 136, 630, 467, True
        t.Add hVBIDE, "#32770", "工程组另存为", 263, 136, 630, 467, True
        
        t.Add hVBIDE, "#32770", "打开工程", 258, 123, 652, 523, True
        
        t.Add hVBIDE, "#32770", "添加文件", 263, 123, 630, 494, True
        
        '注意以下两项的父窗口并不是VB主窗口，而是部件/引用窗口
        t.Add , "#32770", "添加引用", 263, 136, 630, 467, True
        t.Add , "#32770", "添加 ActiveX 控件", 263, 136, 630, 467, True
    
    End If
    
End Sub

Public Sub Msg_Windows(ByVal hWnd As Long)
    t.Find hWnd
End Sub


Public Sub Initialize(ByVal IDEIns As VBE, ByVal DBForm As Object, ByVal WindowRedesign As Boolean)
    
    EnhanceWindow = WindowRedesign
    
    Set VBIns = IDEIns
    Let hVBIDE = VBIns.MainWindow.hWnd
    
    Set DebugForm = DBForm
        
    Set t = New TasksCollection
    
    Load_Basic_Tasks
    
    Set clsLinker = New clsLinker
    
    Set FileControlEvents = VBIns.Events.FileControlEvents(Nothing)
        
End Sub

Private Sub Class_Terminate()
    Set t = Nothing
    Set clsLinker = Nothing
    Set FileControlEvents = Nothing
End Sub


'此类窗口不触发WM_CLOSE消息
Public Sub Msg_WM_DESTROY_If_32770(ByVal hWnd As Long)
    '根据实际测试没有执行CleanUp操作的必要
    Select Case iGetCaption(hWnd)
    Case "生成工程"
        't.CleanUp hWnd
        clsLinker.Msg_CompileWindowDestory
    Case "文件另存为", "工程另存为", "工程组另存为", "打开工程", "添加文件", "添加引用", "添加 ActiveX 控件"
        't.CleanUp hWnd
    End Select
End Sub


Private Sub FileControlEvents_AfterWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal Result As Integer)
    If FileType = vbext_ft_Exe And FileName = TargetEXEFileName Then
        Shell LinkerPath & "iCodeLinker.exe " & FileName & "#" & LinkerData.pIcon & "#" & LinkerData.pManifest
        TargetEXEFileName = ""
        
        VBProject.WriteProperty iSection, "LastBuildFile", FileName
    End If
End Sub

Private Sub FileControlEvents_DoGetNewFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, NewName As String, ByVal OldName As String, CancelDefault As Boolean)
    
    Set clsLinker.CurProject = VBProject
    
    On Error Resume Next '避免没有记录项时ReadProperty引发错误
    
    Dim LastBuildFile As String
    LastBuildFile = VBProject.ReadProperty(iSection, "LastBuildFile")
    If LastBuildFile = "" Then Exit Sub
    
    If FileType = vbext_ft_Exe Then
        NewName = LastBuildFile
        VBProject.WriteProperty iSection, "LastComponentFile", OldName '保存此时目录
    Else
        If OldName = LastBuildFile Then '仅有由自动切换切回时恢复文件目录
            Dim LastComponentFile As String
            LastComponentFile = VBProject.ReadProperty(iSection, "LastComponentFile")
            If LastComponentFile <> "" Then NewName = LastComponentFile
        End If
    End If
    
End Sub

Private Sub FileControlEvents_RequestWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileName As String, Cancel As Boolean)
    If LCase(Right(FileName, 3)) = "exe" Then
        'clsLinker.Msg_StartCompile
        TargetEXEFileName = FileName
    End If
End Sub

Private Sub T_Find(Task As Task, ByVal hWnd As Long)
    
    Select Case Task.sClass
    
    Case "#32770"
    
        Select Case Task.sName
        
        Case ""
            If Task.Tag = "BuildWindow" Then t.Add hWnd, "Button", "选项(&O)...", 450, 380, 77, 23
            If Task.Tag = "AddFileWindow" Then t.Add hWnd, "Button", "添加为相关文档", 8, 440, 108, 15
        
        Case "生成工程"
            
            If EnhanceWindow Then
            
                t.Add hWnd, "#32770", "", 0, 0, 614, 439, , "BuildWindow"
            
                t.Add hWnd, "Static", "保存在(&I):", 11, 10, 72, 14
                t.Add hWnd, "ComboBox", "*Any*", 86, 5, 420, 24, , "WS_EX_ACCEPTFILES"
                t.Add hWnd, "ToolbarWindow32", "", 513, 5, 132, 26
                t.Add hWnd, "Static", "文件名(&N):", 8, 386, 66, 14
                t.Add hWnd, "Edit", "*Any*", 89, 383, 350, 20
                t.Add hWnd, "Button", "确定", 541, 380, 75, 23
                t.Add hWnd, "Button", "取消", 541, 408, 75, 23
                t.Add hWnd, "Button", "帮助(&H)", 6, 408, 70, 23
                t.Add hWnd, "SHELLDLL_DefView", "", 6, 36, 610, 337
                
                If Not VBIns.ActiveVBProject Is Nothing Then
                    If VBIns.ActiveVBProject.Type = vbext_pt_StandardExe Then
                        If clsLinker.Msg_CompileWindowShow(hWnd) Then
                            MoveWindow clsLinker.hContainer, 407, 408, 120, 23, True
                        End If
                    End If
                End If
            
            Else
                
                If Not VBIns.ActiveVBProject Is Nothing Then
                    If VBIns.ActiveVBProject.Type = vbext_pt_StandardExe Then
                        If clsLinker.Msg_CompileWindowShow(hWnd) Then
                            
                            Dim r As RECT
                            GetClientRect hWnd, r
                            
                            MoveWindow clsLinker.hContainer, 10, r.Bottom - r.Top - 33, 120, 23, True
                        End If
                    End If
                End If
                
            End If
            
        Case "添加文件"

            t.Add hWnd, "#32770", "", 0, 0, 614, 466, , "AddFileWindow"

            t.Add hWnd, "Static", "查找范围(&I):", 11, 10, 72, 14

            t.Add hWnd, "ComboBox", "*Any*", 86, 5, 420, 24, , "WS_EX_ACCEPTFILES"
            t.Add hWnd, "ComboBox", "*Any*", 89, 411, 350, 20

            t.Add hWnd, "ToolbarWindow32", "", 513, 5, 100, 26

            t.Add hWnd, "Static", "文件名(&N):", 8, 386, 66, 14
            t.Add hWnd, "Edit", "", 89, 383, 437, 20

            t.Add hWnd, "Static", "文件类型(&T):", 8, 414, 80, 14


            t.Add hWnd, "Button", "打开(&O)", 541, 380, 75, 23
            t.Add hWnd, "Button", "取消", 541, 408, 75, 23
            t.Add hWnd, "Button", "帮助(&H)", 451, 408, 75, 23

            t.Add hWnd, "SHELLDLL_DefView", "", 6, 36, 610, 337
        
            
            
        Case "文件另存为", "工程另存为", "工程组另存为"
            
            t.Add hWnd, "Static", "保存在(&I):", 11, 10, 72, 14
    
            t.Add hWnd, "ComboBox", "*Any*", 86, 5, 420, 24, , "WS_EX_ACCEPTFILES"
            t.Add hWnd, "ComboBox", "*Any*", 89, 411, 350, 20
            
            t.Add hWnd, "ToolbarWindow32", "", 513, 5, 100, 26
            
            t.Add hWnd, "Static", "文件名(&N):", 8, 386, 66, 14
            t.Add hWnd, "Edit", "*Any*", 99 - 10, 409 - 26, 437, 20
            
            t.Add hWnd, "Static", "保存类型(&T):", 8, 414, 80, 14
            
            t.Add hWnd, "Button", "保存(&S)", 541, 380, 75, 23
            t.Add hWnd, "Button", "取消", 541, 408, 75, 23
            t.Add , "Button", "帮助(&H)", 451, 408, 75, 23
            
            t.Add hWnd, "SHELLDLL_DefView", "", 6, 36, 610, 337
        
        Case "打开工程"
        
            t.Add hWnd, "Static", "查找范围(&I):", 21, 36, 72, 14
    
            t.Add hWnd, "ComboBox", "*Any*", 96, 31, 420, 24, , "WS_EX_ACCEPTFILES"
            
            t.Add hWnd, "ToolbarWindow32", "", 523, 31, 100, 26
            
            t.Add hWnd, "Static", "文件名(&N):", 18, 412, 66, 14
            t.Add hWnd, "Edit", "", 99, 409, 437, 20
            
            t.Add hWnd, "Static", "文件类型(&T):", 18, 440, 80, 14
            t.Add hWnd, "ComboBox", "*Any*", 99, 437, 350, 20
            
            t.Add hWnd, "Button", "打开(&O)", 551, 406, 75, 23
            t.Add hWnd, "Button", "取消", 551, 434, 75, 23
            t.Add hWnd, "Button", "帮助(&H)", 461, 434, 75, 23
            
            t.Add hWnd, "SHELLDLL_DefView", "", 16, 62, 610, 337
            
            t.Add hWnd, "SysListView32", "", 16, 32, 610, 367
            
            t.Add hWnd, "SysTabControl32", "", 6, 6, 634, 469
        
        
        Case "添加引用", "添加 ActiveX 控件"
            
            t.Add hWnd, "Static", "查找范围(&I):", 11, 10, 72, 14
    
            t.Add hWnd, "ComboBox", "*Any*", 86, 5, 420, 24, , "WS_EX_ACCEPTFILES"
            
            t.Add hWnd, "ToolbarWindow32", "", 513, 5, 100, 26
            
            t.Add hWnd, "Static", "文件名(&N):", 8, 386, 66, 14
            t.Add hWnd, "Edit", "", 99 - 10, 409 - 26, 437, 20
            
            t.Add hWnd, "Static", "文件类型(&T):", 8, 414, 80, 14
            t.Add hWnd, "ComboBox", "*Any*", 89, 411, 350, 20
            
            t.Add hWnd, "Button", "打开(&O)", 541, 380, 75, 23
            t.Add hWnd, "Button", "取消", 541, 408, 75, 23
            t.Add hWnd, "Button", "帮助(&H)", 451, 408, 75, 23
            
            t.Add hWnd, "SHELLDLL_DefView", "", 6, 36, 610, 337
        
        End Select
    
    Case "ComboBox"
        
        Dim lngStyle As Long
        lngStyle = GetWindowLong(hWnd, GWL_STYLE)
        
        If Task.Tag = "WS_EX_ACCEPTFILES" Then
            If (lngStyle And WS_EX_ACCEPTFILES) = False Then Exit Sub
        Else
            If (lngStyle And WS_EX_ACCEPTFILES) Then Exit Sub
        End If
    
    End Select
    
    If EnhanceWindow Then
        MoveWindow hWnd, Task.x, Task.y, Task.Width, Task.Height, False
    End If
    
End Sub
