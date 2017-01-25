VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColorCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private SettingFileName As String

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private iCodeMenu As CommandBarPopup
Private iPopup As CommandBarPopup
Private btnSetting As CommandBarButton
Private WithEvents btnSetting_Events As CommandBarEvents
Attribute btnSetting_Events.VB_VarHelpID = -1

Private BaseAddr_Property As Long

Private WithEvents Buttons As ButtonCollection
Attribute Buttons.VB_VarHelpID = -1

Private Sub LoadColors()
    If Dir(SettingFileName) <> "" Then
        Open SettingFileName For Input As #1
        Dim i As Long, j As Long, s As String, p() As String
        Input #1, i
        modPublic.Color_Count = 1 + i
        For i = 1 To modPublic.Color_Count - 1
            Input #1, Colors(i).Caption
            Input #1, Colors(i).Creator
            Input #1, s
            p = Split(s, ",")
            For j = 0 To 29
                Colors(i).Color(j) = CLng(p(j))
            Next
        Next
        Close #1
    End If
End Sub

Public Sub WriteColors()

    Open SettingFileName For Output As #1
    Dim i As Long, j As Long, s As String
    Write #1, modPublic.Color_Count - 1
    For i = 1 To modPublic.Color_Count - 1
        Write #1, Colors(i).Caption
        Write #1, Colors(i).Creator
        If Colors(i).Color(0) = "-1" Then s = "-1" Else s = "&H" & Hex(Colors(i).Color(0))
        For j = 1 To 29
            If Colors(i).Color(j) = "-1" Then s = s & ",-1" Else s = s & ",&H" & Hex(Colors(i).Color(j))
        Next
        Write #1, s
    Next
    Close #1
    
    LoadButtons
    
End Sub

Private Sub LoadButtons()
    Dim i As Long, btn As CommandBarControl
    For i = 1 To Buttons.Count
        Buttons.Item(i).Delete
    Next
    Buttons.Clear
    For i = 0 To Color_Count - 1
        Set btn = iPopup.Controls.Add(msoControlButton, , , , True)
        btn.Caption = Colors(i).Caption
        btn.Tag = i
        btn.Visible = True
        If i = 0 Then btn.BeginGroup = True
        Buttons.Push btn
    Next
End Sub

Private Sub GetBaseAddr()
    CopyMemory BaseAddr_Property, ByVal (GetModuleHandle("VBA6.DLL") + &H17D0F4), 4
End Sub

Private Sub SetProperty(ByVal Offset As Long, ByVal Value As Long)
    CopyMemory ByVal (BaseAddr_Property + Offset), Value, 4
End Sub

Public Sub InitColors()
    Colors(0).Caption = "默认"
    Dim i As Long, j As Long
    For i = 0 To 2
        For j = 0 To 9
            CopyMemory Colors(0).Color(i * 10 + j), ByVal (BaseAddr_Property + (&H19C + &H40 * i) + 4 * j), 4
        Next
    Next
End Sub

Public Sub Initialize(ByVal VBIDE As Object, ByVal DebugForm As Object, ByVal iMenu As Object)
    
    Set iFunctions.VBIns = VBIDE
    Let iFunctions.hVBIDE = VBIns.MainWindow.hWnd
    Set iFunctions.DebugForm = DebugForm
    Set iCodeMenu = iMenu
    
    GetBaseAddr
    SettingFileName = Environ("APPDATA") & "\iCode\ColorSetting.ic_c"
    
    Set iPopup = iCodeMenu.Controls.Add(msoControlPopup, , , , True)
    iPopup.Caption = "代码配色"
    iPopup.BeginGroup = True
    iPopup.Visible = True
    
    Set btnSetting = iPopup.Controls.Add(msoControlButton, , , , True)
    btnSetting.Caption = "设置配色方案"
    btnSetting.Visible = True
    Clipboard.SetData LoadResPicture(102, 0)
    btnSetting.PasteFace
    Set btnSetting_Events = VBIns.Events.CommandBarEvents(btnSetting)
    
    Set Buttons = New ButtonCollection
    
    Color_Count = 1
    InitColors
    LoadColors
    LoadButtons

    
End Sub

Private Sub btnSetting_Events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    InitColors
    frmSetting.Show
    Set frmSetting.clsParent = Me
End Sub

Private Sub Buttons_Click(ByVal Button As Office.CommandBarControl)
    Dim i As Long, j As Long
    For i = 0 To 2
        For j = 0 To 9
            SetProperty (&H19C + &H40 * i) + 4 * j, Colors(Button.Tag).Color(i * 10 + j)
        Next
    Next
End Sub
