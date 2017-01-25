VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call CreateActiveMenu
End Sub

Sub CreateActiveMenu()
    Dim hMenu As Long, hSubMenu As Long
    Dim hPopMenuTmp As Long
    ReDim MenuText(0)
    
    hMenu = GetMenu(Me.hwnd) '窗体级菜单句柄
    If hMenu = 0 Then
        '窗体上没有菜单时，创建菜单。这种情况下需在设计阶段设置窗体的NegotiatMenu=False菜单才能显示出来。
        hMenu = CreateMenu()
    End If
    
    '添加到0级菜单
    hSubMenu = hMenu
    FullAllSubMenu hSubMenu
    
    '添加到1级菜单
    hSubMenu = GetSubMenu(hSubMenu, GetMenuItemCount(hSubMenu) - 1) '获取最后一个0级菜单的句柄
    FullAllSubMenu hSubMenu
    
    '添加到2级菜单
    hSubMenu = GetSubMenu(hSubMenu, GetMenuItemCount(hSubMenu) - 1)
    FullAllSubMenu hSubMenu
    
    '添加到3级菜单
    hSubMenu = GetSubMenu(hSubMenu, GetMenuItemCount(hSubMenu) - 1)
    FullAllSubMenu hSubMenu
    
    SetMenu Me.hwnd, hMenu
    DrawMenuBar Me.hwnd
    Me.Refresh
    
    OldWinProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf OnMenu)
End Sub

Sub FullAllSubMenu(hFather As Long)
    '加入全部子菜单
    Dim hPopMenuTmp As Long
    Dim i As Integer
    hPopMenuTmp = CreatePopupMenu()
    For i = 0 To 4
        MenuCount = MenuCount + 1
        '保存菜单文本，用于菜单事件触发时识别出被选择的菜单对象
        ReDim Preserve MenuText(MenuCount)
        MenuText(MenuCount) = "文件" & MenuCount
        '加入子菜单，令其ID>1000，说明其为自动生成的菜单
        AppendMenu1 hPopMenuTmp, MF_STRING, 1000 + MenuCount, MenuText(MenuCount)
        '如果是间隔线，则wFlags=MF_SEPARATOR
        '如果要Check，则wFlags=MF_STRING + MF_CHECKED，若令不可用，则再加MF_GRAYED
    Next
    AppendMenu1 hFather, MF_POPUP, hPopMenuTmp, "&Files"
End Sub
