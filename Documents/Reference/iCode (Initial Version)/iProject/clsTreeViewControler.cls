VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTreeViewControler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Structures
Private Type TVITEM
    mask As Long
    hItem As Long
    state As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
End Type

'Private constants

Private Const TVIF_TEXT = 1
Private Const TVIF_HANDLE = &H10
Private Const TV_FIRST = &H1100&
Private Const TVM_GETCOUNT = TV_FIRST + 5
Private Const TVM_SELECTITEM = TV_FIRST + 11
Private Const TVM_GETITEM = TV_FIRST + 12
Private Const TVM_GETNEXTITEM = TV_FIRST + 10
Private Const TVGN_ROOT = 0
Private Const TVGN_NEXT = 1
Private Const TVGN_CHILD = 4
Private Const TVGN_CARET = 9

'API declarations
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'============NT Shared memory staff======================
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_VM_WRITE = &H20
Private Const PROCESS_ALL_ACCESS = 0
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const MEM_DECOMMIT = &H4000
Private Const MEM_RELEASE = &H8000
Private Const MEM_FREE = &H10000
Private Const MEM_PRIVATE = &H20000
Private Const MEM_MAPPED = &H40000
Private Const MEM_TOP_DOWN = &H100000

'==========Memory access private constants===========
Private Const PAGE_NOACCESS = &H1&
Private Const PAGE_READONLY = &H2&
Private Const PAGE_READWRITE = &H4&
Private Const PAGE_WRITECOPY = &H8&
Private Const PAGE_EXECUTE = &H10&
Private Const PAGE_EXECUTE_READ = &H20&
Private Const PAGE_EXECUTE_READWRITE = &H40&
Private Const PAGE_EXECUTE_WRITECOPY = &H80&
Private Const PAGE_GUARD = &H100&
Private Const PAGE_NOCACHE = &H200&





Dim tvHwnd As Long
Dim tvProcess As Long
Dim mlpTreeItemRemote As Long, mlpTextRemote As Long
Dim mszBuf() As Byte

Private Const mnMaxLen As Long = 1023

Private mStackFlag As CStack
Private mStackNode As CStack


Event ViewItem(hwnd As Long, sText As String, nLevel As Long)

Public Function GetNodeText(ByVal hNode As Long) As String
    On Error Resume Next
    
    Dim dwBytesRead As Long, dwBytesWrite As Long
    Dim lvItemLocal As TVITEM
    
    ReDim mszBuf(mnMaxLen)
    Call WriteProcessMemory(ByVal tvProcess, ByVal mlpTextRemote, mszBuf(0), mnMaxLen + 1, dwBytesWrite)
    
    lvItemLocal.hItem = hNode
    lvItemLocal.mask = TVIF_TEXT + TVIF_HANDLE
    lvItemLocal.cchTextMax = mnMaxLen
    lvItemLocal.pszText = mlpTextRemote
    
    Call WriteProcessMemory(ByVal tvProcess, ByVal mlpTreeItemRemote, ByVal VarPtr(lvItemLocal), Len(lvItemLocal), dwBytesWrite)
    
    Call SendMessage(tvHwnd, TVM_GETITEM, 0&, ByVal mlpTreeItemRemote)
    
    Call ReadProcessMemory(ByVal tvProcess, ByVal mlpTextRemote, mszBuf(0), mnMaxLen + 1, dwBytesRead)
    
    GetNodeText = StrConv(LeftB(mszBuf, InStrB(mszBuf, ChrB(0)) - 1), vbUnicode)
End Function

Public Sub SetTreeViewHwnd(ByVal nHwnd)
    tvHwnd = nHwnd
End Sub

Private Sub OpenTVProcess()
    Dim dwProcessId As Long
    GetWindowThreadProcessId tvHwnd, dwProcessId
    tvProcess = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, 0&, dwProcessId)
End Sub

Private Sub CloseTVProcess()
    CloseHandle tvProcess
    tvProcess = 0
End Sub

Private Function WalkNode(ByVal hNode As Long) As Boolean
    On Error GoTo iError
    
    Dim lLevel As Long
    Dim fChild As Boolean, fEnd As Boolean
    Dim hTmpNode As Long
    Dim hChild As Long, hSibling As Long
    Dim sItemText As String
    hTmpNode = hNode
    
bgChild:
    
    fChild = False
    
    sItemText = GetNodeText(hTmpNode)
    
    RaiseEvent ViewItem(hTmpNode, sItemText, lLevel)
    
    hChild = SendMessage(tvHwnd, TVM_GETNEXTITEM, TVGN_CHILD, ByVal hTmpNode)
    If hChild <> 0 Then
        fChild = True
        mStackNode.Push hTmpNode
        mStackFlag.Push fChild
        hTmpNode = hChild
        lLevel = lLevel + 1
        GoTo bgChild
    End If
    
bgSibling:
    hSibling = SendMessage(tvHwnd, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hTmpNode)
    If hSibling <> 0 Then
        If lLevel = 0 And Not fEnd Then
            mStackNode.Push hTmpNode
            mStackFlag.Push fChild
        End If
        hTmpNode = hSibling
        GoTo bgChild
    Else
        If lLevel = 0 Then
            fEnd = (lLevel = 0)
            GoTo iFinish
        End If
    End If
    
    
    If mStackNode.Pop(hTmpNode) And mStackFlag.Pop(fChild) Then
        If Not fChild Then
            GoTo bgChild
        Else
            If lLevel > 0 Then
                lLevel = lLevel - 1
            End If
            GoTo bgSibling
        End If
    End If
    
iFinish:
    WalkNode = True
    Exit Function
iError:
    WalkNode = False
    Exit Function
End Function

Public Function ViewAllNode() As Boolean
    Dim dwBytesRead As Long, dwBytesWrite As Long
    Dim lpTreeItemRemote As Long, lpTextRemote As Long
    
    ReDim mszBuf(mnMaxLen)
    Dim lvItemLocal As TVITEM
    
    OpenTVProcess
    
    If tvProcess <> 0 Then
        mlpTextRemote = VirtualAllocEx(ByVal tvProcess, ByVal 0&, mnMaxLen + 1, MEM_COMMIT, PAGE_READWRITE)
        mlpTreeItemRemote = VirtualAllocEx(ByVal tvProcess, ByVal 0&, Len(lvItemLocal), MEM_COMMIT, PAGE_READWRITE)
        
        Dim hRoot As Long
        
        Set mStackNode = New CStack
        Set mStackFlag = New CStack
        
        hRoot = SendMessage(tvHwnd, TVM_GETNEXTITEM, TVGN_ROOT, ByVal 0&)
        If hRoot <> 0 Then
            ViewAllNode = WalkNode(hRoot)
        Else
            ViewAllNode = False
        End If
        
        Do Until hRoot = 0
            hRoot = SendMessage(tvHwnd, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hRoot)
            If hRoot <> 0 Then
                ViewAllNode = ViewAllNode Or WalkNode(hRoot)
            End If
        Loop
        
        Set mStackNode = Nothing
        Set mStackFlag = Nothing
        
        Call VirtualFreeEx(tvProcess, ByVal mlpTreeItemRemote, 0, MEM_DECOMMIT)
        Call VirtualFreeEx(tvProcess, ByVal mlpTextRemote, 0, MEM_DECOMMIT)
    End If
    
    CloseTVProcess
End Function

Public Function SelectNode(ByVal hNode As Long) As Boolean
    SelectNode = SendMessage(tvHwnd, TVM_SELECTITEM, TVGN_CARET, ByVal hNode)
End Function
