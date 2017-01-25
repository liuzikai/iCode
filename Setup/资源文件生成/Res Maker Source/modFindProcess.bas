Attribute VB_Name = "modFindProcess"
Option Explicit

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwsize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Const TH32CS_SNAPheaplist = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPthread = &H4
Private Const TH32CS_SNAPmodule = &H8
Private Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule

Public Function FindPID(ByVal PID As Long) As Boolean
    
    Dim ret As Long
    Dim Proc As PROCESSENTRY32
    Dim hSnapshot As Long
    
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
    
    ret = ProcessFirst(hSnapshot, Proc)
    Do While ret <> 0
        If Proc.th32ProcessID = PID Then
            FindPID = True
            GoTo SubEnd
        End If
        ret = ProcessNext(hSnapshot, Proc) '循环获取下一个进程的PROCESSENTRY32结构信息数据
    Loop
    
SubEnd:
    CloseHandle hSnapshot '关闭进程“快照”句柄
    
End Function

