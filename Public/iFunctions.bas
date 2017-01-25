Attribute VB_Name = "iFunctions"
Option Explicit

Public VBIns As VBE
Public hVBIDE As Long

Public DebugForm As Form

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function iSetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameW" (ByVal hWnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef pguid As guid) As Long

Private Type guid
    Data1 As Long
    Data4(0 To 7) As Byte
    Data3 As Integer
    Data2 As Integer
End Type

Public Function GetGUID() As String
    Dim lRetVal As Long
    Dim udtGuid As guid
    
    Dim sPartOne As String
    Dim sPartTwo As String
    Dim sPartThree As String
    Dim sPartFour As String
    Dim iDataLen As Integer
    Dim iStrLen As Integer
    Dim iCtr As Integer
    Dim sAns As String
    
    On Error GoTo ErrorHandler
    
    sAns = ""
    
    lRetVal = CoCreateGuid(udtGuid)
    
    If lRetVal = 0 Then
        
        sPartOne = Hex$(udtGuid.Data1)
        iStrLen = Len(sPartOne)
        iDataLen = Len(udtGuid.Data1)
        sPartOne = String((iDataLen * 2) - iStrLen, "0") & Trim$(sPartOne)
        
        sPartTwo = Hex$(udtGuid.Data2)
        iStrLen = Len(sPartTwo)
        iDataLen = Len(udtGuid.Data2)
        sPartTwo = String((iDataLen * 2) - iStrLen, "0") & Trim$(sPartTwo)
        
        sPartThree = Hex$(udtGuid.Data3)
        iStrLen = Len(sPartThree)
        iDataLen = Len(udtGuid.Data3)
        sPartThree = String((iDataLen * 2) - iStrLen, "0") & Trim$(sPartThree)
        
        For iCtr = 0 To 7
            sPartFour = sPartFour & Format$(Hex$(udtGuid.Data4(iCtr)), "00")
        Next
        
        sAns = sPartOne & sPartTwo & sPartThree & sPartFour
    End If
    
    GetGUID = sAns
    Exit Function
    
ErrorHandler:
    Exit Function
End Function

Public Function SetOnTop(ByVal hWnd As Long, ByVal IsOnTop As Boolean) As Long
    If IsOnTop Then
        SetOnTop = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Else
        SetOnTop = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If
End Function

Public Function iGetClassName(ByVal hWnd As Long) As String
    Dim s As String: s = String(256, 0)
    iGetClassName = Left(s, GetClassName(hWnd, StrPtr(s), 256))
End Function

Public Function iGetCaption(ByVal hWnd As Long) As String
    Dim s As String: s = String(256, 0)
    iGetCaption = Left(s, GetWindowText(hWnd, StrPtr(s), 256))
End Function

Public Function Max(ByVal a, ByVal b)
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function

Public Function Min(ByVal a, ByVal b)
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function

Public Function LeftIs(ByVal Str1 As String, ByVal Str2 As String) As Boolean
    LeftIs = (Left(LCase(Str1), Len(Str2)) = LCase(Str2))
End Function

Public Function iReplaceAll(ByVal str As String, ByVal Find As String, ByVal Replace As String) As String
    If str = "" Then Exit Function
    
    Dim a() As String
    a = Split(str, Find)
    
    iReplaceAll = a(0)
    
    Dim i As Long
    For i = 1 To UBound(a)
        iReplaceAll = iReplaceAll & Replace & a(i)
    Next
End Function

Public Sub DBPrint(ByVal str)
    If DebugForm Is Nothing Then Exit Sub
    
    DebugForm.txtMain.Text = DebugForm.txtMain.Text & str & vbCrLf
    DoEvents
End Sub

Public Function IsIDEMode() As Boolean
    IsIDEMode = (App.LogMode = 0)
End Function
