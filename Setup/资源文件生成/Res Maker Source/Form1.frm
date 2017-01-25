VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Res制作"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17670
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   17670
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text5 
      Height          =   7035
      Left            =   12060
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   420
      Width           =   5475
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   4680
      TabIndex        =   10
      Text            =   "0"
      Top             =   7860
      Width           =   1395
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3.制作Res文件"
      Height          =   435
      Left            =   14580
      TabIndex        =   5
      Top             =   7680
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2.翻译"
      Height          =   435
      Left            =   10800
      TabIndex        =   4
      Top             =   7680
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1.生成文件列表"
      Height          =   435
      Left            =   6900
      TabIndex        =   3
      Top             =   7680
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   7860
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   7035
      Left            =   6360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   420
      Width           =   5475
   End
   Begin VB.TextBox Text1 
      Height          =   7035
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   420
      Width           =   5955
   End
   Begin VB.Label Label5 
      Caption         =   "Uninstall.rc："
      Height          =   195
      Left            =   12060
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "文件夹个数："
      Height          =   195
      Left            =   4680
      TabIndex        =   9
      Top             =   7560
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "Setup.rc："
      Height          =   195
      Left            =   6360
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "省略前缀："
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "原始文件列表："
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim File(100) As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long


Private Sub Command1_Click()

    Dim PID As Long, s As String
    
    PID = Shell("cmd /c dir """ & App.Path & "\File" & """ /OG /s /b >""" & App.Path & "\File List.txt""", vbHide)
    Do While FindPID(PID)
        DoEvents
    Loop
    DoEvents
    
    Sleep 200
    
    Open App.Path & "\File List.txt" For Binary As #1
    s = Space(LOF(1))
    Get #1, , s
    Close #1
    
    Text1 = s
    
End Sub

Private Sub Command2_Click()

    Dim i As Long, j As Long, k As Long, c As Long
    Dim p() As String, s As String
    
    p = Split(Text1, vbCrLf)
    For i = 0 To UBound(p)
        If p(i) = "" Then Exit For
        File(i) = Right(p(i), Len(p(i)) - Len(Text3))
        c = i
    Next
    
    Text2 = "STRINGTABLE" & vbCrLf & "BEGIN" & vbCrLf & "100, """ & Text4 & "," & c + 1 & """" & vbCrLf
    For i = 0 To c
        Text2 = Text2 & 101 + i & ", """ & File(i) & """" & vbCrLf
    Next
    
    Text2 = Text2 & "END" & vbCrLf
    
    Text5 = Text2
    
    Text2 = Text2 & vbCrLf
    
    For i = 0 To c
        If i >= Text4 Then
            Text2 = Text2 & 101 + i & " CUSTOM """ & File(i) & """" & vbCrLf
        End If
    Next
    
End Sub

Private Sub Command3_Click()

    Dim s As String, PID As Long

    Open App.Path & "\File\Setup.rc" For Output As #1
    s = Text2
    Print #1, s
    Close #1
    
    Open App.Path & "\File\Uninstall.rc" For Output As #1
    s = Text5
    Print #1, s
    Close #1
    
    DoEvents
    
    PID = Shell("cmd /k " & App.Path & "\RC.exe /v /l804 /fo " & App.Path & "\Setup.res " & App.Path & "\File\Setup.rc", vbNormalFocus)
    Do While FindPID(PID)
        DoEvents
    Loop
    Sleep 200
    DoEvents
    
    PID = Shell("cmd /k " & App.Path & "\RC.exe /v /l804 /fo " & App.Path & "\Uninstall.res " & App.Path & "\File\Uninstall.rc", vbNormalFocus)
    Do While FindPID(PID)
        DoEvents
    Loop
    Sleep 200
    DoEvents
    
    MsgBox "编译指令已发送！点击确定清理文件", 0, "提示"
    
    DeleteFile App.Path & "\File\Setup.rc"
    DeleteFile App.Path & "\File\Uninstall.rc"
    
End Sub

Private Sub Form_Load()
    Text3 = App.Path & "\File\"
End Sub
