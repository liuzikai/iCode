VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpW" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long                                   '上下文相关的帮助
Private Const HH_TP_HELP_CONTEXTMENU = &H10
Private Const HH_GET_LAST_ERROR = &H14
Private Const HH_DISPLAY_TOC = &H1

Const HH_DISPLAY_TOPIC = &H0
Const HH_DISPLAY_INDEX = &H2
Const HH_HELP_CONTEXT = &HF
Const HH_DISPLAY_SEARCH = &H3
Const HH_DISPLAY_TEXT_POPUP = &HE

Private Type HH_LAST_ERROR
  cbStruct As Long
  hr As Long
  Description As String
End Type

Private Const HH_INITIALIZE = &H1C

Private Sub Form_Load()
    Dim sHelpFile As String
    sHelpFile = "E:\MSDN Library\98VS\2052\msdnvs98.col"
    Dim a(0 To 4) As Long
    a(0) = 0
    a(1) = 1243800
    a(2) = 0
    a(3) = -1
    a(4) = 0
    
    
    
'    CommonDialog1.HelpFile = sHelpFile
'    CommonDialog1.HelpContext = 0
'    'CommonDialog1.HelpCommand = &H1
'    CommonDialog1.showHelp
    
    
    
    Call Htmlhelp(Me.hWnd, sHelpFile, HH_DISPLAY_TOPIC, 0&)
    
'    Dim e As HH_LAST_ERROR
'
'    e.cbStruct = Len(e)
'    HtmlHelp Me.hWnd, sHelpFile, HH_GET_LAST_ERROR, e
'
'    MsgBox e.hr
'    MsgBox e.Description
End Sub
