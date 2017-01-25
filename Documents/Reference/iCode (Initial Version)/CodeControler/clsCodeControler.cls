VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeControler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const lCodeQuick As Boolean = True
Private Const lCodeSort As Boolean = False
Private Const lCWE As Boolean = False
Private Const lAC As Boolean = False
Private Const lStatistics As Boolean = True

Private m_AC As AC_AutoCompleteControl

Private m_CodeSort As clsCodeSort

Private m_CWE As clsCodeWindowEnhance

Private m_CodeQuick As clsCodeQuick

Private m_Statistics As clsCodeStatistics

Friend Property Get AC() As AC_AutoCompleteControl
    Set AC = m_AC
End Property

Friend Property Get CodeSort() As clsCodeSort
    Set CodeSort = m_CodeSort
End Property

Friend Property Get CWE() As clsCodeWindowEnhance
    Set CWE = m_CWE
End Property

Friend Property Get CodeQuick() As clsCodeQuick
    Set CodeQuick = CodeQuick
End Property

Friend Property Get Statistics() As clsCodeStatistics
    Set Statistics = m_Statistics
End Property


Private Sub Class_Initialize()
    If lCodeQuick Then Set m_CodeQuick = New clsCodeQuick
    If lCodeSort Then Set m_CodeSort = New clsCodeSort
    If lCWE Then Set m_CWE = New clsCodeWindowEnhance
    If lAC Then Set m_AC = New AC_AutoCompleteControl
    If lStatistics Then Set m_Statistics = New clsCodeStatistics
End Sub
