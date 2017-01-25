VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iProject_Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#Const HaveDebugForm = True

Dim iProject As clsiProject


Public Sub Initialize(ByVal IDEIns As VBE)

    Set VBIns = IDEIns
    hVBIDE = VBIns.MainWindow.hwnd
    
    #If HaveDebugForm Then
        Set DebugForm = New frmDebug
        DebugForm.Show
    #End If
    
    Set iProject = New clsiProject
End Sub

Public Sub Terminate()
    Set iProject = Nothing
End Sub

