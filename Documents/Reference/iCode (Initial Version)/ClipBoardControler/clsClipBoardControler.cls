VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClipBoardControler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim sText As String
Dim dData As IPictureDisp

Dim bText As Boolean, bData As Boolean

Public Sub GetClipBoard()
    With Clipboard
        If .GetFormat(vbCFText) = True Then sText = .GetText: bText = True
        If .GetFormat(vbCFLink) Or .GetFormat(vbCFBitmap) Or .GetFormat(vbCFMetafile) Or .GetFormat(vbCFDIB) Or .GetFormat(vbCFPalette) Then Set dData = .GetData: bData = True
    End With
End Sub

Public Sub SetClipBoard()
    If bText = True Then Clipboard.SetText sText
    If bData = True Then Clipboard.SetData dData
End Sub

