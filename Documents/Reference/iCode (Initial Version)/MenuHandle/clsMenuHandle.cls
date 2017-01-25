VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private RootMenu As ME_Menu

Private Sub LoadMenu(ByVal filename As String)
    Dim k As Long
    k = FreeFile
    Open filename For Input As #k
    
    Dim l As String
    Dim TargetMenu As ME_Menu
    
    Set TargetMenu = RootMenu
    
    Do Until EOF(k) = True
        Line Input #k, l
        
        l = Trim(l)
        
        If LeftIs(l, "Begin VB.Menu ") Then
            Set TargetMenu = TargetMenu.AddChildMenu(Mid(l, Len("Begin VB.Menu ") + 1, Len(l) - Len("Begin VB.Menu ")))
        Else
            If Not (TargetMenu Is RootMenu) Then
                
                If LeftIs(l, "Caption ") Then
                    TargetMenu.Caption = Mid(l, InStr(1, l, Chr(34)) + 1, InStr(InStr(1, l, Chr(34)) + 1, l, Chr(34)) - InStr(1, l, Chr(34)) - 1)
                ElseIf LeftIs(l, "Checked ") Then
                    TargetMenu.Checked = CBool(Trim(Mid(l, InStr(1, l, "=") + 1, InStr(1, l, "'") - (InStr(1, l, "=") + 1))))
                ElseIf LeftIs(l, "Enabled ") Then
                    TargetMenu.Enabled = CBool(Trim(Mid(l, InStr(1, l, "=") + 1, InStr(1, l, "'") - (InStr(1, l, "=") + 1))))
                ElseIf LeftIs(l, "Visible ") Then
                    TargetMenu.Visible = CBool(Trim(Mid(l, InStr(1, l, "=") + 1, InStr(1, l, "'") - (InStr(1, l, "=") + 1))))
                ElseIf LeftIs(l, "Index ") Then
                    TargetMenu.Index = CLng(Trim(Mid(l, InStr(1, l, "=") + 1, Len(l) - InStr(1, l, "="))))
                ElseIf LeftIs(l, "HelpContextID ") Then
                    TargetMenu.HelpContextID = CLng(Trim(Mid(l, InStr(1, l, "=") + 1, Len(l) - InStr(1, l, "="))))
                ElseIf LeftIs(l, "NegotiatePosition") Then
                    TargetMenu.NegotiatePosition = CLng(Trim(Mid(l, InStr(1, l, "=") + 1, InStr(1, l, "'") - (InStr(1, l, "=") + 1))))
                ElseIf LeftIs(l, "WindowList") Then
                    TargetMenu.WindowList = CBool(Trim(Mid(l, InStr(1, l, "=") + 1, InStr(1, l, "'") - (InStr(1, l, "=") + 1))))
                ElseIf LeftIs(l, "ShortCut") Then
                    TargetMenu.ShortCut = TranslateShortCut(Trim(Mid(l, InStr(1, l, "=") + 1, Len(l) - InStr(1, l, "="))))
                ElseIf LeftIs(l, "End") Then
                    Set TargetMenu = TargetMenu.Parent
                    
                ElseIf LeftIs(l, "Attribute ") Then
                    Exit Do
                End If
            End If
        End If
    Loop
    
    Close #k
End Sub

Private Function TranslateShortCut(ByVal Str As String) As String
    If LeftIs(Str, "^") Then
        TranslateShortCut = "Ctrl+" & Right(Str, 1)
    ElseIf LeftIs(Str, "Ctrl+") Then
        TranslateShortCut = "^" & Right(Str, 1)
    End If
End Function

