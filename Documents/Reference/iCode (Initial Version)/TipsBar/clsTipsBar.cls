VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipsBarControler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private hTipsBar As Long

Private WithEvents m_TipsBar As TipsBar
Attribute m_TipsBar.VB_VarHelpID = -1

Public Enum enmTipsBarMode
    tbNormal = 0
    tbCollection = 1
End Enum

Public Mode As enmTipsBarMode

Private WithEvents ComponentsEvents As VBComponentsEvents
Attribute ComponentsEvents.VB_VarHelpID = -1

Dim WithEvents TB_Timer As XTimer
Attribute TB_Timer.VB_VarHelpID = -1

Private IDEWorkArea As typControlArea

Private Const TimeToCheckTipsBar As Long = 1000


Public Sub DealMessage(ByVal Msg As Long, ByVal hWnd As Long, ByVal Caption As String, ByVal ClassName As String)
    Select Case Msg
        
        
    Case HCBT_MOVESIZE, HCBT_MINMAX
        
        
        If hWnd = hVBIDE Then
            Me.UpdataTipsBarRightNow
        End If
        
        
    Case HCBT_SETFOCUS
        
        If hWnd = hVBIDE Then
            Me.UpdataTipsBarRightNow
            Exit Sub
        End If
        
        Dim tTip As Long, tChild As Long
        
        Select Case ClassName
            
        Case "VbaWindow"
            
            If Caption = "对象浏览器" Or Caption = "立即" Then Exit Sub
            
            If Me.Mode = tbNormal Then
                tTip = TipsBar.ReutrnIndexByKey(hWnd)
                If tTip = 0 Then
                    tTip = TipsBar.CreateTip(DealCaption(Caption) & " [Code]", hWnd, True)
                End If
                
                TipsBar.ActivateTip tTip, 0
                
            Else
                tTip = Me.TipsBar.ReutrnIndexByKey(DealCaption(Caption, True))
                If tTip = 0 Then
                    tTip = Me.TipsBar.CreateTip(DealCaption(Caption), DealCaption(Caption, True), True)
                    tChild = Me.TipsBar.Tips(tTip).CreateChildTip("Code", hWnd, True)
                Else
                    tChild = Me.TipsBar.Tips(tTip).ReutrnIndexByKey(hWnd)
                    If tChild = 0 Then
                        tChild = Me.TipsBar.Tips(tTip).CreateChildTip("Code", hWnd, True)
                    End If
                End If
                
                Me.TipsBar.ActivateTip tTip, tChild
            End If
            
            Me.TipsBar.DrawTips
            
            
        Case "DesignerWindow"
            
            If Me.Mode = tbNormal Then
                tTip = Me.TipsBar.ReutrnIndexByKey(hWnd)
                If tTip = 0 Then
                    tTip = Me.TipsBar.CreateTip(DealCaption(Caption) & " [Design]", hWnd, True)
                End If
                
                Me.TipsBar.ActivateTip tTip, 0
                
            Else
                tTip = Me.TipsBar.ReutrnIndexByKey(DealCaption(Caption, True))
                If tTip = 0 Then
                    tTip = Me.TipsBar.CreateTip(DealCaption(Caption), DealCaption(Caption, True), True)
                    tChild = Me.TipsBar.Tips(tTip).CreateChildTip("Design", hWnd, True)
                Else
                    tChild = Me.TipsBar.Tips(tTip).ReutrnIndexByKey(hWnd)
                    
                    If tChild = 0 Then
                        tChild = Me.TipsBar.Tips(tTip).CreateChildTip("Design", hWnd, True)
                    End If
                    
                End If
                
                Me.TipsBar.ActivateTip tTip, tChild
                
            End If
            
            Me.TipsBar.DrawTips
            
            
        Case "DockingView"
            
            If Caption = "对象浏览器" Then
                
                tTip = Me.TipsBar.ReutrnIndexByKey(hWnd)
                
                If tTip = 0 Then
                    tTip = Me.TipsBar.CreateTip("对象浏览器", hWnd, True)
                End If
                
                Me.TipsBar.ActivateTip tTip, 0
                
                Me.TipsBar.DrawTips
                
            End If
            
        End Select
        
    Case HCBT_DESTROYWND
        
        Select Case ClassName
            
        Case "VbaWindow"
            
            If Caption = "对象浏览器" Or Caption = "立即" Then Exit Sub
            
            If Me.Mode = tbNormal Then
                Me.TipsBar.DeleteTip hWnd
            Else
                Me.TipsBar.DeleteTip DealCaption(Caption, True), hWnd
            End If
            Me.TipsBar.DrawTips
            
        Case "DesignerWindow"
            
            If Me.Mode = tbNormal Then
                Me.TipsBar.DeleteTip hWnd
            Else
                Me.TipsBar.DeleteTip DealCaption(Caption, True), hWnd
            End If
            Me.TipsBar.DrawTips
            
        Case "DockingView"
            
            If Caption = "对象浏览器" Then
                Me.TipsBar.DeleteTip hWnd
                Me.TipsBar.DrawTips
            End If
        End Select
    End Select
    
End Sub

Public Function DealCaption(sCaption As String, Optional bSpecial As Boolean = False) As String
    If InStr(1, sCaption, " - ") <> 0 And bSpecial = False Then
        DealCaption = Mid(sCaption, InStr(1, sCaption, " - ") + 3, InStrRev(sCaption, " ") - (InStr(1, sCaption, " - ") + 3))
    Else
        DealCaption = Left$(sCaption, InStrRev(sCaption, " ") - 1)
    End If
End Function


Public Sub UnLoad()
    SetParent hTipsBar, frmHidden.hWnd
    Set m_TipsBar = Nothing
    TB_Timer.Pause
    Set TB_Timer = Nothing
End Sub

Public Property Get hIDEWorkArea() As Long
    hIDEWorkArea = IDEWorkArea.hWnd
End Property

Public Sub UpdataWorkArea()
    Dim rRect As RECT
    
    rRect = GetControlRect(IDEWorkArea.hWnd, hVBIDE)
    
    If Not (IDEWorkArea.Top = rRect.Top And _
        IDEWorkArea.Width = rRect.Right - rRect.Left And _
        IDEWorkArea.Height = rRect.Bottom - rRect.Top And _
        IDEWorkArea.Left = rRect.Left) Then
        
        IDEWorkArea.Top = rRect.Top
        IDEWorkArea.Width = rRect.Right - rRect.Left
        IDEWorkArea.Height = rRect.Bottom - rRect.Top
        IDEWorkArea.Left = rRect.Left
        
        IDEWorkArea.bFixed = False
        
        MoveTipsBar
        
    End If
End Sub

Public Sub UpdataTipsBarRightNow()
    TB_Timer.Interval = 50
End Sub

Friend Property Get TipsBar() As TipsBar
    Set TipsBar = m_TipsBar
End Property

Friend Property Set TipsBar(NewValue As TipsBar)
    Set m_TipsBar = NewValue
End Property

Public Sub Init()
    Set ComponentsEvents = VBIns.Events.VBComponentsEvents(Nothing)
    
    Set Me.TipsBar = frmHidden.TipsBar
    
    hTipsBar = frmHidden.TipsBar.hWnd
    
    SetParent hTipsBar, hVBIDE
    
    IDEWorkArea.hWnd = FindWindowEx(hVBIDE, 0, "MDIClient", "")
    
    DBPrint "IDEWorkArea.hWnd = " & IDEWorkArea.hWnd
    
    Set TB_Timer = New XTimer
    TB_Timer.Interval = TimeToCheckTipsBar
    TB_Timer.Start
    
    iTipsBar.UpdataWorkArea
End Sub

Public Sub MoveTipsBar()
    If IDEWorkArea.bFixed = False Then
        MoveWindow hTipsBar, IDEWorkArea.Left, IDEWorkArea.Top, IDEWorkArea.Width, frmHidden.TipsBar.Height, True
        MoveWindow IDEWorkArea.hWnd, IDEWorkArea.Left, IDEWorkArea.Top + frmHidden.TipsBar.Height, IDEWorkArea.Width, IDEWorkArea.Height - frmHidden.TipsBar.Height, True
        IDEWorkArea.Top = IDEWorkArea.Top + frmHidden.TipsBar.Height
        IDEWorkArea.Height = IDEWorkArea.Height - frmHidden.TipsBar.Height
        IDEWorkArea.bFixed = True
    Else
        MoveWindow hTipsBar, IDEWorkArea.Left, IDEWorkArea.Top - frmHidden.TipsBar.Height, IDEWorkArea.Width, frmHidden.TipsBar.Height, True
        MoveWindow IDEWorkArea.hWnd, IDEWorkArea.Left, IDEWorkArea.Top, IDEWorkArea.Width, IDEWorkArea.Height, True
    End If
    TipsBar.DrawTips True
End Sub

Private Sub ComponentsEvents_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
    MoveTipsBar
End Sub

Private Sub m_TipsBar_DblClick()
    iTipsBar.TipsBar.DrawTips
End Sub

Private Sub m_TipsBar_TipMouseUp(ByVal nTip As Variant, ByVal nChild As Variant, bActivate As Boolean, ByVal nButton As Long, ByVal nShift As Long, ByVal nX As Single, ByVal nY As Single)
    If nButton = 1 Then
        If nChild = 0 And Mode = tbNormal Then
            SetFocus CLng(TipsBar.Tips(nTip).Key)
        ElseIf nChild <> 0 And Mode = tbCollection Then
            SetFocus CLng(TipsBar.Tips(nTip).ChildTips(nChild).Key)
        End If
    ElseIf nButton = 2 Then
        frmHidden.TB_TipsIndex = nTip
        frmHidden.PopupMenu frmHidden.mnuTipsBar
    End If
    bActivate = False
End Sub


Public Sub LoadWindows()
    Dim i As Long
    For i = 1 To VBIns.Windows.Count
        With VBIns.Windows.item(i)
            If .type = vbext_wt_Designer Then
                Me.TipsBar.CreateTip LW_DealCaption(.Caption) & " [Design]", .hWnd, True
            ElseIf VBIns.Windows.item(i).type = vbext_wt_CodeWindow Then
                Me.TipsBar.CreateTip LW_DealCaption(.Caption) & " [Code]", .hWnd, True
            End If
        End With
    Next
End Sub

Private Function LW_DealCaption(sCaption As String) As String
    LW_DealCaption = Left(sCaption, InStr(1, sCaption, " (") - 1)
End Function

Public Sub DelectLeft(nTip As Long)
    Dim i As Long
    For i = 1 To nTip - 1
        SendMessage TipsBar.Tips(1).Key, WM_CLOSE, 0, 0
    Next
End Sub

Public Sub DelectRight(nTip As Long)
    Dim i As Long
    For i = nTip + 1 To iTipsBar.TipsBar.TipsTotal
        SendMessage TipsBar.Tips(nTip + 1).Key, WM_CLOSE, 0, 0
    Next
End Sub

Private Sub m_TipsBar_TipXClick(ByVal nTip As Variant, ByVal nChild As Variant, bDelete As Boolean)
    If nChild = 0 And Mode = tbNormal Then
        SendMessage CLng(TipsBar.Tips(nTip).Key), WM_CLOSE, 0, 0
    ElseIf nChild <> 0 And Mode = tbCollection Then
        SendMessage CLng(TipsBar.Tips(nTip).ChildTips(nChild).Key), WM_CLOSE, 0, 0
    End If
    
    bDelete = False
End Sub

Private Sub TB_Timer_Timer()
    UpdataWorkArea
    If TB_Timer.Interval <> TimeToCheckTipsBar Then TB_Timer.Interval = TimeToCheckTipsBar
End Sub
