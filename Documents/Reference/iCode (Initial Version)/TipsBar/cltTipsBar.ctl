VERSION 5.00
Begin VB.UserControl TipsBar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10725
   ScaleHeight     =   608
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   715
End
Attribute VB_Name = "TipsBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Pi = 3.14159265358

Private Const cR As Single = 3                                                  '圆角矩形圆弧对应的圆的半径

Private Const cDisX As Single = 6                                               'Tip标签字符与标签左右的间距
Private Const cDisY As Single = 2                                               'Tip标签字符与标签上下的间距

Private Const cXCR As Single = 3                                                '关闭按钮所在圆的半径
Private Const cXDisX As Single = 3                                              '关闭按钮与字符的间距
Private Const cXDrawWidth As Long = 2                                           '关闭按钮的线条粗细

Private Const cChildTipsTextDis As Long = 5                                     '子标签与父标签文字的间距


Private Const cCTDisX As Long = 5
Private Const cCTDisY As Single = 1.5
Private Const cCTXCR As Single = 2                                              '关闭按钮所在圆的半径
Private Const cCTXDisX As Single = 3                                            '关闭按钮与字符的间距
Private Const cCTXDrawWidth As Long = 2                                         '关闭按钮的线条粗细
Private Const cCTParentXDis As Long = 2

Private Const cTipsTop As Single = 1

Private Const cTextAColor As Long = vbBlack
Private Const cTextUAColor As Long = &H696969

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function lStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Const strTest As String = "中文"


Private m_Tips() As TB_Tip

Dim pTextY As Single
Dim pTipsH As Single


Dim pCTTextY As Single
Dim pCTTipsH As Single

Private Const cIconXDis As Single = 4
Private Const cIconTextXDis As Single = 3

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Dim mLastTip As String
Dim mLastX As Single, mLastY As Single


Dim nOldForeColor As Long

Public Enum enmMouseOnTipMode
    mNone
    mNormal
    mOnX
End Enum

Public Enum enmMouseAction
    maMove
    maClick
    maDblClick
    maMouseDown
    maMouseUp
End Enum

Public Enum enmTipsBarAlignMode
    tbaTop = 1
    tbaBottom = 2
    tbaLeft = 3
    tbaRight = 4
End Enum

'缺省属性值:
Const m_def_AutoBottomLine = False
Const m_def_AutoBottomLineColor = vbBlack

Private Const IconWidth = 15
Private Const IconHeight = 15

Private Const XAColor = vbRed
Private Const XUAColor = &H80000010

Private Const TipsAColor = vbWhite
Private Const TipsUAColor = &HF9F9F9

Private Const ChildTipsDis = 2

Private Const TipsDis = 0

Private Const UnEnabledColor = &H8000000A
Private Const LineColor = &HB4B4B4
Private Const FirstTipX = 3

Const m_def_OnlyOneActive = True
Const m_def_Align = 1
Private Const TipsFontSize = 9
Private Const ChildTipsFontSize = 8

'属性变量:
Dim m_nActiveTip As Single
Dim m_AutoBottomLine As Boolean
Dim m_AutoBottomLineColor As OLE_COLOR
Dim m_OnlyOneActive As Boolean
Dim m_Align As enmTipsBarAlignMode

'事件声明:

Event TipMouseDown(ByVal nTip, ByVal nChild, ByRef bActivate As Boolean, ByVal nButton As Long, ByVal nShift As Long, ByVal nX As Single, ByVal nY As Single)
Event TipMouseUp(ByVal nTip, ByVal nChild, ByRef bActivate As Boolean, ByVal nButton As Long, ByVal nShift As Long, ByVal nX As Single, ByVal nY As Single)

Event TipClick(ByVal nTip, ByVal nChild, ByRef bActivate As Boolean)
Event TipDblClick(ByVal nTip, ByVal nChild)
Event TipXClick(ByVal nTip, ByVal nChild, ByRef bDelete As Boolean)

Event MouseIn(ByVal X As Single, ByVal Y As Single)
Event MouseOut(ByVal X As Single, ByVal Y As Single)

Event TipCreate(ByRef sCaption As String, ByRef Key, _
    ByRef bCloseButton As Boolean, _
    ByRef nParent, _
    ByRef sPopUp As String, _
    ByRef bCreate As Boolean)

Event TipDelete(ByVal nTip, ByVal nChild, ByRef bDelete As Boolean, ByRef bActivateAnother As Boolean)

Event TipMouseIn(ByVal nTip, ByVal nChild, ByRef bActivate As Boolean)
Event TipMouseOut(ByVal nTip, ByVal nChild, ByRef bUnActivate As Boolean)
Event TipMouseMove(ByVal nTip, ByVal nChild, ByVal nX As Single, ByVal nY As Single)

Event TipActivate(ByVal nTip, ByVal nChild, ByRef bActivate As Boolean)
Event TipUnActivate(ByVal nTip, ByVal nChild, ByRef bUnActivate As Boolean)

Event TipXMouseIn(ByVal nTip, ByVal nChild, ByRef bXActivate As Boolean)
Event TipXMouseOut(ByVal nTip, ByVal nChild, ByRef bXUnActivate As Boolean)

Event Click()                                                                   'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick()                                                                'MappingInfo=UserControl,UserControl,-1,DblClick
Event Hide()                                                                    'MappingInfo=UserControl,UserControl,-1,Hide
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)  'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)  'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)    'MappingInfo=UserControl,UserControl,-1,MouseUp
    
    Friend Property Get ActiveTip() As TB_Tip
    If Int(nActiveTip) = nActiveTip Then
        Set ActiveTip = Tips(nActiveTip)
    Else
        Set ActiveTip = Tips(Int(nActiveTip)).ChildTips(GetChildTipIndexFromStr(nActiveTip))
    End If
End Property

Public Property Get TipsTotal() As Long
    TipsTotal = UBound(m_Tips)
End Property

Public Sub Clear()
    Erase m_Tips
    ReDim m_Tips(0)
End Sub

'Property Get Tips
'输入: ByVal n —— Index或Key
'返回值: TB_Tip或Nothing

Public Property Get Tips(ByVal n) As TB_Tip
    If KeyToIndex(n) = True Then
        Set Tips = m_Tips(n)
    Else
        Set Tips = Nothing
    End If
End Property

'Property Get Tips
'输入: ByVal n —— Index或Key
'      Byval New_TB_Tip as TB_Tip —— 新Tip
'返回值: 无
'错误： 如找不到对应的Tip,则会引发错误3611

Public Property Set Tips(ByVal n, ByVal New_TB_Tip As TB_Tip)
    If KeyToIndex(n) = True Then
        Set m_Tips(n) = New_TB_Tip
    Else
        Err.Raise "3611", "TipsBar Control - Property Set Tips", "找不到 Tip ！"
    End If
End Property


Private Sub UserControl_DblClick()
    RaiseEvent DblClick
    mDealAction mLastX, mLastY, mLastTip, maDblClick
End Sub

Private Sub UserControl_Initialize()
    ReDim m_Tips(0)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If X > 0 And X < UserControl.ScaleWidth And Y > 0 And Y < UserControl.ScaleHeight Then
        
        
        mLastX = X
        mLastY = Y
        
        SetCapture UserControl.hWnd
        
        Dim sRes As String
        If mLastTip = -1 Or mLastTip = 0 Then
            If mLastTip = -1 Then RaiseEvent MouseIn(X, Y)
            RaiseEvent MouseMove(Button, Shift, X, Y)
            mLastTip = mDealAction(X, Y)
        Else
            RaiseEvent MouseMove(Button, Shift, X, Y)
            mLastTip = mDealAction(X, Y, mLastTip)
        End If
        
        
    Else
        If mLastTip <> -1 Then
            
            If mLastTip <> 0 Then
                CallEventTipMouseOut Int(mLastTip), GetChildTipIndexFromStr(mLastTip)
                
                If GetChildTipIndexFromStr(mLastTip) > 0 Then
                    CallEventTipMouseOut Int(mLastTip), 0
                    Tips(mLastTip).ChildTips(GetChildTipIndexFromStr(mLastTip)).bMouseOnX = False
                End If
                
                Tips(mLastTip).bMouseOnX = False
            End If
            
            RaiseEvent MouseOut(X, Y)
            
            ReleaseCapture
            mLastTip = -1
            
        End If
    End If
End Sub

Public Sub CallEventTipMouseDown(ByVal nTip, ByVal nChild, nButton, nShift, nX, nY)
    Dim bRes As Boolean: bRes = False
    RaiseEvent TipMouseDown(nTip, nChild, bRes, nButton, nShift, nX, nY)
    If bRes = True Then ActivateTip nTip, nChild
End Sub

Public Sub CallEventTipMouseUp(ByVal nTip, ByVal nChild, nButton, nShift, nX, nY)
    Dim bRes As Boolean: bRes = False
    RaiseEvent TipMouseUp(nTip, nChild, bRes, nButton, nShift, nX, nY)
    If bRes = True Then UnActivateTip nTip, nChild
End Sub

Private Sub CallEventTipMouseIn(ByVal nTip, Optional ByVal nChild = 0)
    Dim bRes As Boolean: bRes = False
    RaiseEvent TipMouseIn(nTip, nChild, bRes)
    If bRes = True Then ActivateTip nTip, nChild
End Sub

Private Sub CallEventTipMouseOut(ByVal nTip, Optional ByVal nChild = 0)
    Dim bRes As Boolean: bRes = False
    RaiseEvent TipMouseOut(nTip, nChild, bRes)
    If bRes = True Then UnActivateTip nTip, nChild
End Sub

Private Function CallEventTipActivate(ByRef nTip, Optional ByRef nChild = 0) As Boolean
    Dim bRes As Boolean: bRes = True
    RaiseEvent TipActivate(nTip, nChild, bRes)
    CallEventTipActivate = bRes
End Function

'Function ActivateTip
'输入: ByVal n —— Tip的Index或Key
'      Optional ByVal nChild = 0 —— ChildTip的Index或Key
'返回值: Bool —— True表示成功，False表示失败

Public Function ActivateTip(ByVal n, Optional ByVal nChild = 0) As Boolean
    ActivateTip = False
    
    On Error GoTo iErr
    
    
    If KeyToIndex(n) = False Then Exit Function
    
    If nChild <> 0 Then If Tips(n).KeyToIndex(nChild) = False Then Exit Function
    
    
    If CallEventTipActivate(n, nChild) = False Then Exit Function
    
    If nChild = 0 Then
        
        If Int(nActiveTip) = n Then ActivateTip = True: Exit Function
        
        Tips(n).Active = True
        
        If Tips(n).ChildTipsTotal > 0 Then
            ActivateTip n, 1
        Else
            If Me.OnlyOneActive = True Then UnActivateTip Int(nActiveTip), Me.GetChildTipIndexFromStr(nActiveTip)
            nActiveTip = n
        End If
        
    Else
        
        If nActiveTip = n & "." & nChild Then ActivateTip = True: Exit Function
        
        Tips(n).ChildTips(nChild).Active = True
        
        If Me.OnlyOneActive = True Then UnActivateTip Int(nActiveTip), Me.GetChildTipIndexFromStr(nActiveTip)
        
        nActiveTip = n & "." & nChild
        
        If Tips(n).Active = False Then Tips(n).Active = True
        
    End If
    
    ActivateTip = True
    
    Exit Function
iErr:
    DBErr "TipsBar - Function ActivateTip", "n = " & n, "nChild = " & nChild
End Function

'KeyToIndex
'输入:ByRef n —— Key或Index → Index
'返回值: Bool —— True表示成功,False表示失败

Public Function KeyToIndex(ByRef n) As Boolean
    If IsNumeric(n) And Val(n) <= Me.TipsTotal Then
        KeyToIndex = True
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To UBound(m_Tips)
        If Tips(i).Key = n Then
            n = i
            KeyToIndex = True
            Exit Function
        End If
    Next
End Function

'ReutrnIndexByKey
'输入:ByVal n —— Key或Index
'返回值: Lond —— Index,0表示失败

Public Function ReutrnIndexByKey(ByVal n) As Long
    If IsNumeric(n) And Val(n) <= Me.TipsTotal Then
        ReutrnIndexByKey = n
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To Me.TipsTotal
        If Tips(i).Key = n Then
            ReutrnIndexByKey = i
            Exit Function
        End If
    Next
End Function

'GetIndexBySearchKey
'输入:ByVal n —— Key或Index
'ByRef nTip —— →找到的Tip的Index
'ByRef nChild  —— →找到的ChildTip的Index
'返回值: Bool —— True表示成功,False表示失败

Public Function GetIndexBySearchKey(ByVal n, ByRef nTip, ByRef nChild) As Boolean
    Dim i As Long
    For i = 1 To UBound(m_Tips)
        If Tips(i).Key = n Then
            nTip = i
            nChild = 0
            GetIndexBySearchKey = True
            Exit Function
        Else
            Dim j As Long
            For j = 1 To Tips(i).ChildTipsTotal
                If Tips(i).ChildTips(j).Key = n Then
                    nTip = i
                    nChild = j
                    GetIndexBySearchKey = True
                    Exit Function
                End If
            Next
        End If
    Next
End Function


Private Function CallEventUnActivateTip(ByRef n, Optional ByRef nChild = 0) As Boolean
    Dim bRes As Boolean: bRes = True
    RaiseEvent TipUnActivate(n, nChild, bRes)
    CallEventUnActivateTip = bRes
End Function

'Function UnActivateTip
'输入: ByVal n —— Tip的Index或Key
'      Optional ByVal nChild = 0 —— ChildTip的Index或Key
'返回值: Bool —— True表示成功，False表示失败
'注意: 此函数仅用于反激活标签，不考虑接下来的跳转之类的问题

Public Function UnActivateTip(ByVal n, Optional ByVal nChild = 0) As Boolean
    UnActivateTip = False
    
    If n = 0 Then Exit Function                                                 '为了尚未有激活标签时调用此函数而不出错
    
    
    If KeyToIndex(n) = False Then Exit Function
    If nChild <> 0 Then If Tips(n).KeyToIndex(nChild) = False Then Exit Function
    
    If CallEventUnActivateTip(n, nChild) = False Then Exit Function
    
    If nChild = 0 Then
        Tips(n).Active = False
    Else
        Tips(n).ChildTips(nChild).Active = False
        
        If Me.OnlyOneActive = True Then Tips(n).Active = False
    End If
    
    
    UnActivateTip = True
End Function

Private Sub CallEventTipXMouseIn(ByVal nTip, Optional ByVal nChild = 0)
    Dim bRes As Boolean: bRes = True
    RaiseEvent TipXMouseIn(nTip, nChild, bRes)
    
    If bRes = False Then Exit Sub
    
    '调用此函数前已进行关闭按钮是否已激活的判断
    
    If nChild = 0 Then
        Tips(nTip).bMouseOnX = True
        DrawTip Tips(nTip)
    Else
        Tips(nTip).ChildTips(nChild).bMouseOnX = True
        DrawTip Tips(nTip).ChildTips(nChild)
    End If
    
    UserControl.Refresh
End Sub

Private Sub CallEventTipXMouseOut(ByVal nTip, Optional ByVal nChild = 0)
    Dim bRes As Boolean: bRes = True
    RaiseEvent TipXMouseOut(nTip, nChild, bRes)
    
    If bRes = False Then Exit Sub
    
    '调用此函数前已进行关闭按钮是否未激活的判断
    
    If nChild = 0 Then
        Tips(nTip).bMouseOnX = False
        DrawTip Tips(nTip)
    Else
        Tips(nTip).ChildTips(nChild).bMouseOnX = False
        DrawTip Tips(nTip).ChildTips(nChild)
    End If
    
    UserControl.Refresh
End Sub

Public Function mMouseOnTip(ByVal nX As Single, ByVal nY As Single, tTip As TB_Tip) As enmMouseOnTipMode
    mMouseOnTip = mNone
    
    Select Case Me.Align
        
    Case tbaTop
        If (nX > tTip.nX And nX < tTip.nX + tTip.nWidth) = False Then Exit Function
        
        If tTip.ParentTip = 0 Then
            If ((UserControl.ScaleHeight - nY) > 0 And (UserControl.ScaleHeight - nY) < pTipsH) = False Then Exit Function
        Else
            If ((UserControl.ScaleHeight - nY) > 0 And (UserControl.ScaleHeight - nY) < pCTTipsH) = False Then Exit Function
        End If
        
        mMouseOnTip = mNormal
        
        If (nX > tTip.nXX - 2) = False Then Exit Function
        
        If tTip.ParentTip = 0 Then
            If (nX < (tTip.nXX + cXCR * 2 + 2)) = False Then Exit Function
            If ((UserControl.ScaleHeight - nY) > cDisY And (UserControl.ScaleHeight - nY) < (cDisY + 2 * cXCR + 4)) = False Then Exit Function
        Else
            If (nX < tTip.nXX + cCTXCR * 2 + 3) = False Then Exit Function
            If ((UserControl.ScaleHeight - nY) > cCTDisY And (UserControl.ScaleHeight - nY) < (cCTDisY + 3 * cCTXCR + 4)) = False Then Exit Function
        End If
        
        mMouseOnTip = mOnX
        
    Case tbaBottom
        If (nX > tTip.nX And nX < tTip.nX + tTip.nWidth) = False Then Exit Function
        
        If tTip.ParentTip = 0 Then
            If (nY > 0 And nY < pTipsH) = False Then Exit Function
        Else
            If (nY > 0 And nY < pCTTipsH) = False Then Exit Function
        End If
        
        mMouseOnTip = mNormal
        
        If (nX > tTip.nXX - 2) = False Then Exit Function
        
        If tTip.ParentTip = 0 Then
            If (nX < (tTip.nXX + cXCR * 2 + 2)) = False Then Exit Function
            If (nY > cDisY And nY < (cDisY + 2 * cXCR + 4)) = False Then Exit Function
        Else
            If (nX < tTip.nXX + cCTXCR * 2 + 3) = False Then Exit Function
            If (nY > cCTDisY And nY < (cCTDisY + 3 * cCTXCR + 4)) = False Then Exit Function
        End If
        
        mMouseOnTip = mOnX
        
    Case tbaRight
        
        If (nY > tTip.nX And nY < tTip.nX + tTip.nWidth) = False Then Exit Function
        
        If tTip.ParentTip = 0 Then
            If (nX > 0 And nX < pTipsH) = False Then Exit Function
        Else
            If (nX > 0 And nX < pCTTipsH) = False Then Exit Function
        End If
        
        mMouseOnTip = mNormal
        
        If (nY > tTip.nXX - 2) = False Then Exit Function
        
        If tTip.ParentTip = 0 Then
            If (nY < (tTip.nXX + cXCR * 2 + 2)) = False Then Exit Function
            If (nX > cDisY And nX < (cDisY + 2 * cXCR + 4)) = False Then Exit Function
        Else
            If (nY < tTip.nXX + cCTXCR * 2 + 3) = False Then Exit Function
            If (nX > cCTDisY And nX < (cCTDisY + 3 * cCTXCR + 4)) = False Then Exit Function
        End If
        
        mMouseOnTip = mOnX
        
    Case tbaLeft
        
        If (nY > tTip.nX And nY < tTip.nX + tTip.nWidth) = False Then Exit Function
        
        If tTip.ParentTip = 0 Then
            If ((UserControl.ScaleWidth - nX) > 0 And (UserControl.ScaleWidth - nX) < pTipsH) = False Then Exit Function
        Else
            If ((UserControl.ScaleWidth - nX) > 0 And (UserControl.ScaleWidth - nX) < pCTTipsH) = False Then Exit Function
        End If
        
        mMouseOnTip = mNormal
        
        If (nY > tTip.nXX - 2) = False Then Exit Function
        
        If tTip.ParentTip = 0 Then
            If (nY < (tTip.nXX + cXCR * 2 + 2)) = False Then Exit Function
            If ((UserControl.ScaleWidth - nX) > cDisY And (UserControl.ScaleWidth - nX) < (cDisY + 2 * cXCR + 4)) = False Then Exit Function
        Else
            If (nY < tTip.nXX + cCTXCR * 2 + 3) = False Then Exit Function
            If ((UserControl.ScaleWidth - nX) > cCTDisY And (UserControl.ScaleWidth - nX) < (cCTDisY + 3 * cCTXCR + 4)) = False Then Exit Function
        End If
        
        mMouseOnTip = mOnX
        
    End Select
    
End Function

Public Function GetChildTipIndexFromStr(sStr) As Long
    If InStr(1, sStr, ".") <> 0 Then
        GetChildTipIndexFromStr = Right(sStr, Len(sStr) - InStr(1, sStr, "."))
    End If
End Function

Public Function mDealAction(ByVal nX As Single, ByVal nY As Single, Optional ByVal nStart As Long = 0, _
    Optional ByVal nMouseAction As enmMouseAction = maMove, _
    Optional eCode1, Optional eCode2) As String
    Dim bRes As Boolean
    
    On Error GoTo iErr
    
    Dim i As Long
    
    
    For i = IIf(Int(nStart) > 1, Int(nStart) - 1, 1) To IIf(Int(nStart) > 0 And Int(nStart) < UBound(m_Tips), Int(nStart) + 1, UBound(m_Tips))
        Select Case mMouseOnTip(nX, nY, Tips(i))
        Case mOnX
            
            
            
            '激活OnX必先经过Normal状态
            '则不会出现由一个标签直接跳到另一个标签的关闭按钮上
            '故无需激活TipMouseOut等事件
            
            mDealAction = i
            
            
            'Debug.Print "mOnX!"
            
            Select Case nMouseAction
            Case maMove                                                         '如果不是maMove事件,不会产生MouseIn/Out事件
                If nStart = 0 Or nStart = -1 Then
                    CallEventTipMouseIn i, 0
                ElseIf Int(nStart) <> i Then
                    CallEventTipMouseOut Int(nStart), GetChildTipIndexFromStr(nStart)
                    CallEventTipMouseIn i, 0
                End If
                
                If Tips(i).bMouseOnX = False Then CallEventTipXMouseIn i, 0
                RaiseEvent TipMouseMove(i, 0, nX, nY)
            Case maDblClick
                RaiseEvent TipDblClick(i, 0)
            Case maMouseDown
                CallEventTipMouseDown i, 0, eCode1, eCode2, nX, nY
            Case maMouseUp
                CallEventTipMouseUp i, 0, eCode1, eCode2, nX, nY
                CallEventTipXClick i, 0
            End Select
            
            
            Exit Function
            
        Case mNormal
            
            mDealAction = i
            
            
            If nMouseAction = maMove Then
                If Int(nStart) <> i Then
                    If Int(nStart) <> 0 And Int(nStart) <> -1 Then
                        CallEventTipMouseOut Int(nStart), GetChildTipIndexFromStr(nStart)
                        CallEventTipXMouseOut Int(nStart), 0
                    End If
                End If
                
                If Tips(i).bMouseOnX = True Then CallEventTipXMouseOut i, 0
                
            End If
            
            If Tips(i).ChildTipsTotal > 0 Then
                
                Dim j As Long
                
                For j = 1 To Tips(i).ChildTipsTotal
                    
                    Select Case mMouseOnTip(nX, nY, Tips(i).ChildTips(j))
                        
                    Case mOnX
                        
                        mDealAction = mDealAction & "." & j
                        
                        Select Case nMouseAction
                        Case maMove
                            If nStart = 0 Or nStart = -1 Then
                                CallEventTipMouseIn i, j
                            ElseIf nStart <> mDealAction Then
                                '此处不可能由一个Tip直接跳到另一个ChildTip
                                '故无需考虑TipMouseOut重复激活
                                '此处只针对同一个ParentTip中鼠标由一个ChildTip移动至另一个
                                CallEventTipMouseOut Int(nStart), GetChildTipIndexFromStr(nStart)
                                CallEventTipMouseIn i, j
                            End If
                            
                            If Tips(i).ChildTips(j).bMouseOnX = False Then CallEventTipXMouseIn i, j
                            
                            'RaiseEvent TipMouseMove(i, j, nX, nY)
                        Case maClick
                            '在关闭按钮上Click时只需激活TipXClick而无需激活TipClick
                            CallEventTipXClick i, j
                        Case maDblClick
                            RaiseEvent TipDblClick(i, j)
                        Case maMouseDown
                            CallEventTipMouseDown i, j, eCode1, eCode2, nX, nY
                        Case maMouseUp
                            CallEventTipMouseUp i, j, eCode1, eCode2, nX, nY
                            'CallEventTipClick i, j
                            
                        End Select
                        
                        Exit Function
                        
                    Case mNormal
                        
                        mDealAction = mDealAction & "." & j
                        Select Case nMouseAction
                            
                        Case maMove
                            
                            If nStart = 0 Or nStart = -1 Then
                                CallEventTipMouseIn i, 0
                                CallEventTipMouseIn i, j
                            ElseIf nStart <> mDealAction Then
                                If InStr(1, nStart, ".") <> 0 Then CallEventTipMouseOut Int(nStart), GetChildTipIndexFromStr(nStart)
                                CallEventTipMouseIn i, j
                            End If
                            
                            If Tips(i).ChildTips(j).bMouseOnX = True Then CallEventTipXMouseOut i, j
                            
                            
                            RaiseEvent TipMouseMove(i, j, nX, nY)
                            
                            'Case maClick
                            
                        Case maDblClick
                            RaiseEvent TipDblClick(i, j)
                        Case maMouseDown
                            CallEventTipMouseDown i, j, eCode1, eCode2, nX, nY
                        Case maMouseUp
                            CallEventTipMouseUp i, j, eCode1, eCode2, nX, nY
                            CallEventTipClick i, j
                        End Select
                        
                        
                    End Select
                    
                    
                    
                Next
                
                If InStr(1, mDealAction, ".") = 0 Then
                    If InStr(1, nStart, ".") <> 0 Then
                        CallEventTipMouseOut Int(nStart), GetChildTipIndexFromStr(nStart)
                    End If
                End If
                
            End If
            
            If InStr(1, mDealAction, ".") = 0 Then                              '当Tip没有ChildTip或鼠标不在ChildTip上时
                Select Case nMouseAction
                Case maMove
                    CallEventTipMouseIn i, 0
                    
                    RaiseEvent TipMouseMove(i, 0, nX, nY)
                    
                    'Case maClick
                    
                Case maDblClick
                    RaiseEvent TipDblClick(i, 0)
                Case maMouseDown
                    CallEventTipMouseDown i, 0, eCode1, eCode2, nX, nY
                Case maMouseUp
                    CallEventTipMouseUp i, 0, eCode1, eCode2, nX, nY
                    CallEventTipClick i, 0
                End Select
                
            End If
            
            
        End Select
        
        
    Next
    
    If mDealAction = "" Then
        mDealAction = 0
        If nStart > 0 Then
            CallEventTipMouseOut Int(nStart), GetChildTipIndexFromStr(nStart)
            CallEventTipXMouseOut Int(nStart), GetChildTipIndexFromStr(nStart)
        End If
    End If
    Exit Function
iErr:
    DBErr "TipsBar - Function mDealAction", "nX = " & nX, "nY = " & nY, "nStart = " & nStart, "nMouseAction = " & nMouseAction, "eCode1 = " & eCode1, "eCode2 = " & eCode2
End Function

'Function DrawTips
'输入: Optional bNoCalc As Boolean = False —— bNoCalc=True则仅绘制而不重算
'返回值: Bool —— True表示成功，False表示失败

Public Function DrawTips(Optional bNoCalc As Boolean = False) As Boolean
    DrawTips = False
    
    UserControl.Cls
    
    If Me.TipsTotal <= 0 Then Exit Function
    
    If bNoCalc = False Then CalcTips
    
    Dim i As Long
    Dim j As Long
    
    
    For i = 1 To Me.TipsTotal
        DrawTip Tips(i)
        If Tips(i).ChildTipsTotal > 0 Then
            For j = 1 To Tips(i).ChildTipsTotal
                DrawTip Tips(i).ChildTips(j)
            Next
        End If
    Next
    
    If Me.Align = tbaLeft Or Me.Align = tbaTop Then
        If AutoBottomLine Then iLine 0, 0.5, Tips(Me.TipsTotal).nX + Tips(Me.TipsTotal).nWidth + FirstTipX, 0.5, AutoBottomLineColor
    Else
        If AutoBottomLine Then iLine 0, 0, Tips(Me.TipsTotal).nX + Tips(Me.TipsTotal).nWidth + FirstTipX, 0, AutoBottomLineColor
    End If
    
    UserControl.Refresh
    
    DrawTips = True
End Function


Private Sub CalcPublicVar()
    SetFont False, TipsFontSize
    pTextY = cDisY + iTextHeight(strTest) - 1                                   '1为Tip标题向下偏移量
    pTipsH = pTextY + cDisY
    
    Select Case Me.Align
    Case tbaTop, tbaBottom
        
        SetFont False, TipsFontSize
        pTextY = cDisY + iTextHeight(strTest) - 1                               '1为Tip标题向下偏移量
        pTipsH = pTextY + cDisY
        
        ChangeSize
        
        SetFont False, ChildTipsFontSize
        pCTTextY = cCTDisY + iTextHeight(strTest) - 0.8                         '0.8为ChildTip标题向下偏移量
        pCTTipsH = pCTTextY + cCTDisY
        
    Case tbaLeft, tbaRight
        
        SetFont False, TipsFontSize
        pTextY = cDisY + iTextHeight(strTest) - 1                               '1为Tip标题向下偏移量
        pTipsH = pTextY + cDisY
        
        ChangeSize
        
        SetFont False, ChildTipsFontSize
        pCTTextY = cCTDisY + iTextHeight(strTest)                               '0.8为ChildTip标题向下偏移量
        pCTTipsH = pCTTextY + cCTDisY
        
    End Select
    
    
End Sub

Private Sub iTextOut(ByVal nX, ByVal nY, ByVal sString As String)
    Dim i As Long, sChar As String
    
    Select Case Me.Align
        
    Case tbaTop
        TextOut UserControl.hDC, nX, UserControl.ScaleHeight - nY, sString, lStrLen(sString)
    Case tbaBottom
        TextOut UserControl.hDC, nX, nY - iTextHeight(strTest), sString, lStrLen(sString)
    Case tbaRight
        
        For i = 1 To Len(sString)
            sChar = Mid(sString, i, 1)
            TextOut UserControl.hDC, nY - iTextWidth(sChar) / 2 - iTextHeight(sChar) / 2, nX + iTextWidth(sChar) * (i - 1), sChar, lStrLen(sChar)
        Next
        
    Case tbaLeft
        
        For i = 1 To Len(sString)
            sChar = Mid(sString, i, 1)
            TextOut UserControl.hDC, UserControl.ScaleWidth - (nY - iTextWidth(sChar) / 2) - iTextHeight(sChar) / 2, nX + iTextWidth(sChar) * (i - 1), sChar, lStrLen(sChar)
        Next
        
    End Select
End Sub

Private Sub iCircle(nX, nY, nR, Optional lColor, Optional nStart, Optional nEnd, Optional nAspect)
    Select Case Me.Align
    Case tbaTop
        iCircleSub nX, UserControl.ScaleHeight - nY, nR, lColor, nStart, nEnd, nAspect
    Case tbaBottom
        iCircleSub nX, nY, nR, lColor, nStart, nEnd, nAspect
    Case tbaRight
        iCircleSub nY, nX, nR, lColor, nStart, nEnd, nAspect
    Case tbaLeft
        iCircleSub UserControl.ScaleWidth - nY, nX, nR, lColor, nStart, nEnd, nAspect
    End Select
End Sub

Private Sub iLine(nX1, nY1, nX2, nY2, Optional lColor, Optional bB As Boolean = False, Optional bF As Boolean = False)
    Select Case Me.Align
    Case tbaTop
        iLineSub nX1, UserControl.ScaleHeight - nY1, nX2, UserControl.ScaleHeight - nY2, lColor, bB, bF
    Case tbaBottom
        iLineSub nX1, nY1, nX2, nY2, lColor, bB, bF
    Case tbaRight
        iLineSub nY1, nX1, nY2, nX2, lColor, bB, bF
    Case tbaLeft
        iLineSub UserControl.ScaleWidth - nY1, nX1, UserControl.ScaleWidth - nY2, nX2, lColor, bB, bF
    End Select
End Sub


Private Function iTextWidth(ByVal sStr As String) As Single
    Select Case Me.Align
    Case tbaTop, tbaBottom
        iTextWidth = UserControl.TextWidth(sStr)
    Case tbaLeft, tbaRight
        If IsAllEnglish(sStr) Then
            iTextWidth = (UserControl.TextHeight(sStr) - 0.3) * Len(sStr)
        Else
            iTextWidth = (UserControl.TextHeight(sStr)) * Len(sStr)
        End If
    End Select
End Function

Private Function iTextHeight(ByVal sStr As String) As Single
    Select Case Me.Align
    Case tbaTop, tbaBottom
        iTextHeight = UserControl.TextHeight(sStr)
    Case tbaLeft, tbaRight
        If IsAllEnglish(sStr) Then
            iTextHeight = UserControl.TextWidth("A") - 0
        Else
            iTextHeight = UserControl.TextWidth("中")
        End If
    End Select
End Function

Private Sub CalcTips()
    
    '○以下开始所有Tips运算
    
    Dim n As Long
    
    For n = 1 To UBound(m_Tips)
        
        '○以下开始Tip横坐标运算
        
        Dim nTipX As Long
        
        If n = 1 Then                                                           '通过第一个Tip的初始横坐标值 求得当前Tip的横坐标
            nTipX = FirstTipX
        Else                                                                    '通过上一个Tip的横坐标与宽度 求得当前Tip的横坐标
            nTipX = Tips(n - 1).nX + Tips(n - 1).nWidth + TipsDis
        End If
        
        Tips(n).nX = nTipX                                                      '存储Tip横坐标
        
        nTipX = 0                                                               '初始化变量
        
        '○以上完成Tip横坐标运算
        
        '○以下开始Tip的宽度运算
        
        Dim nTipWidth As Long
        
        If Tips(n).bHaveIcon = True Then
            nTipWidth = cIconXDis
            Tips(n).nIconX = Tips(n).nX + nTipWidth
            nTipWidth = nTipWidth + IconWidth
            
            nTipWidth = nTipWidth + cIconTextXDis
        Else
            nTipWidth = nTipWidth + cDisX
        End If
        
        
        
        Tips(n).nStrX = Tips(n).nX + nTipWidth
        
        SetFont Tips(n).Active, TipsFontSize
        
        nTipWidth = nTipWidth + iTextWidth(Tips(n).Caption)                     'nTipWidth位移至字符后
        
        If Tips(n).ChildTipsTotal > 0 Then                                      '如果此Tip有ChildTips则进行运算
            
            nTipWidth = nTipWidth + cChildTipsTextDis                           'nTipWidth位移至第一个ChildTip前
            
            '○以下开始第所有ChildTips的运算
            
            Dim m As Long
            
            For m = 1 To Tips(n).ChildTipsTotal
                
                '○以下开始当前ChildTip的横坐标运算
                
                Dim nCTX As Long
                
                
                
                If m = 1 Then                                                   '通过第一个ChildTip的初始横坐标值 求得当前Tip的横坐标
                    nCTX = Tips(n).nX + nTipWidth
                Else                                                            '通过上一个ChildTip的横坐标与宽度 求得当前Tip的横坐标
                    nCTX = Tips(n).ChildTips(m - 1).nX + Tips(n).ChildTips(m - 1).nWidth + ChildTipsDis
                End If
                
                Tips(n).ChildTips(m).nX = nCTX
                
                nCTX = 0                                                        '初始化变量
                
                '○以上完成当前ChildTip的横坐标运算
                
                '○以下开始当前ChildTip的宽度运算
                
                Dim nCTWidth As Long                                            'nCTWidth从ChildTip开头开始位移
                
                nCTWidth = cCTDisX                                              'nCTWidth位移至字符前
                
                
                
                
                Tips(n).ChildTips(m).nStrX = Tips(n).ChildTips(m).nX + nCTWidth
                
                SetFont Tips(n).ChildTips(m).Active, ChildTipsFontSize
                
                nCTWidth = nCTWidth + iTextWidth(Tips(n).ChildTips(m).Caption)  'nCTWidth位移至字符后
                
                If Tips(n).ChildTips(m).ShowCloseButton = True Then
                    
                    '运算ChildTip关闭按钮
                    
                    nCTWidth = nCTWidth + cCTXDisX                              'nCTWidth位移至关闭按钮前
                    
                    Tips(n).ChildTips(m).nXX = Tips(n).ChildTips(m).nX + nCTWidth '完成当前ChildTip的关闭按钮的横坐标运算
                    
                    nCTWidth = nCTWidth + cCTXCR * 2 + 2                        'nCTWidth位移至关闭按钮后
                    '2为为关闭按钮右侧额外增加的空间
                    
                End If
                
                nCTWidth = nCTWidth + cCTDisX                                   'nCTWidth位移至ChildTip末尾
                
                Tips(n).ChildTips(m).nWidth = nCTWidth                          '存储ChildTip宽度
                
                '○以上完成当前ChildTip的宽度运算
                
                nTipWidth = nTipWidth + nCTWidth + ChildTipsDis                 'nTipWidth位移至下一个ChildTip的开头
                
                nCTWidth = 0                                                    '初始化变量
                
            Next                                                                '○以上完成所有ChildTips的横坐标和宽度运算
            
            '○以上完成所有ChildTips的运算
            
            'nTipWidth位移至所有ChildTips后
            
        End If
        
        If Tips(n).ShowCloseButton = True Then
            
            '运算Tip关闭按钮
            
            '关闭按钮前的间距因是否有ChildTip而异
            If Tips(n).ChildTipsTotal = 0 Then
                nTipWidth = nTipWidth + cXDisX
            Else
                nTipWidth = nTipWidth + cCTParentXDis
            End If
            'nTipWidth位移至关闭按钮前
            
            Tips(n).nXX = Tips(n).nX + nTipWidth                                '存储Tip宽度
            
            nTipWidth = nTipWidth + cXCR * 2                                    'nTipWidth位移至关闭按钮后
        End If
        
        nTipWidth = nTipWidth + cDisX                                           'nTipWidth位移至Tip末尾
        
        Tips(n).nWidth = nTipWidth
        
        nTipWidth = 0                                                           '初始化变量
        
        '○以上完成Tip宽度运算
        
    Next
    
    '○以上完成所有Tips运算
    
End Sub

Private Sub DrawTip(objTip As TB_Tip)
    On Error GoTo iErr
    
    With objTip
        
        Dim nDrawTipH As Single
        Dim nDrawTextX As Single, nDrawTextY As Single
        Dim nDrawTipColor As Long, nDrawXColor As Long, nDrawLineColor As Long
        Dim nDrawTextColor As Long
        
        If .ParentTip = 0 Then
            nDrawTipH = pTipsH
            nDrawTextY = pTextY
            
            nDrawTextX = .nStrX + 1                                             '1为Tip标题向右偏移量
            
            SetFont .Active, TipsFontSize
        Else
            nDrawTipH = pCTTipsH
            nDrawTextY = pCTTextY
            
            nDrawTextX = .nStrX + 1                                             '1为Tip标题向右偏移量
            
            SetFont .Active, ChildTipsFontSize
        End If
        
        If .Active = False Then
            nDrawTipColor = TipsUAColor
            UserControl.ForeColor = cTextUAColor
        Else
            nDrawTipColor = TipsAColor
            UserControl.ForeColor = cTextAColor
        End If
        
        If Me.Enabled = True Then
            If objTip.bMouseOnX = False Then
                nDrawXColor = XUAColor
            Else
                nDrawXColor = XAColor
            End If
        Else
            nDrawXColor = UnEnabledColor
        End If
        
        If Me.Enabled = True Then
            nDrawLineColor = LineColor
        Else
            nDrawLineColor = UnEnabledColor
        End If
        
        DrawRect objTip.nX, objTip.nWidth, nDrawTipH, nDrawLineColor, nDrawTipColor
        
        iTextOut nDrawTextX, nDrawTextY, objTip.Caption
        
        If objTip.ShowCloseButton = True Then
            Dim nOldDW As Long
            
            nOldDW = UserControl.DrawWidth
            
            UserControl.DrawWidth = cXDrawWidth
            
            iLine .nXX, nDrawTipH / 2 + cXCR - 0.5, .nXX + 2 * cXCR, nDrawTipH / 2 - cXCR - 0.5, nDrawXColor
            iLine .nXX, nDrawTipH / 2 - cXCR - 0.5, .nXX + 2 * cXCR, nDrawTipH / 2 + cXCR - 0.5, nDrawXColor
            '0.5为关闭按钮向下偏移量
            
            UserControl.DrawWidth = nOldDW
        End If
        
        If objTip.nIconX <> 0 Then
            iPaintIcon objTip.Icon, objTip.nIconX, nDrawTipH / 2, IconWidth, IconHeight
        End If
        
    End With
    
    Exit Sub
iErr:
    DBErr "TipsBar - Sub DrawTip", "objTip.Caption = " & objTip.Caption
End Sub

Private Sub MakePictureGray(nX0, nY0, nWidth, nHeight)
    Dim nRed As Integer, nGreen As Integer, nBlue As Integer
    Dim nColor As Long, nGray As Long
    Dim nX1 As Integer, nY1 As Integer
    For nX1 = 0 To nWidth
        For nY1 = 0 To nHeight
            nColor = UserControl.Point(nX0 + nX1, nY0 + nY1)
            nRed = (nColor And &HFF)
            nGreen = (nColor And 62580) / 256
            nBlue = (nColor And &HFF00) / 65536
            nGray = (nRed + nGreen + nBlue) / 3
            UserControl.PSet (nX0 + nX1, nY0 + nY1), RGB(nGray, nGray, nGray)
            DoEvents
        Next
    Next
    
End Sub

Private Sub iPaintIcon(picIcon As StdPicture, ByVal nX, ByVal nMY, ByVal nWidth, ByVal nHeight)
    Select Case Me.Align
    Case tbaTop
        UserControl.PaintPicture picIcon, nX, UserControl.ScaleHeight - (nMY + nHeight / 2 - 0.7), nWidth, nHeight
        If Me.Enabled = False Then MakePictureGray nX, UserControl.ScaleHeight - (nMY + nHeight / 2 - 0.7), nWidth, nHeight
        '0.7为图标向下偏移量
    Case tbaBottom
        UserControl.PaintPicture picIcon, nX, nMY - nHeight / 2 + 0.1, nWidth, nHeight
        If Me.Enabled = False Then MakePictureGray nX, nMY - nHeight / 2 + 0.1, nWidth, nHeight
        '0.1为图标向下偏移量
    Case tbaRight
        UserControl.PaintPicture picIcon, nMY - nHeight / 2, nX, nWidth, nHeight
        If Me.Enabled = False Then MakePictureGray nMY - nHeight / 2, nX, nWidth, nHeight
    Case tbaLeft
        UserControl.PaintPicture picIcon, UserControl.ScaleWidth - (nMY - nHeight / 2) - nHeight, nX, nWidth, nHeight
        If Me.Enabled = False Then MakePictureGray UserControl.ScaleWidth - (nMY - nHeight / 2) - nHeight, nX, nWidth, nHeight
    End Select
End Sub

Friend Function CallEventCreateTip(ByRef sCaption As String, ByRef Key, _
    Optional ByRef bCloseButton As Boolean = False, _
    Optional ByRef nParent = 0, _
    Optional ByRef sPopUp As String) As Boolean
    
    Dim bRes As Boolean
    bRes = True
    
    RaiseEvent TipCreate(sCaption, Key, bCloseButton, nParent, sPopUp, bRes)
    
    CallEventCreateTip = bRes
End Function

'Function CreateTip
'输入: ByVal sCaption As String
'      ByVal Key
'      Optional ByVal bCloseButton As Boolean = True
'      Optional ByVal nParent = 0 —— ParentTip的Index或Key,若为0,则创建常规标签,若不存在,则返回0
'      Optional ByVal sPopUp As String
'返回值: Long —— 返回创建的Tip的Index,若为0,则表示失败,若创建子标签,则返回父标签Index

Public Function CreateTip(ByVal sCaption As String, ByVal Key, _
    Optional ByVal bCloseButton As Boolean = True, _
    Optional ByVal nParent = 0, _
    Optional ByVal sPopUp As String) As Long
    
    On Error GoTo iErr
    
    If KeyToIndex(Key) = True Then Exit Function
    
    If nParent = 0 Then
        
        If CallEventCreateTip(sCaption, Key, bCloseButton, nParent, sPopUp) = False Then Exit Function
        
        ReDim Preserve m_Tips(Me.TipsTotal + 1)
        
        Dim NewTipIndex As Long
        
        NewTipIndex = Me.TipsTotal
        
        Set Tips(NewTipIndex) = New TB_Tip
        
        With Tips(NewTipIndex)
            .iTipIndex = NewTipIndex
            Set .ObjTipsBar = Me
            .Caption = sCaption
            .Key = Key
            .ShowCloseButton = bCloseButton
            .PopUpStr = sPopUp
        End With
        
        CreateTip = NewTipIndex
    Else
        
        If KeyToIndex(nParent) = False Then Exit Function
        
        If CallEventCreateTip(sCaption, Key, bCloseButton, nParent, sPopUp) = False Then Exit Function
        
        Tips(nParent).ChildTips(Tips(nParent).CreateChildTip(sCaption, Key, bCloseButton)).PopUpStr = sPopUp
        
        CreateTip = nParent
        
    End If
    
iErr:
    DBErr "TipsBar - Function CreateTip", "sCaption = " & sCaption, "Key = " & Key, "bCloseButton = " & bCloseButton, "nParent = " & nParent, "Key = " & sPopUp
    
End Function

'Function DeleteTip
'输入: ByVal n —— Tip的Index或Key
'      Optional ByVal nChild = 0 —— ChildTip的Index或Key
'返回值: Bool —— True表示成功，False表示失败
'注意: 此函数考虑删除Tip后焦点切换，可在DeleteTip事件中阻止焦点自动切换

Public Function DeleteTip(ByVal n, Optional ByVal nChild = 0) As Boolean
    DeleteTip = False
    
    On Error GoTo iErr
    
    If KeyToIndex(n) = False Then Exit Function
    If nChild <> 0 Then If Tips(n).KeyToIndex(nChild) = False Then Exit Function
    
    Dim bActivateAnother As Boolean: bActivateAnother = True
    If CallEventTipDelete(n, nChild, bActivateAnother) = False Then Exit Function
    
    If nChild = 0 Then
        
        If nActiveTip = n Then
            nActiveTip = 0
        ElseIf nActiveTip > n Then
            nActiveTip = nActiveTip - 1
        End If
        
        If mLastTip <> -1 Then mLastTip = 0
        
        Dim i As Long
        For i = n To Me.TipsTotal - 1
            Set Tips(i) = Tips(i + 1)
            Tips(i).iTipIndex = i
        Next
        
        ReDim Preserve m_Tips(Me.TipsTotal - 1)
        
    Else
        
        Tips(n).DeleteChildTip nChild
        
        If CStr(nActiveTip) = n & "." & nChild Then
            
            If Tips(n).ChildTipsTotal > 0 Then
                
                If bActivateAnother = True Then
                    If nChild = 1 Then
                        ActivateTip n, 1
                    Else
                        ActivateTip n, nChild - 1
                    End If
                End If
            Else
                If bActivateAnother = True Then ActivateTip n
            End If
            
        End If
    End If
    
    DeleteTip = True
    
    Exit Function
iErr:
    DBErr "TipsBar - Function DeleteTip", "n = " & n, "nChild = " & nChild
End Function


Friend Function CallEventTipDelete(ByVal nTip, Optional ByVal nChild = 0, Optional ByRef bActivateAnother As Boolean = True) As Boolean
    Dim bRes As Boolean
    
    bRes = True
    
    RaiseEvent TipDelete(nTip, nChild, bRes, bActivateAnother)
    
    CallEventTipDelete = bRes
End Function


Private Function IsAllEnglish(ByVal sStr As String) As Boolean
    Dim i As Long
    
    IsAllEnglish = True
    
    For i = 1 To Len(sStr)
        If AscW(Mid(sStr, i, 1)) > 127 Then IsAllEnglish = False: Exit Function
    Next
End Function

Private Sub DrawRect(nLeft, nWidth, nHeight, nLColor As Long, nFColor As Long)
    
    UserControl.FillStyle = 0                                                   '设置UserControl的填充模式为实心填充
    UserControl.FillColor = nFColor                                             '设置UserControl的填充颜色
    
    
    iCircle nLeft + cR, nHeight - cR, cR, nLColor                               '绘制并填充左上角的扇形
    iCircle nLeft + nWidth - cR, nHeight - cR, cR, nLColor                      '绘制并填充右上角的扇形
    
    iLine nLeft + cR, nHeight - cR, nLeft + nWidth - cR, nHeight, nFColor, True, True '绘制并填充顶部实心矩形
    iLine nLeft, 0, nLeft + nWidth, nHeight - cR, nFColor, True, True           '绘制并填充下方大实心矩形
    
    UserControl.FillStyle = 1                                                   '设置UserControl的填充模式为无填充
    
    iLine nLeft, 0, nLeft, nHeight - cR + 1, nLColor                            '绘制左方Tip边框线条
    iLine nLeft + cR, nHeight, nLeft + nWidth - cR + 1, nHeight, nLColor        '绘制上方Tip边框线条
    iLine nLeft + nWidth, 0, nLeft + nWidth, nHeight - cR + 1, nLColor          '绘制右方Tip边框线条
    
End Sub

Private Sub SetFont(Optional bBold As Boolean = False, Optional nSize As Long = 9)
    UserControl.Font.Bold = bBold
    UserControl.Font.size = nSize
End Sub

Private Function AtoR(ByVal nA) As Single
    AtoR = nA * Pi / 180
End Function

Private Sub iCircleSub(nX, nY, nR, Optional lColor, Optional nStart, Optional nEnd, Optional nAspect)
    If Not IsMissing(nStart) And Not IsMissing(nEnd) Then
        If IsMissing(nAspect) Then
            UserControl.Circle (nX, nY), nR, lColor, nStart, nEnd
        Else
            UserControl.Circle (nX, nY), nR, lColor, nStart, nEnd, nAspect
        End If
    ElseIf IsMissing(nStart) And IsMissing(nEnd) Then
        If IsMissing(nAspect) Then
            UserControl.Circle (nX, nY), nR, lColor
        Else
            UserControl.Circle (nX, nY), nR, lColor, , , nAspect
        End If
    Else
        Err.Raise "36102", , "TipsBar : Circle函数不能只有Start或End其中之一!"
    End If
End Sub


Private Sub iLineSub(nX1, nY1, nX2, nY2, Optional lColor, Optional bB As Boolean = False, Optional bF As Boolean = False)
    If IsMissing(lColor) Then
        If bB = True And bF = True Then
            UserControl.Line (nX1, nY1)-(nX2, nY2), , BF
        ElseIf bB = False And bF = False Then
            UserControl.Line (nX1, nY1)-(nX2, nY2)
        ElseIf bB = True And bF = False Then
            UserControl.Line (nX1, nY1)-(nX2, nY2), , B
        ElseIf bF = True And bF = False Then
            Err.Raise "36101", , "TipsBar : Line函数不能只使用F关键字而不使用B关键字!"
        End If
    Else
        If bB = True And bF = True Then
            UserControl.Line (nX1, nY1)-(nX2, nY2), lColor, BF
        ElseIf bB = False And bF = False Then
            UserControl.Line (nX1, nY1)-(nX2, nY2), lColor
        ElseIf bB = True And bF = False Then
            UserControl.Line (nX1, nY1)-(nX2, nY2), lColor, B
        ElseIf bF = True And bF = False Then
            Err.Raise "36101", , "TipsBar : Line函数不能只使用F关键字而不使用B关键字!"
        End If
    End If
    
End Sub


Private Sub UserControl_Resize()
    ChangeSize
End Sub

Private Sub ChangeSize()
    Select Case Me.Align
    Case tbaTop, tbaBottom
        UserControl.Height = UserControl.ScaleY(pTipsH + 2 + cTipsTop, 3, 1)
    Case tbaLeft, tbaRight
        UserControl.Width = UserControl.ScaleX(pTipsH + 4 + cTipsTop, 3, 1)
    End Select
End Sub

Private Sub UserControl_Show()
    
    mLastTip = 0
    
    CalcPublicVar
    
    
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BackStyle() As Integer
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property



Private Sub UserControl_Click()
    RaiseEvent Click
    
    '    mDealAction mLastX, mLastY, mLastTip, maClick
End Sub

Public Sub CallEventTipClick(nTip, Optional nChild = 0)
    Dim bRes As Boolean
    
    bRes = True
    
    RaiseEvent TipClick(nTip, nChild, bRes)
    
    If bRes = True Then ActivateTip nTip, nChild
    
End Sub

Public Sub CallEventTipXClick(nTip, Optional nChild = 0)
    Dim bRes As Boolean
    bRes = True
    RaiseEvent TipXClick(nTip, nChild, bRes)
    If bRes = True Then DeleteTip nTip, nChild
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    If New_Enabled = False Then
        nOldForeColor = Me.ForeColor
        Me.ForeColor = UnEnabledColor
    Else
        Me.ForeColor = nOldForeColor
    End If
    DrawTips
    PropertyChanged "Enabled"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    mDealAction X, Y, mLastTip, maMouseDown, Button, Shift
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    mDealAction X, Y, mLastTip, maMouseUp, Button, Shift
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_Align = m_def_Align
    m_OnlyOneActive = m_def_OnlyOneActive
    m_AutoBottomLine = m_def_AutoBottomLine
    m_AutoBottomLineColor = m_def_AutoBottomLineColor
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    
    m_Align = PropBag.ReadProperty("Align", m_def_Align)
    m_OnlyOneActive = PropBag.ReadProperty("OnlyOneActive", m_def_OnlyOneActive)
    m_AutoBottomLine = PropBag.ReadProperty("AutoBottomLine", m_def_AutoBottomLine)
    m_AutoBottomLineColor = PropBag.ReadProperty("AutoBottomLineColor", m_def_AutoBottomLineColor)
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    
    Call PropBag.WriteProperty("Align", m_Align, m_def_Align)
    Call PropBag.WriteProperty("OnlyOneActive", m_OnlyOneActive, m_def_OnlyOneActive)
    Call PropBag.WriteProperty("AutoBottomLine", m_AutoBottomLine, m_def_AutoBottomLine)
    Call PropBag.WriteProperty("AutoBottomLineColor", m_AutoBottomLineColor, m_def_AutoBottomLineColor)
End Sub

Public Property Get Align() As enmTipsBarAlignMode
    Align = m_Align
End Property

Public Property Let Align(ByVal New_Align As enmTipsBarAlignMode)
    m_Align = New_Align
    PropertyChanged "Align"
End Property

Public Property Get OnlyOneActive() As Boolean
    OnlyOneActive = m_OnlyOneActive
End Property

Public Property Let OnlyOneActive(ByVal New_OnlyOneActive As Boolean)
    m_OnlyOneActive = New_OnlyOneActive
    PropertyChanged "OnlyOneActive"
End Property

Public Property Get AutoBottomLine() As Boolean
    AutoBottomLine = m_AutoBottomLine
End Property

Public Property Let AutoBottomLine(ByVal New_AutoBottomLine As Boolean)
    m_AutoBottomLine = New_AutoBottomLine
    PropertyChanged "AutoBottomLine"
End Property

Public Property Get AutoBottomLineColor() As OLE_COLOR
    AutoBottomLineColor = m_AutoBottomLineColor
End Property

Public Property Let AutoBottomLineColor(ByVal New_AutoBottomLineColor As OLE_COLOR)
    m_AutoBottomLineColor = New_AutoBottomLineColor
    PropertyChanged "AutoBottomLineColor"
End Property

Friend Property Get nActiveTip() As Single
    nActiveTip = m_nActiveTip
End Property

Friend Property Let nActiveTip(ByVal New_nActiveTip As Single)
    m_nActiveTip = New_nActiveTip
    PropertyChanged "nActiveTip"
End Property



