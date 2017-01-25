VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_Tip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_StrCaption As String
Private m_StrPopUpStr As String
Private m_bShowCloseButton As Boolean
Private m_LonParentTip As Long
Private m_ChildTips() As TB_Tip
Private m_bActive As Boolean
Private m_Icon As StdPicture

Public iTipIndex As Single

Private m_bHaveIcon As Boolean
Private m_bMouseOnX As Boolean
Private m_nXX As Single
Private m_nStrX As Single
Private m_nIconX As Single
Private m_nWidth As Single
Private m_nX As Single
Private m_ObjTipsBar As TipsBar

Private m_Key As String

Friend Property Get Key() As String
    Key = m_Key
End Property

Friend Property Let Key(ByVal NewValue As String)
    m_Key = NewValue
End Property

Friend Property Get ObjTipsBar() As TipsBar
    Set ObjTipsBar = m_ObjTipsBar
End Property

Friend Property Set ObjTipsBar(ByVal XCTValue As TipsBar)
    Set m_ObjTipsBar = XCTValue
End Property

Friend Property Get nX() As Single
    nX = m_nX
End Property

Friend Property Let nX(ByVal NewValue As Single)
    m_nX = NewValue
End Property

Friend Property Get nWidth() As Single
    nWidth = m_nWidth
End Property

Friend Property Let nWidth(ByVal NewValue As Single)
    m_nWidth = NewValue
End Property

Friend Property Get nIconX() As Single
    nIconX = m_nIconX
End Property

Friend Property Let nIconX(ByVal NewValue As Single)
    m_nIconX = NewValue
End Property

Friend Property Get nStrX() As Single
    nStrX = m_nStrX
End Property

Friend Property Let nStrX(ByVal NewValue As Single)
    m_nStrX = NewValue
End Property

Friend Property Get nXX() As Single
    nXX = m_nXX
End Property

Friend Property Let nXX(ByVal NewValue As Single)
    m_nXX = NewValue
End Property

Friend Property Get bMouseOnX() As Boolean
    bMouseOnX = m_bMouseOnX
End Property

Friend Property Let bMouseOnX(ByVal bValue As Boolean)
    m_bMouseOnX = bValue
End Property

Friend Property Get bHaveIcon() As Boolean
    bHaveIcon = m_bHaveIcon
End Property

Friend Property Let bHaveIcon(ByVal bValue As Boolean)
    m_bHaveIcon = bValue
End Property



Public Property Get Icon() As StdPicture
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal picValue As StdPicture)
    bHaveIcon = True
    Set m_Icon = picValue
End Property


Public Sub LoadIconFromFile(ByVal sFileName As String)
    Set Me.Icon = LoadPicture(sFileName)
End Sub


Public Property Get Active() As Boolean
    Active = m_bActive
End Property

Public Property Let Active(ByVal bValue As Boolean)
    m_bActive = bValue
End Property

Public Sub Activate()
    ObjTipsBar.ActivateTip Int(iTipIndex), ObjTipsBar.GetChildTipIndexFromStr(iTipIndex)
    ObjTipsBar.DrawTips
End Sub

Public Sub UnActivate()
    ObjTipsBar.UnActivateTip Int(iTipIndex), ObjTipsBar.GetChildTipIndexFromStr(iTipIndex)
    ObjTipsBar.DrawTips
End Sub



Public Property Get ChildTips(ByVal n) As TB_Tip
    If Int(iTipIndex) <> CSng(iTipIndex) Then Err.Raise 3612, "TipsBar Control - Tips(" & Me.iTipIndex & ").ChildTips", "不能嵌套ChildTip！"
    
    If KeyToIndex(n) = True Then
        Set ChildTips = m_ChildTips(n)
    Else
        Set ChildTips = Nothing
    End If
End Property

Public Function CreateChildTip(ByVal Caption As String, ByVal Key, Optional ByVal CloseButton As Boolean = True) As Long
    
    If ObjTipsBar.CallEventCreateTip(Caption, Key, CloseButton, iTipIndex, "") = False Then Exit Function
    
    ReDim Preserve m_ChildTips(Me.ChildTipsTotal + 1)
    
    Dim NewChildTipIndex As Long
    
    NewChildTipIndex = Me.ChildTipsTotal
    
    Set m_ChildTips(NewChildTipIndex) = New TB_Tip
    
    With m_ChildTips(NewChildTipIndex)
        
        Set .ObjTipsBar = ObjTipsBar
        
        .iTipIndex = iTipIndex & "." & NewChildTipIndex
        .ParentTip = iTipIndex
        
        .Caption = Caption
        .Key = Key
        .ShowCloseButton = CloseButton
        
    End With
    
    CreateChildTip = NewChildTipIndex
End Function


'KeyToIndex
'输入:ByRef n —— Key或Index → Index
'返回值: Bool —— True表示成功,False表示失败

Public Function KeyToIndex(ByRef n) As Boolean
    If IsNumeric(n) And Val(n) <= Me.ChildTipsTotal Then
        KeyToIndex = True
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To Me.ChildTipsTotal
        If ChildTips(i).Key = n Then
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
    If IsNumeric(n) And Val(n) <= Me.ChildTipsTotal Then
        ReutrnIndexByKey = n
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To Me.ChildTipsTotal
        If ChildTips(i).Key = n Then
            ReutrnIndexByKey = i
            Exit Function
        End If
    Next
End Function


Public Property Set ChildTips(ByVal n, NewValue As TB_Tip)
    If Int(iTipIndex) <> CSng(iTipIndex) Then Err.Raise 36201, , "TipsBar : 不能嵌套ChildTip！"
    
    If KeyToIndex(n) = True Then
        m_ChildTips(n) = NewValue
    Else
        Err.Raise "3611", "TipsBar Control - Tips(" & Me.iTipIndex & "Property Set ChildTips", "找不到 Tip ！"
    End If
    
End Property

Public Property Get ChildTipsTotal() As Long
    ChildTipsTotal = UBound(m_ChildTips)
End Property


Public Function DeleteChildTip(ByVal n) As Boolean
    DeleteChildTip = False
    
    If KeyToIndex(n) = False Then Exit Function
    
    If ObjTipsBar.CallEventTipDelete(iTipIndex, n) = False Then Exit Function
    
    Dim i As Long
    For i = n To Me.ChildTipsTotal - 1
        Set m_ChildTips(i) = m_ChildTips(i + 1)
        m_ChildTips(i).iTipIndex = iTipIndex & "." & i
    Next
    ReDim Preserve m_ChildTips(Me.ChildTipsTotal - 1)
    
    DeleteChildTip = True
End Function


Public Sub Delete()
    If Int(iTipIndex) = CSng(iTipIndex) Then
        ObjTipsBar.DeleteTip iTipIndex
    Else
        Call ObjTipsBar.Tips(Me.ParentTip).DeleteChildTip(ObjTipsBar.GetChildTipIndexFromStr(iTipIndex))
    End If
End Sub


Public Property Get ParentTip() As Long
    ParentTip = m_LonParentTip
End Property

Public Property Let ParentTip(ByVal LonValue As Long)
    m_LonParentTip = LonValue
End Property




Public Property Get ShowCloseButton() As Boolean
    ShowCloseButton = m_bShowCloseButton
End Property

Public Property Let ShowCloseButton(ByVal bValue As Boolean)
    m_bShowCloseButton = bValue
End Property



Public Property Get PopUpStr() As String
    PopUpStr = m_StrPopUpStr
End Property

Public Property Let PopUpStr(ByVal StrValue As String)
    m_StrPopUpStr = StrValue
End Property

Public Property Get Caption() As String
    Caption = m_StrCaption
End Property

Public Property Let Caption(ByVal StrValue As String)
    m_StrCaption = StrValue
End Property

Private Sub Class_Initialize()
    ReDim m_ChildTips(0)
End Sub

