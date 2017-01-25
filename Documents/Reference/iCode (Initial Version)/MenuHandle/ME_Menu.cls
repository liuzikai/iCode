VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ME_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mvarParent As ME_Menu
Private mvarMenuName As String
Private mvarCaption As String
Private mvarChecked As Boolean
Private mvarEnabled As Boolean
Private mvarVisible As Boolean
Private mvarHelpContextID As Long
Private mvarIndex As Long
Private mvarNegotiatePosition As Long
Private mvarWindowList As Boolean
Private mvarImage As Variant
Private mvarSelectImage As Variant
Private mvarChildren() As ME_Menu
Private mvarShortCut As String

Public Property Let ShortCut(ByVal vData As String)
    mvarShortCut = vData
End Property

Public Property Get ShortCut() As String
    ShortCut = mvarShortCut
End Property


Public Property Get Children(ByVal Index) As ME_Menu
    Set Children = mvarChildren(Index)
End Property

Public Property Set Children(ByVal Index, ByVal vData As ME_Menu)
    Set mvarChildren(Index) = vData
End Property

Public Property Get ChildrenTotal() As Long
    ChildrenTotal = UBound(mvarChildren)
End Property

Public Function AddChildMenu(Optional ByVal Name, Optional ByVal Caption, Optional ByVal Checked As Boolean = False, _
    Optional ByVal Enabled As Boolean = True, Optional ByVal Visible As Boolean = True, Optional ByVal Index As Long, _
    Optional ByVal Image As Variant, Optional ByVal NegotiatePosition As Long = 0, Optional ByVal WindowList As Boolean = False) _
    As ME_Menu
    
    ReDim Preserve mvarChildren(ChildrenTotal + 1)
    
    Set mvarChildren(ChildrenTotal) = New ME_Menu
    
    With mvarChildren(ChildrenTotal)
        Set .Parent = Me
        
        .MenuName = Name
        .Caption = Caption
        .Checked = Checked
        .Enabled = Enabled
        .Visible = Visible
        .Index = Index
        Set .Image = Image
        .NegotiatePosition = NegotiatePosition
        .WindowList = WindowList
    End With
    
    Set AddChildMenu = mvarChildren(ChildrenTotal)
End Function

Public Sub RemoveChildMenu(ByVal Index)
    Dim i As Long
    For i = Index To ChildrenTotal - 1
        Set mvarChildren(i) = mvarChildren(i + 1)
    Next
    ReDim Preserve mvarChildren(ChildrenTotal - 1)
End Sub

Public Property Set Image(ByVal vData As Variant)
    Set mvarImage = vData
End Property

Public Property Get Image() As Variant
    Set Image = mvarImage
End Property

Public Property Set SelectImage(ByVal vData As Variant)
    Set mvarSelectImage = vData
End Property

Public Property Get SelectImage() As Variant
    Set SelectImage = mvarSelectImage
End Property


Public Property Let WindowList(ByVal vData As Boolean)
    mvarWindowList = vData
End Property

Public Property Get WindowList() As Boolean
    WindowList = mvarWindowList
End Property


Public Property Let NegotiatePosition(ByVal vData As Long)
    mvarNegotiatePosition = vData
End Property

Public Property Get NegotiatePosition() As Long
    NegotiatePosition = mvarNegotiatePosition
End Property


Public Property Let Index(ByVal vData As Long)
    mvarIndex = vData
End Property

Public Property Get Index() As Long
    Index = mvarIndex
End Property


Public Property Let HelpContextID(ByVal vData As Long)
    mvarHelpContextID = vData
End Property

Public Property Get HelpContextID() As Long
    HelpContextID = mvarHelpContextID
End Property


Public Property Let Visible(ByVal vData As Boolean)
    mvarVisible = vData
End Property

Public Property Get Visible() As Boolean
    Visible = mvarVisible
End Property


Public Property Let Enabled(ByVal vData As Boolean)
    mvarEnabled = vData
End Property

Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property


Public Property Let Checked(ByVal vData As Boolean)
    mvarChecked = vData
End Property

Public Property Get Checked() As Boolean
    Checked = mvarChecked
End Property


Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
End Property

Public Property Get Caption() As String
    Caption = mvarCaption
End Property


Public Property Let MenuName(ByVal vData As String)
    mvarMenuName = vData
End Property

Public Property Get MenuName() As String
    MenuName = mvarMenuName
End Property


Public Property Set Parent(ByVal vData As ME_Menu)
    Set mvarParent = vData
End Property


Public Property Get Parent() As ME_Menu
    Set Parent = mvarParent
End Property



Private Sub Class_Initialize()
    ReDim mvarChildren(0)
    Enabled = True
    Visible = True
End Sub
