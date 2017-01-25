VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AC_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit



Private m_ItemName As String
Private m_Kind As AC_ItemKind
Private m_Code As String

Public Property Let Kind(ByVal vData As AC_ItemKind)
    m_Kind = vData
End Property

Public Property Get Kind() As AC_ItemKind
    Kind = m_Kind
End Property

Public Property Let ItemName(ByVal vData As String)
    m_ItemName = vData
End Property

Public Property Get ItemName() As String
Attribute ItemName.VB_UserMemId = 0
    ItemName = m_ItemName
End Property

'注意:Collection标签上显示的将会是Item的Code！

Public Property Get Code() As String
    Code = m_Code
End Property

Public Property Let Code(vData As String)
    m_Code = vData
End Property
