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

Private mvarStr As String
Private mvarKind As AC_ItemKind
Private mvarCode As String

Public Property Let Kind(ByVal vData As AC_ItemKind)
    mvarKind = vData
End Property

Public Property Get Kind() As AC_ItemKind
    Kind = mvarKind
End Property

Public Property Let Str(ByVal vData As String)
    mvarStr = vData
End Property

Public Property Get Str() As String
Attribute Str.VB_UserMemId = 0
    Str = mvarStr
End Property

Public Property Get Code() As String
    Code = mvarCode
End Property

Public Property Let Code(vData As String)
    mvarCode = vData
End Property
