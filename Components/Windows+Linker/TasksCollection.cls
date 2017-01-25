VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TasksCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'sName支持以下指令
'“*Any*”—— 任意
'“*Include* ” + ... —— 包含...
'“*Not Include* ” + ... —— 不包含...

Public Type Task
    hParent As Long
    sClass As String
    sName As String
    x As Long
    y As Long
    Width As Long
    Height As Long
    Keep As Boolean
    Tag As String
End Type

Event Find(Task As Task, ByVal hWnd As Long)

Public Count As Long
Private Task(100) As Task

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long


Public Sub Add(Optional ByVal hParent As Long = 0, _
    Optional ByVal sClass As String, _
    Optional ByVal sName As String, _
    Optional ByVal x As Long, _
    Optional ByVal y As Long, _
    Optional ByVal Width As Long, _
    Optional ByVal Height As Long, _
    Optional ByVal Keep As Boolean, _
    Optional ByVal Tag As String)
    
    Count = Count + 1
    With Task(Count)
        .hParent = hParent
        .sClass = sClass
        .sName = sName
        .x = x
        .y = y
        .Width = Width
        .Height = Height
        .Keep = Keep
        .Tag = Tag
    End With
End Sub

Public Sub CleanUp(ByVal hParent As Long)
    Dim i As Long, j As Long
    For i = 1 To Count
        If Task(i).hParent <> hParent Then
            j = j + 1
            If i <> j Then Task(j) = Task(i)
        End If
    Next
    'DBPrint "Clean Up " & Count & "->" & j
    Count = j
End Sub

Public Sub Clear()
    Count = 0
End Sub

Public Sub Find(ByVal hWnd As Long)
       
    Dim i As Long
    
    For i = 1 To Count Step 1

        If Task(i).sClass <> iGetClassName(hWnd) Then GoTo iNext
        
        If Task(i).sName <> "*Any*" Then
            If Task(i).sName <> iGetCaption(hWnd) Then GoTo iNext
        End If

        If Task(i).hParent <> 0 Then
            If Task(i).hParent <> GetParent(hWnd) Then GoTo iNext
        End If
                
        RaiseEvent Find(Task(i), hWnd)
        
        If i > Count Then Exit For
        
        If Task(i).Keep = False Then
            If i < Count Then Task(i) = Task(Count)
            Count = Count - 1
            i = i - 1
        End If
        
        Exit Sub '发现一个匹配即完成
        
iNext:
    
    If i >= Count Then Exit For

    Next
    
End Sub
