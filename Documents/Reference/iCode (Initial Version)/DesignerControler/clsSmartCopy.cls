VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSmartCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const SC_CodeFileName As String = "\iCode - Smart Copy - Code.Code"

Private m_SC_Buttons() As New SC_Button

Private SmartCopySign As String

Private SmartCopying As Boolean
Private SmartPasting As Boolean

Friend Sub SC_Buttons_Add(ByVal ButtonType As SC_ButtonType, ByVal Bar As String, ByVal Before As Long)
    ReDim Preserve m_SC_Buttons(UBound(m_SC_Buttons) + 1)
    'Set m_SC_Buttons(UBound(m_SC_Buttons)) = New SC_Button
    m_SC_Buttons(UBound(m_SC_Buttons)).Load ButtonType, Bar, Before
End Sub

Friend Property Let SC_Buttons_Enabled(ByVal ButtonType As SC_ButtonType, ByVal NewValue As Boolean)
    Dim i As Long
    For i = 1 To UBound(m_SC_Buttons)
        If m_SC_Buttons(i).ButtonType = ButtonType Then m_SC_Buttons(i).Enabled = NewValue
    Next
End Property

Friend Sub SC_Buttons_Clear()
    Dim i As Long
    For i = 1 To UBound(m_SC_Buttons)
        Set m_SC_Buttons(i) = Nothing
    Next
    ReDim m_SC_Buttons(0)
End Sub

Public Sub Init()
    
    SC_Buttons_Add SC_BT_Cut, "Edit", 3 + 1
    SC_Buttons_Add SC_BT_Copy, "Edit", 4 + 1 + 1
    SC_Buttons_Add SC_BT_Paste, "Edit", 5 + 1 + 2
    
    'SC_Buttons_Add SC_BT_Paste, "Forms", 4 + 1
    
    'SC_Buttons_Add SC_BT_Cut, "Controls", 1 + 1
    'SC_Buttons_Add SC_BT_Copy, "Controls", 2 + 1 + 1
    'SC_Buttons_Add SC_BT_Paste, "Controls", 3 + 1 + 2
End Sub

Public Sub UnLoad()
    SC_Buttons_Clear
End Sub

Public Sub SC_MsgMenuShow()
    SC_Buttons_Enabled(SC_BT_Copy) = (iState = isDesign)
    SC_Buttons_Enabled(SC_BT_Cut) = (iState = isDesign)
    SC_Buttons_Enabled(SC_BT_Paste) = SC_CanPasteCode
End Sub

Private Function SC_GetCode() As String
    If Dir(App.Path & SC_CodeFileName) = "" Then
        SC_GetCode = ""
    Else
        Dim i As Long, s As String
        i = FreeFile
        
        Open App.Path & SC_CodeFileName For Input As #i
        Input #i, s
        Close #i
        
        SC_GetCode = s
    End If
End Function

Private Sub SC_SetCode(ByVal sCode As String)
    
    Dim i As Long
    i = FreeFile
    
    Open App.Path & SC_CodeFileName For Output As #i
    Write #i, sCode
    Close #i
    
End Sub

Public Function SC_CanPasteCode() As Boolean
    Dim s As String
    s = SC_GetCode
    SC_CanPasteCode = (iState = isDesign) And VBIns.CommandBars("Edit").Controls("Õ³Ìù(&P)").Enabled = True And LeftIs(s, SmartCopySign)
End Function

Public Function SC_MsgCopyCut(Optional ByVal bCut As Boolean = False) As Boolean
    If iState = isDesign And VBIns.SelectedVBComponent.Designer.SelectedVBControls.Count > 0 And SmartCopying = False Then
        Select Case MsgBox("ÊÇ·ñ¸´ÖÆ´úÂë£¿", vbYesNoCancel, "iCode - SmartCopy")
        Case vbYes
            SmartCopying = True
            SC_MsgCopyCut = True
            SC_CopyCutCode bCut
        Case vbNo
            SC_MsgCopyCut = True
        Case vbCancel
            SC_MsgCopyCut = False
        End Select
    ElseIf SmartCopying = True Then
        SC_MsgCopyCut = True
        SmartCopying = False
    End If
End Function


Public Function SC_MsgPaste() As Boolean
    If SC_CanPasteCode And SmartPasting = False Then
        
        Dim s As String
        s = SC_GetCode
        
        SmartPasting = True
        SC_MsgPaste = True
        SC_PasteCode s
        
    ElseIf SmartPasting = True Then
        SC_MsgPaste = True
        SmartPasting = False
    End If
End Function

Public Sub SC_CopyCutCode(Optional ByVal bCut As Boolean = False)
    Dim s As String
    
    Dim n As Long
    
    Dim AD As VBForm
    Set AD = VBIns.SelectedVBComponent.Designer
    
    If AD.SelectedVBControls.Count <> 0 Then s = SmartCopySign
    
    For n = 1 To AD.SelectedVBControls.Count
        
        Dim i As Long
        With VBIns.SelectedVBComponent.CodeModule
            For i = 1 To .Members.Count
                If i > .Members.Count Then Exit For
                If LeftIs(.Members.item(i).Name, AD.SelectedVBControls.item(n - 1).Properties("Name") & "_") And .Members.item(i).type = vbext_mt_Method Then
                    Dim ls As Long, lc As Long
                    ls = .ProcStartLine(.Members.item(i).Name, vbext_pk_Proc)
                    lc = .ProcCountLines(.Members.item(i).Name, vbext_pk_Proc)
                    s = s & .Lines(ls, lc)
                    If bCut Then .DeleteLines ls, lc
                End If
            Next
        End With
        
    Next
    
    DoEvents
    
    SC_Clipboard_SetText s
End Sub

Public Sub SC_Clipboard_SetText(ByVal s As String)
    On Error GoTo KK
    
    DoEvents
    SC_SetCode s
KK:
    If Err.Number = 521 Then Resume 0
End Sub

Public Sub SC_PasteCode(ByVal sCode As String)
    sCode = Right(sCode, Len(sCode) - Len(SmartCopySign))
    CodeOpe.AddCodeToDeclaration sCode
End Sub

Private Sub Class_Initialize()
    SmartCopySign = "iCode SmartCopy V" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
    ReDim m_SC_Buttons(0)
End Sub

Friend Sub SC_ButtonEvent(ByVal btnType As SC_ButtonType)
    Select Case btnType
    Case SC_BT_Cut
        DBPrint "SC_ButtonEvent - Cut"
        
        SmartCopying = True
        VBIns.CommandBars("Edit").Controls("¼ôÇÐ(&T)").Execute
        
        SmartCopying = True
        SC_CopyCutCode True
        
        SmartCopying = False
    Case SC_BT_Copy
        DBPrint "SC_ButtonEvent - Copy"
        
        SmartCopying = True
        VBIns.CommandBars("Edit").Controls("¸´ÖÆ(&C)").Execute
        
        SmartCopying = True
        SC_CopyCutCode
        
        SmartCopying = False
    Case SC_BT_Paste
        DBPrint "SC_ButtonEvent - Paste"
        
        Dim s As String
        s = SC_GetCode
        
        SmartPasting = True
        SC_PasteCode s
        
        SmartPasting = True
        VBIns.CommandBars("Edit").Controls("Õ³Ìù(&P)").Execute
        
        SmartPasting = False
    End Select
End Sub
