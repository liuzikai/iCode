VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeQuick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const filename As String = "\CodeQuick.iCode"

Private WithEvents Buttons As ButtonsCollection
Attribute Buttons.VB_VarHelpID = -1

Private WithEvents SC As ScriptControl
Attribute SC.VB_VarHelpID = -1

Public Enum CQ_Proc_Type
    tFind = 1
    tReplace = 2
    tButton = 3
    tInstant = 4
End Enum

Public Type CQ_Proc_Info
    Caption As String
    ProcName As String
    ProcType As Long
    Code As String
    CodeJudge As String
    Tag As String
End Type

Public Type CQ_ButtonsInfo
    Proc As CQ_Proc_Info
    ButtonIndex As Long
    BeginGroup As Boolean
    Tag As String
End Type

Public Type CQ
    Find() As CQ_Proc_Info
    Replace() As CQ_Proc_Info
    Buttons() As CQ_Proc_Info
    Instant() As CQ_Proc_Info
    CButtons() As CQ_ButtonsInfo
    CButtonsCount As Long
End Type

Private r As CQ

Private WithEvents BC As ButtonsCollection
Attribute BC.VB_VarHelpID = -1

Private Sub BC_ButtonClick(ByVal CommandBarControl As Office.CommandBarControl)
    RunProc Right(CommandBarControl.Tag, Len(CommandBarControl.Tag) - 1), Left(CommandBarControl.Tag, 1)
End Sub

Public Sub JudegeButtons()
    Dim i As Long
    For i = 1 To r.CButtonsCount
        With r.CButtons(i)
            If .Proc.ProcType = tButton Then
                BC.Items(.ButtonIndex).Button.Enabled = SC.Modules("ButtonsJudge").Run(.Proc.ProcName)
            End If
        End With
    Next
End Sub

Private Function RunProc(ByVal ProcName As String, ByVal ProcType As CQ_Proc_Type) As Variant
    RunProc = SC.Modules(Choose(ProcType, "Find", "Replace", "ButtonsAction", "Instant")).Run(ProcName)
End Function

Private Sub Class_Initialize()
    Set SC = frmHidden.CQ_SC
    Set BC = New ButtonsCollection
    LoadRecords
End Sub

Private Sub LoadRecords()
    If Dir(App.Path & filename) <> "" Then
        Dim fn As Long: fn = FreeFile
        Open App.Path & filename For Binary As #fn
        Get #fn, , r
        Close #fn
    Else
        If Dir(App.Path & filename) = "" Then
            r = GetDefultRecords
            SetRecords
        End If
    End If
    
    BC.AddTargetBar "CodeQucik", True
    Dim i As Long
    For i = 1 To r.CButtonsCount
        With r.CButtons(i)
            BC.Add msoControlButton, .Proc.Caption, msoButtonCaption, .BeginGroup, CStr(.Proc.ProcType) & .Proc.ProcName, .ButtonIndex
        End With
    Next
    
    SC.AddObject "CodeOpe", CodeOpe, True
    
    SC.Modules.Add "Find"
    For i = 1 To UBound(r.Find)
        SC.Modules("Find").AddCode r.Find(i).Code
    Next
    
    SC.Modules.Add "ButtonsJudge"
    SC.Modules.Add "ButtonsAction"
    For i = 1 To UBound(r.Buttons)
        SC.Modules("ButtonsJudge").AddCode r.Buttons(i).CodeJudge
        SC.Modules("ButtonsAction").AddCode r.Buttons(i).Code
    Next
    
    SC.Modules.Add "Instant"
    For i = 1 To UBound(r.Instant)
        SC.Modules("Instant").AddCode r.Instant(i).Code
    Next
End Sub

Function Private_Public() As Boolean
    
End Function

Public Function SetRecords()
    Dim fn As Long: fn = FreeFile
    Open App.Path & filename For Binary As #fn
    Put #fn, , r
    Close #fn
End Function

Public Function GetDefultRecords() As CQ
    Dim t As CQ
    ReDim t.Find(0)
    ReDim t.Replace(0)
    ReDim t.Buttons(0)
    ReDim t.Instant(0)
    
    
    ReDim Preserve t.Buttons(1)
    
    With t.Buttons(1)
        .Caption = "Private¡û¡úPublic"
        .ProcName = "Private_Public"
        .ProcType = tButton
        
        .CodeJudge = _
        "Function Private_Public()" & vbCrLf & _
        "    Private_Public = True" & vbCrLf & _
        "End Function"
        
        .Code = Clipboard.GetText
    End With
    
    ReDim t.CButtons(1)
    t.CButtonsCount = 1
    t.CButtons(1).Proc = t.Buttons(1)
    
    GetDefultRecords = t
End Function

'Sub Private_Public_DealString(ByRef s)
'    Do Until InStr(1, s, "Private") = 0
'        s = Replace(s, "Private", "#3607#")
'    Loop
'    Do Until InStr(1, s, "Public") = 0
'        s = Replace(s, "Public", "Private")
'    Loop
'    Do Until InStr(1, s, "#3607#") = 0
'        s = Replace(s, "#3607#", "Public")
'    Loop
'End Sub
'
'Sub Private_Public()
'    Dim s
'
'    If CodeOpe.Selection <> "" Then
'
'        s = CodeOpe.Selection
'
'        Private_Public_DealString s
'
'        CodeOpe.Selection = s
'
'    Else
'        If CodeOpe.bInDeclaration = True Then
'            s = CodeOpe.Lines(CodeOpe.SL)
'            Private_Public_DealString s
'            CodeOpe.ReplaceLine CodeOpe.SL, s
'        Else
'            Dim PN
'            PN = CodeOpe.TargetCodePane.CodeModule.ProcOfLine(CodeOpe.SL, CodeOpe.m)
'            Dim PSL
'            PSL = CodeOpe.TargetCodePane.CodeModule.ProcBodyLine(PN, CodeOpe.m)
'            s = CodeOpe.Lines(PSL)
'            Private_Public_DealString s
'            CodeOpe.ReplaceLine PSL, s
'        End If
'    End If
'End Sub


