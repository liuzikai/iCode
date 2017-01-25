VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_lProc As Long

Private Enum enmSortMode
    smLine = 0
    smRect = 1
End Enum

Private Enum enmStructure
    sIf = 11
    mIf = 31
    eIf = 21
    sSelect = 12
    mSelect = 32
    eSelect = 22
    sWith = 13
    eWith = 23
    sFor = 14
    eFor = 24
    sProc = 15
    eProc = 25
    sType = 16
    etype = 26
    sEnum = 17
    eEnum = 27
    sDo = 18
    eDo = 28
End Enum

Dim s As New CStack

Public Sub DealMessage(ByVal Msg As Long, ByVal wParam As Long, ByVal Caption As String, ByVal ClassName As String)
    
    Select Case Msg
        
    Case WM_KEYUP
        
        Select Case wParam
            
        Case vbKeyReturn
            If iState = isCode Then
                If iCode.AC.Visible = False Then
                    Me.SortMouseKeyEvent
                End If
            End If
            
        End Select
        
    Case WM_LBUTTONUP
        
        If iState = isCode Then iCode.CodeSort.SortMouseKeyEvent
        
    End Select
    
End Sub

Friend Property Get lProc() As Long
    lProc = m_lProc
End Property

Friend Property Let lProc(Value As Long)
    m_lProc = Value
End Property

Public Function SortProc(ByVal sProc As String, ByVal kProc As vbext_ProcKind) As Long
    DBPN "SortProc"
    With VBIns.ActiveCodePane.CodeModule
        If sProc <> "" Then
            SortLines .ProcBodyLine(sProc, kProc), .ProcStartLine(sProc, kProc) + .ProcCountLines(sProc, kProc) - 1
            SortProc = .ProcStartLine(sProc, kProc) + .ProcCountLines(sProc, kProc) - 1
        Else
            SortLines 1, .CountOfDeclarationLines
            SortProc = .CountOfDeclarationLines
        End If
        
    End With
End Function


Public Function SortLines(lStart As Long, lEnd As Long) As Boolean
    
    On Error GoTo iErr
    
    
    Dim i As Long
    Dim r As enmStructure, v As enmStructure
    
    CodeOpe.UpdataSelectionInfo
    
    For i = lStart To lEnd
        
        SortLine i
        
    Next
    
    
    VBIns.ActiveCodePane.SetSelection CodeOpe.SL, CodeOpe.SC, CodeOpe.EL, CodeOpe.EC
    
    SortLines = True
    
    Exit Function
iErr:
    SortLines = False
    DBErr "iCode - Function SortLines", "Code :" & vbCrLf & VBIns.ActiveCodePane.CodeModule.Lines(1, VBIns.ActiveCodePane.CodeModule.CountOfLines), "lStart = " & lStart, "lEnd = " & lEnd
End Function

Public Function SortString(ByVal sLine As String, ByVal i As Long, ByVal c As Long, Optional ByVal AppointLevel As Long = 0) As Long
    
    On Error GoTo iErr
    
    Dim r As enmStructure, v As enmStructure
    
    With VBIns.ActiveCodePane.CodeModule
        
        
        r = CheckCode(sLine)
        
        If r = 0 Then
            If AppointLevel = 0 Then
                SetLine i, s.Count, c
                SortString = s.Count
            ElseIf AppointLevel = -1 Then
                
            Else
                SetLine i, AppointLevel, c
                SortString = AppointLevel
            End If
        ElseIf (r > 10 And r < 20) Then
            If AppointLevel = 0 Then
                SetLine i, s.Count, c
                SortString = s.Count
            ElseIf AppointLevel = -1 Then
                
            Else
                SetLine i, AppointLevel, c
                SortString = AppointLevel
            End If
            
            s.Push r
        ElseIf (r > 20 And r < 30) Then
            
            If r = eProc Then
                Do Until s.Pop(v) = False
                Loop
                If AppointLevel <> -1 Then SetLine i, 0, c
            Else
                Do
                    If (s.Count <= 1 And s.First = sProc) Or s.Count <= 0 Then
                        If AppointLevel = 0 Then
                            SetLine i, s.Count - 1, c
                            SortString = s.Count
                        ElseIf AppointLevel = -1 Then
                            
                        Else
                            SetLine i, AppointLevel, c
                            SortString = AppointLevel
                        End If
                        
                        Exit Do
                    End If
                    
                    s.Pop v
                    
                    DoEvents
                    
                    If (v Mod 10) = (r Mod 10) Then
                        If AppointLevel = 0 Then
                            SetLine i, s.Count, c
                            SortString = s.Count
                        ElseIf AppointLevel = -1 Then
                            
                        Else
                            SetLine i, AppointLevel, c
                            SortString = AppointLevel
                        End If
                        
                        Exit Do
                    End If
                    
                Loop
            End If
            
            
        ElseIf (r > 30) Then
            Do
                If (s.Count <= 1 And s.First = sProc) Or s.Count <= 0 Then
                    If AppointLevel = 0 Then
                        SetLine i, s.Count - 1, c
                        SortString = s.Count
                    ElseIf AppointLevel = -1 Then
                        
                    Else
                        SetLine i, AppointLevel, c
                        SortString = AppointLevel
                    End If
                    
                    Exit Do
                End If
                
                s.Pop v
                DoEvents
                
                If (v Mod 10) = (r Mod 10) Then
                    If AppointLevel = 0 Then
                        SetLine i, s.Count, c
                        SortString = s.Count
                    ElseIf AppointLevel = -1 Then
                        
                    Else
                        SetLine i, AppointLevel, c
                        SortString = AppointLevel
                    End If
                    
                    Exit Do
                End If
            Loop
            
            s.Push r
        End If
        
        
    End With
    
    
    Exit Function
iErr:
    SortString = -1
    DBErr "iCode - Function SortString", "sLine = " & sLine
End Function

Private Function SortLine(ByRef n As Long) As Boolean
    SortLine = False
    
    On Error GoTo iErr
    
    Dim r As String, c As Long
    
    c = 0
    
    With VBIns.ActiveCodePane.CodeModule
        
        Dim i As Long, bNoFirstOfLine As Boolean, nFirstLineLevel As Long
        
        i = 0: bNoFirstOfLine = False: nFirstLineLevel = -1
        
        
        Do
            If c = 0 Then
                r = r & .Lines(n, 1)
            Else
                r = r & LTrim(.Lines(n, 1))
            End If
            
            c = c + 1
            
            
            Do
                i = InStr(1, r, ":")
                
                If i = 0 Then
                    Exit Do
                Else
                    If bNoFirstOfLine = False Then
                        If nFirstLineLevel = -1 Then
                            nFirstLineLevel = SortString(Left(r, i - 1), n, 1, 0)
                        Else
                            SortString Left(r, i - 1), n, 1, nFirstLineLevel + 1
                        End If
                        bNoFirstOfLine = True
                    Else
                        SortString Left(r, i - 1), n, 1, -1
                    End If
                    
                    r = Right(r, Len(r) - i)
                    c = 1
                    
                End If
            Loop
            
            r = RTrim(r)
            
            
            If Right(r, 1) <> "_" Then
                If bNoFirstOfLine = False Then
                    If nFirstLineLevel = -1 Then
                        nFirstLineLevel = SortString(r, n, c, 0)
                    Else
                        SortString r, n, 1, nFirstLineLevel + 1
                    End If
                Else
                    SortString r, n, 1, -1
                End If
                
                Exit Do
            Else
                r = Left(r, Len(r) - 1)
                n = n + 1
                bNoFirstOfLine = False
            End If
            
        Loop
    End With
    
    SortLine = True
    
    Exit Function
iErr:
    SortLine = False
    DBErr "iCode - Function SortLine", "Code :" & vbCrLf & VBIns.ActiveCodePane.CodeModule.Lines(1, VBIns.ActiveCodePane.CodeModule.CountOfLines), "n = " & n
    
End Function

Private Function SetLine(ByVal nLine As Long, ByVal nLevel As Long, ByVal lCount As Long) As Boolean
    '这里需要引入lCount以对多行代码块进行整理
    
    If nLevel < 0 Then nLevel = 0
    
    Dim i As Long
    
    For i = nLine - lCount + 1 To nLine Step 1
        
        Dim l As String
        
        l = VBIns.ActiveCodePane.CodeModule.Lines(i, 1)
        '这里需要使用Lines,因为使用换行符的代码块需要每行都进行整理
        
        Dim n As Long, m As Long
        
        m = 0
        
        For n = 1 To Len(l)
            If Mid(l, n, 1) <> " " Then Exit For
            m = m + 1
        Next
        
        If m = nLevel * 4 Then GoTo iGoNext
        
iGoSort:
        
        VBIns.ActiveCodePane.CodeModule.ReplaceLine i, Space(nLevel * 4) & LTrim(l)
        
        If nLine = CodeOpe.SL Then
            CodeOpe.SC = CodeOpe.SC + (nLevel * 4 - m)
            If CodeOpe.SC < 0 Then CodeOpe.SC = 0
        End If
        
        If nLine = CodeOpe.EL Then
            CodeOpe.EC = CodeOpe.EC + (nLevel * 4 - m)
            If CodeOpe.EC < 0 Then CodeOpe.EC = 0
        End If
        
iGoNext:
        
        If i = nLine - lCount + 1 Then nLevel = nLevel + 1
        '增加nLevel使得由多行组成的代码块除第一行外向后位移
        
    Next
    
End Function

Private Function CheckCode(sLine As String) As enmStructure
    Dim s As String
    s = Trim(sLine)
    If LeftIs(s, "If ") Then
        Dim ts As String
        ts = LTrim(Right(s, Len(s) - (InStr(1, LCase(s), LCase("Then")) + 4) + 1))
        If ts = "" Or LeftIs(ts, "'") Then
            CheckCode = sIf
        Else
            CheckCode = 0
        End If
    ElseIf LCase(s) = LCase("Else") Or LeftIs(s, "ElseIf ") Then
        CheckCode = mIf
    ElseIf LeftIs(s, "End If") Or LeftIs(s, "EndIf ") Or LeftIs(s, "EndIf'") Then
        CheckCode = eIf
    ElseIf LeftIs(s, "Select Case ") Then
        CheckCode = sSelect
    ElseIf LeftIs(s, "Case ") Then
        CheckCode = mSelect
    ElseIf LeftIs(s, "End Select") Then
        CheckCode = eSelect
    ElseIf LeftIs(s, "With ") Then
        CheckCode = sWith
    ElseIf LeftIs(s, "End With") Then
        CheckCode = eWith
    ElseIf LeftIs(s, "For ") Then
        CheckCode = sFor
    ElseIf LCase(s) = LCase("Next") Or LeftIs(s, "Next ") Or LeftIs(s, "Next'") Then
        CheckCode = eFor
    ElseIf LeftIs(s, "Private Sub ") Or LeftIs(s, "Private Function ") Or LeftIs(s, "Private Property ") _
        Or LeftIs(s, "Public Sub ") Or LeftIs(s, "Public Function ") Or LeftIs(s, "Public Property ") _
        Or LeftIs(s, "Friend Sub ") Or LeftIs(s, "Friend Function ") Or LeftIs(s, "Friend Property ") _
        Or LeftIs(s, "Sub ") Or LeftIs(s, "Function ") Or LeftIs(s, "Property ") Then
        CheckCode = sProc
    ElseIf LeftIs(s, "End Sub") Or LeftIs(s, "End Function") Or LeftIs(s, "End Property") Then
        CheckCode = eProc
    ElseIf LeftIs(s, "Private Type ") Or LeftIs(s, "Public Type ") Or LeftIs(s, "Firend Type ") Or LeftIs(s, "Type ") Then
        CheckCode = sType
    ElseIf LeftIs(s, "End Type") Then
        CheckCode = etype
    ElseIf LeftIs(s, "Private Enum ") Or LeftIs(s, "Public Enum ") Or LeftIs(s, "Firend Enum ") Or LeftIs(s, "Enum ") Then
        CheckCode = sEnum
    ElseIf LeftIs(s, "End Enum") Then
        CheckCode = eEnum
    ElseIf LeftIs(s, "Do ") Or LCase(s) = LCase("Do") Or LeftIs(s, "While ") Or LCase(s) = LCase("While") Then
        CheckCode = sDo
    ElseIf LeftIs(s, "Loop ") Or LCase(s) = LCase("Loop") Or LeftIs(s, "Wend ") Or LCase(s) = LCase("Wend") Then
        CheckCode = eDo
    End If
End Function

Public Function GetActiveLine() As Long
    Dim n As Long
    VBIns.ActiveCodePane.GetSelection n, 0, 0, 0
    GetActiveLine = n
End Function

Public Function SortMouseKeyEvent() As Long
    If lProc = GetActiveLine Then Exit Function
    
    Dim ProcOld As String, ProcNew As String
    Dim KOld As vbext_ProcKind, KNew As vbext_ProcKind
    
    If lProc > 0 Then CodeOpe.GetProcByLine lProc, ProcOld, KOld
    
    CodeOpe.GetProcByLine GetActiveLine, ProcNew, KNew
    
    
    
    SortProc ProcNew, KNew
    
    If ProcOld <> ProcNew Or KOld <> KNew Then
        iCode.CWE.ProcNotesAdd ProcNew, KNew
        SortProc ProcOld, KOld
    End If
    
    lProc = GetActiveLine
    
End Function

Public Function SortPasteEvent() As Long
    If lProc = GetActiveLine Then Exit Function
    
    Dim ProcOld As String, ProcNew As String
    Dim KOld As vbext_ProcKind, KNew As vbext_ProcKind
    
    Dim eLine As Long
    
    eLine = GetActiveLine
    
    If lProc > 0 Then CodeOpe.GetProcByLine lProc, ProcOld, KOld
    CodeOpe.GetProcByLine eLine, ProcNew, KNew
    
    Dim s As Long, e As Long
    
    If ProcOld <> "" Then
        s = VBIns.ActiveCodePane.CodeModule.ProcBodyLine(ProcOld, KOld)
    Else
        s = 1
    End If
    
    If ProcNew <> "" Then
        e = VBIns.ActiveCodePane.CodeModule.ProcStartLine(ProcNew, KNew) + VBIns.ActiveCodePane.CodeModule.ProcCountLines(ProcNew, KNew) - 1
    Else
        e = VBIns.ActiveCodePane.CodeModule.CountOfDeclarationLines
    End If
    
    If s > e Then
        Dim t As Long
        t = s: s = e: e = t
    End If
    
    SortLines s, e
    
    lProc = eLine
End Function


