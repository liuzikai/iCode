Attribute VB_Name = "modDebug"
Option Explicit

Private Const ErrLogFileName As String = "\Error.log"
Private ErrLogFN As Long

Private Const PNFileName As String = "\FullLog.mdb"
Dim PNCnn As New ADODB.Connection
Dim PNRS As New ADODB.Recordset

Public Sub DBInit()
    ErrLogFN = FreeFile
    If Dir(Left(ErrLogFileName, InStrRev(ErrLogFileName, "\") - 1), vbDirectory) = "" Then MkDir App.Path & Left(ErrLogFileName, InStrRev(ErrLogFileName, "\") - 1)
    Open App.Path & ErrLogFileName For Output As #ErrLogFN
    
    PNOpenData
End Sub

Public Sub DBUnLoad()
    If ErrLogFN <> 0 Then Close #ErrLogFN
    
    PNCloseData
End Sub

Public Sub DBErr(ByVal Source As String, ParamArray Extra())
    On Error Resume Next
    
    If Err.Number = 0 Then Exit Sub
    Err.Source = Source
    
    Print #ErrLogFN, "Time: " & Time
    Print #ErrLogFN, "Source: " & Err.Source
    Print #ErrLogFN, "Number: " & Err.Number
    Print #ErrLogFN, "Description: " & Err.description
    Print #ErrLogFN, "Extra Info:"
    Dim i As Long
    
    For i = 0 To UBound(Extra)
        Print #ErrLogFN, "     " & Extra(i)
    Next
End Sub

Public Sub DBPN(ByVal ProcName As String)
    '    PNRS.AddNew
    '    PNRS.Fields("Time") = Now
    '    PNRS.Fields("Proc") = ProcName
End Sub

Private Sub PNOpenData()
    'PNCnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & PNFileName
    'PNRS.CursorLocation = adUseClient
    
    'PNCnn.Execute "Create Table " & Now & " (Time Data,Proc BSTR)"
    
    'PNRS.Open "FirstTabel", PNCnn, 3, 3
    'PNRS.Fields.Append "Time", adDate
    'PNRS.Fields.Append "Proc", adBSTR
    
End Sub

Private Sub PNCloseData()
    'PNRS.Close
    'PNCnn.Close
End Sub

Public Sub DBEnd()
    UnSetCBTHook
    UnSetMsgHooks
    Set iTipsBar = Nothing
    XTimerSupport.Scrub
End Sub
