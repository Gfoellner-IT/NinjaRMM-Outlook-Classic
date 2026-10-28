VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileAuditSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FileAuditSink v0.1.0
'
' IAuditSink that appends to one file per month in
' %LOCALAPPDATA%\NinjaRMM-Outlook-Classic\audit\match_audit_yyyy-mm.jsonl (ASCII, one JSON
' object per line). The first write after Outlook starts also deletes audit files older than
' the retention period. BaseFolder and TestNow exist for unit tests.
Implements IAuditSink
Option Explicit

' Folder of the audit files; "" = the default below %LOCALAPPDATA%.
Public BaseFolder As String
' Months a monthly file is kept (the current month plus this many earlier months).
Public RetentionMonths As Long
' Test seam: fixed "now" (0 = real time).
Public TestNow As Date

Private pPruned As Boolean

Private Sub Class_Initialize()
    RetentionMonths = TicketConfig.AUDIT_RETENTION_MONTHS
End Sub

Public Function Folder() As String
    If Len(BaseFolder) > 0 Then
        Folder = BaseFolder
    Else
        Folder = Environ$("LOCALAPPDATA") & "\NinjaRMM-Outlook-Classic\audit"
    End If
End Function

Private Function Clock() As Date
    If TestNow <> 0 Then
        Clock = TestNow
    Else
        Clock = Now
    End If
End Function

Private Sub IAuditSink_Append(ByVal line As String)
    AppendLine line
End Sub

' Appends the line to the file of the current month (also callable without the interface).
Public Sub AppendLine(ByVal line As String)
    Dim fso As Object
    Dim stream As Object
    Dim path As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolder fso, Folder()
    If Not pPruned Then PruneOldFiles fso

    path = Folder() & "\" & AuditLogic.MonthFileName(Clock())
    Set stream = fso.OpenTextFile(path, 8, True, 0)   ' ForAppending, create, ASCII
    stream.WriteLine line
    stream.Close
End Sub

' Deletes monthly files older than the retention period. Never touches other files.
Public Sub PruneOldFiles(ByVal fso As Object)
    Dim f As Object

    pPruned = True
    On Error Resume Next
    For Each f In fso.GetFolder(Folder()).Files
        If AuditLogic.IsExpiredFileName(f.Name, Clock(), RetentionMonths) Then fso.DeleteFile f.Path, True
    Next f
    On Error GoTo 0
End Sub

Private Sub EnsureFolder(ByVal fso As Object, ByVal path As String)
    Dim parent As String

    If fso.FolderExists(path) Then Exit Sub
    parent = fso.GetParentFolderName(path)
    If Len(parent) > 0 Then EnsureFolder fso, parent
    fso.CreateFolder path
End Sub
