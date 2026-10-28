VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeAuditSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FakeAuditSink v0.1.0
'
' Test double for IAuditSink: keeps the lines in memory. FailOnAppend simulates a broken log file.
Implements IAuditSink
Option Explicit

Public Lines As Collection
Public FailOnAppend As Boolean

Private Sub Class_Initialize()
    Set Lines = New Collection
End Sub

Public Property Get Count() As Long
    Count = Lines.Count
End Property

Public Function LastLine() As String
    If Lines.Count > 0 Then LastLine = CStr(Lines(Lines.Count))
End Function

Private Sub IAuditSink_Append(ByVal line As String)
    If FailOnAppend Then Err.Raise 76, "FakeAuditSink", "path not found (simulated)"
    Lines.Add line
End Sub
