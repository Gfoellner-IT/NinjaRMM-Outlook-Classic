VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAuditSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' IAuditSink v0.1.0
'
' Where audit lines go. The real implementation is FileAuditSink, unit tests use FakeAuditSink.
Option Explicit

' Appends one finished audit line (ASCII, no line break). May raise an error; TicketAudit
' swallows it.
Public Sub Append(ByVal line As String)
End Sub
