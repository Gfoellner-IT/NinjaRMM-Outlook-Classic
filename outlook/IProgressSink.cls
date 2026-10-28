VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProgressSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' IProgressSink v0.1.0
'
' Receives progress reports of a long run (archiving many ticket folders). The logic in
' TicketService only knows this interface; the visible implementation is ProgressSinkForm
' (non-modal ProgressForm), tests use FakeProgressSink. Unattended runs pass no sink at all.
Option Explicit

' A run with total steps begins.
Public Sub Begin(ByVal total As Long)
End Sub

' One more step is about to be processed.
Public Sub Advance(ByVal message As String)
End Sub

' The run is over (also called when it ended with a problem).
Public Sub Finish()
End Sub
