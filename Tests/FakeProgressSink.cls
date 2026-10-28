VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeProgressSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FakeProgressSink v0.1.0
'
' Test double for IProgressSink: records what was reported.
Implements IProgressSink
Option Explicit

Public BeginCount As Long
Public TotalReported As Long
Public FinishCount As Long
Public Messages As Collection

Private Sub Class_Initialize()
    Set Messages = New Collection
End Sub

Private Sub IProgressSink_Begin(ByVal total As Long)
    BeginCount = BeginCount + 1
    TotalReported = total
End Sub

Private Sub IProgressSink_Advance(ByVal message As String)
    Messages.Add message
End Sub

Private Sub IProgressSink_Finish()
    FinishCount = FinishCount + 1
End Sub
