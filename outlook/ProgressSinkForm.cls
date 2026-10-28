VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressSinkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ProgressSinkForm v0.1.0
'
' IProgressSink that shows the non-modal ProgressForm (UserForm.Show vbModeless) via
' ProgressHelper. Outlook stays usable while the form is open; the form is always closed
' in Finish. Used by Ticketautomatismus.RunArchiveRuleWithProgress only.
Implements IProgressSink
Option Explicit

Private Sub IProgressSink_Begin(ByVal total As Long)
    ProgressHelper.ProgressStart total
End Sub

Private Sub IProgressSink_Advance(ByVal message As String)
    ProgressHelper.ProgressStep message
End Sub

Private Sub IProgressSink_Finish()
    ProgressHelper.ProgressEnd
End Sub
