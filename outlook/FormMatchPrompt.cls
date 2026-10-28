VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormMatchPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FormMatchPrompt v0.1.0
'
' IMatchPrompt with the review form ReviewForm (dialog variant B of M8: a list with check boxes). The
' form is created by name at run time (VBA.UserForms.Add), so this class compiles without the form;
' if the form is not part of the VBA project, the proposals are shown with message boxes instead
' (MsgBoxMatchPrompt, variant A) and a warning is logged once. The hint question is the same box in
' both variants.
Implements IMatchPrompt
Option Explicit

Private pWarned As Boolean
Private pFallback As MsgBoxMatchPrompt

Private Function IMatchPrompt_AskReviewNow(ByVal pendingCount As Long) As Boolean
    IMatchPrompt_AskReviewNow = (MsgBox(MatchPromptText.HintText(pendingCount), vbYesNo + vbQuestion, MatchPromptText.DialogTitle()) = vbYes)
End Function

Private Function IMatchPrompt_Review(ByVal proposals As Collection) As Collection
    Dim form As Object
    Dim fallback As IMatchPrompt

    On Error Resume Next
    Set form = VBA.UserForms.Add("ReviewForm")
    If Err.Number <> 0 Then Set form = Nothing
    Err.Clear
    On Error GoTo 0

    If form Is Nothing Then
        If Not pWarned Then
            pWarned = True
            TicketLog.Warn "The review form ReviewForm is not in the VBA project: message boxes are used instead"
        End If
        If pFallback Is Nothing Then Set pFallback = New MsgBoxMatchPrompt
        Set fallback = pFallback
        Set IMatchPrompt_Review = fallback.Review(proposals)
        Exit Function
    End If

    form.Setup proposals
    form.Show vbModal
    Set IMatchPrompt_Review = form.Result()
    Unload form
End Function
