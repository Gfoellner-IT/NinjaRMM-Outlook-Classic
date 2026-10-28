VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MsgBoxMatchPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' MsgBoxMatchPrompt v0.1.0
'
' IMatchPrompt with plain message boxes (dialog variant A of M8): one box for the hint question
' and one box per proposal (Yes = accept, No = reject, Cancel = stop the review). The texts and
' the mapping of the answers are in MatchPromptText; this class only shows the boxes.
Implements IMatchPrompt
Option Explicit

Private Function IMatchPrompt_AskReviewNow(ByVal pendingCount As Long) As Boolean
    IMatchPrompt_AskReviewNow = (MsgBox(MatchPromptText.HintText(pendingCount), vbYesNo + vbQuestion, MatchPromptText.DialogTitle()) = vbYes)
End Function

Private Function IMatchPrompt_Review(ByVal proposals As Collection) As Collection
    Dim decisions As Collection
    Dim i As Long
    Dim answer As VbMsgBoxResult
    Dim stopReview As Boolean

    Set decisions = New Collection
    For i = 1 To proposals.Count
        answer = MsgBox(MatchPromptText.ProposalText(proposals(i), i, proposals.Count), vbYesNoCancel + vbQuestion, _
                        MatchPromptText.DialogTitle())
        decisions.Add CLng(MatchPromptText.DecisionForAnswer(answer, stopReview))
        If stopReview Then Exit For
    Next i

    Set IMatchPrompt_Review = decisions
End Function
