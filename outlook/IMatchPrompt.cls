VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMatchPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' IMatchPrompt v0.1.0
'
' The dialogs of the confirmation level (M8): the hint question and the review of the proposals.
' The logic (MatchReview) only talks to this interface; the real dialogs (message box, list form)
' implement it, unit tests use FakeMatchPrompt.
Option Explicit

' "There are new proposals, review them now?" True = yes. pendingCount is the number of waiting
' proposals.
Public Function AskReviewNow(ByVal pendingCount As Long) As Boolean
End Function

' Shows the proposals (Collection of MatchProposal, oldest first) and returns the decision for each
' one: a Collection of Long (MatchDecision values) in the same order. A shorter collection means
' "skip" for the missing ones. Nothing = the user cancelled, nothing is decided.
Public Function Review(ByVal proposals As Collection) As Collection
End Function
