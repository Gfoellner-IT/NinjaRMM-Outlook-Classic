Attribute VB_Name = "MatchReview"
' MatchReview v0.1.0
'
' The confirmation level (M8): shows the pending proposals through IMatchPrompt and carries out
' the decisions of the user.
'
'   OfferReview    asks "review the new proposals now?" (at most once per MATCH_HINT_MINUTES and only
'                  when there are proposals the user was not asked about yet), then reviews on "yes"
'   ReviewPending  shows all proposals and applies the decisions:
'                    accept  moves the mail (same protection as level 3: open ticket only),
'                            the proposal is dropped even if the move fails (it is stale then)
'                    reject  remembers the mail (MatchRejections) so it is not proposed again
'                    skip    leaves the proposal in the list
'                  Cancelling the dialog decides nothing.
'
' Every decision goes to the audit log (user_confirmed / user_rejected). While a dialog is open no
' second dialog is started (mail events keep arriving). No Outlook types.

Option Explicit

Public Enum MatchDecision
    DecisionSkip = 0
    DecisionAccept = 1
    DecisionReject = 2
End Enum

Private pLastHint As Date
Private pHintedCount As Long
Private pBusy As Boolean

' Test seam: fixed "now" (0 = real time).
Public TestNow As Date

' ---------------------------------------------------------------------------
'  Public API
' ---------------------------------------------------------------------------

' Asks whether to review now, if there is something new to review. Never raises an error.
Public Sub OfferReview(ByVal prompt As IMatchPrompt, ByVal store As ITicketStore)
    Dim pending As Long

    On Error GoTo Failed
    If prompt Is Nothing Then Exit Sub
    If pBusy Then Exit Sub

    pending = MatchProposals.Count()
    If pending = 0 Then
        pHintedCount = 0
        Exit Sub
    End If
    If pending <= pHintedCount Then Exit Sub   ' nothing new since the last question

    If pLastHint <> 0 Then
        If (Clock() - pLastHint) * 1440# < TicketConfig.MATCH_HINT_MINUTES Then Exit Sub
    End If

    pLastHint = Clock()
    pHintedCount = pending

    pBusy = True
    If prompt.AskReviewNow(pending) Then
        pBusy = False
        ReviewPending prompt, store
    End If
    pBusy = False
    Exit Sub

Failed:
    pBusy = False
    TicketLog.Fail "Review offer failed: " & Err.Description
End Sub

' Shows all proposals and applies the decisions. Returns the number of moved mails.
Public Function ReviewPending(ByVal prompt As IMatchPrompt, ByVal store As ITicketStore) As Long
    Dim shown As Collection
    Dim decisions As Collection
    Dim i As Long
    Dim moved As Long
    Dim rejected As Long
    Dim skipped As Long
    Dim failed As Long

    On Error GoTo Failed
    If prompt Is Nothing Then Exit Function
    If pBusy Then Exit Function

    If MatchProposals.Count() = 0 Then
        TicketLog.Info "Review: no pending proposals"
        Exit Function
    End If

    Set shown = New Collection
    For i = 1 To MatchProposals.Count()
        shown.Add MatchProposals.Item(i)
    Next i

    pBusy = True
    Set decisions = prompt.Review(shown)
    pBusy = False

    If decisions Is Nothing Then
        TicketLog.Info "Review cancelled, " & shown.Count & " proposals stay"
        Exit Function
    End If

    For i = 1 To shown.Count
        Select Case DecisionAt(decisions, i)
            Case DecisionAccept
                If TicketService.ConfirmProposal(shown(i), store) Then moved = moved + 1 Else failed = failed + 1
                MatchProposals.RemoveItem shown(i)
            Case DecisionReject
                TicketService.RejectProposal shown(i)
                MatchProposals.RemoveItem shown(i)
                rejected = rejected + 1
            Case Else
                skipped = skipped + 1
        End Select
    Next i

    pHintedCount = MatchProposals.Count()
    TicketLog.Info "Review finished: " & moved & " moved, " & failed & " not moved, " & rejected & " rejected, " & skipped & " skipped"
    ReviewPending = moved
    Exit Function

Failed:
    pBusy = False
    TicketLog.Fail "Review failed: " & Err.Description
End Function

Public Sub ResetState()
    pLastHint = 0
    pHintedCount = 0
    pBusy = False
    TestNow = 0
End Sub

' ---------------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------------

Private Function Clock() As Date
    If TestNow <> 0 Then
        Clock = TestNow
    Else
        Clock = Now
    End If
End Function

' Decision number i, DecisionSkip if it is missing or not a valid value.
Private Function DecisionAt(ByVal decisions As Collection, ByVal i As Long) As MatchDecision
    Dim value As Long

    If i > decisions.Count Then Exit Function
    On Error Resume Next
    value = CLng(decisions(i))
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0

    If value = DecisionAccept Or value = DecisionReject Then DecisionAt = value
End Function
