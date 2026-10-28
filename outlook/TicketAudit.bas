Attribute VB_Name = "TicketAudit"
' TicketAudit v0.1.0
'
' Audit log of the match decisions (M7): one JSON line per decision through an IAuditSink.
' It records features only (sender domain, normalized subject, ticket numbers, hashed
' conversation and message ids), never mail contents or person names.
' A problem with the audit log never disturbs the ticket automation: every error is swallowed
' (noted once in TicketLog). No Outlook types; mails come in through IIncomingMail.

Option Explicit

Private pSink As IAuditSink
Private pFailureLogged As Boolean

' ---------------------------------------------------------------------------
'  Public API
' ---------------------------------------------------------------------------

' Records the decision about one foreign mail.
'   source  watcher | sweep | dry_run
'   action  log_only | proposed | moved | stays | none
Public Sub RecordMatchDecision(ByVal mail As IIncomingMail, ByVal outcome As MatchOutcome, ByVal rule As MatchRule, _
                               ByVal ticketId As Long, ByVal action As String, ByVal source As String)
    WriteLine mail, "auto_decision", outcome, rule, ticketId, action, source
End Sub

' Records what the user decided about a proposal (M8): accepted = True is user_confirmed (action
' moved or stays), False is user_rejected (action rejected).
' Records a mail the user moved into a ticket folder by hand (M7). outcome, rule and matcherTicket tell what
' the matcher would have done with it; action is agree (same ticket), differ (another ticket) or missed
' (no or no unambiguous match). ticketId is the ticket of the folder.
Public Sub RecordManualAssignment(ByVal mail As IIncomingMail, ByVal outcome As MatchOutcome, ByVal rule As MatchRule, _
                                  ByVal ticketId As Long, ByVal matcherTicket As Long, ByVal action As String)
    WriteLine mail, "manual_assignment", outcome, rule, ticketId, action, "scan", matcherTicket
End Sub

Public Sub RecordUserDecision(ByVal mail As IIncomingMail, ByVal rule As MatchRule, ByVal ticketId As Long, _
                              ByVal accepted As Boolean, ByVal action As String)
    WriteLine mail, IIf(accepted, "user_confirmed", "user_rejected"), MatchFound, rule, ticketId, action, "review"
End Sub

Private Sub WriteLine(ByVal mail As IIncomingMail, ByVal eventType As String, ByVal outcome As MatchOutcome, _
                      ByVal rule As MatchRule, ByVal ticketId As Long, ByVal action As String, ByVal source As String, _
                      Optional ByVal matcherTicket As Long = 0)
    Dim sink As IAuditSink

    On Error GoTo Failed

    Set sink = ActiveSink()
    If sink Is Nothing Then Exit Sub

    sink.Append BuildLine(mail, eventType, outcome, rule, ticketId, action, source, matcherTicket)
    Exit Sub

Failed:
    If Not pFailureLogged Then
        pFailureLogged = True
        TicketLog.Warn "Audit log: could not write (" & Err.Description & "); further problems are not logged"
    End If
End Sub

' The line RecordMatchDecision would write (public for unit tests).
Public Function BuildLine(ByVal mail As IIncomingMail, ByVal eventType As String, ByVal outcome As MatchOutcome, ByVal rule As MatchRule, _
                          ByVal ticketId As Long, ByVal action As String, ByVal source As String, _
                          Optional ByVal matcherTicket As Long = 0) As String
    Dim subjectText As String
    Dim fullSubject As String
    Dim normalized As String
    Dim ninjaIds As Collection
    Dim externals As Collection

    subjectText = mail.Subject
    ' Privacy: the subject itself is only written for mails the matcher found a ticket for (or was unsure
    ' about) and for decisions of the user; for mails without a match only its hash is logged.
    fullSubject = TicketMatcher.NormalizeSubject(subjectText)
    If TicketConfig.AUDIT_LOG_SUBJECT And (outcome <> MatchNone Or eventType = "manual_assignment" Or TicketConfig.AUDIT_LOG_SUBJECT_ON_NONE) Then
        normalized = Left$(fullSubject, TicketConfig.AUDIT_SUBJECT_MAX_CHARS)
    End If

    Set ninjaIds = TicketMatcher.ExtractNinjaIds(subjectText)
    If ninjaIds.Count = 0 Then Set ninjaIds = TicketMatcher.ExtractNinjaIds(mail.BodyText)
    Set externals = TicketMatcher.ExtractExternalNumbers(subjectText)

    BuildLine = AuditLogic.BuildMatchEvent(Now, eventType, source, OutcomeName(outcome), TicketMatcher.RuleName(rule), ticketId, action, _
                                           mail.SenderDomain, normalized, HashOrEmpty(fullSubject), JoinCollection(externals), _
                                           JoinCollection(ninjaIds), HashOrEmpty(mail.ConversationId), _
                                           mail.MessageKey, mail.HasAutoHeader, matcherTicket)
End Function

' ---------------------------------------------------------------------------
'  Test seams
' ---------------------------------------------------------------------------

Public Sub UseSink(ByVal sink As IAuditSink)
    Set pSink = sink
End Sub

Public Sub ResetState()
    Set pSink = Nothing
    pFailureLogged = False
End Sub

' ---------------------------------------------------------------------------
'  Implementation
' ---------------------------------------------------------------------------

' The injected sink, otherwise the file sink. Nothing if the audit log is switched off, and
' while file logging is disabled (unit tests), so tests never write into the real audit folder.
Private Function ActiveSink() As IAuditSink
    If Not pSink Is Nothing Then
        Set ActiveSink = pSink
        Exit Function
    End If

    If Not TicketConfig.ENABLE_AUDIT_LOG Then Exit Function
    If TicketLog.IsFileLoggingDisabled() Then Exit Function

    Set pSink = New FileAuditSink
    Set ActiveSink = pSink
End Function

Private Function OutcomeName(ByVal outcome As MatchOutcome) As String
    Select Case outcome
        Case MatchFound:      OutcomeName = "found"
        Case MatchAmbiguous:  OutcomeName = "ambiguous"
        Case Else:            OutcomeName = "none"
    End Select
End Function

Private Function HashOrEmpty(ByVal text As String) As String
    If Len(text) > 0 Then HashOrEmpty = AuditLogic.Hash32Hex(text)
End Function

' "a,b,c" (empty string for an empty collection).
Private Function JoinCollection(ByVal items As Collection) As String
    Dim i As Long

    For i = 1 To items.Count
        If i > 1 Then JoinCollection = JoinCollection & ","
        JoinCollection = JoinCollection & CStr(items(i))
    Next i
End Function
