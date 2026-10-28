Attribute VB_Name = "TicketService"
' TicketService v0.2.0
'
' The ticket automation without Outlook types: sorts ticket mails into ticket folders
' and archives closed tickets. Folders and mails come in through ITicketStore, ITicketFolder
' and IIncomingMail, ticket data through NinjaApi. This makes the whole logic unit testable.
' The Outlook entry points are in Ticketautomatismus.bas.

Option Explicit

' After this many consecutive "unavailable" answers the API is not asked again during one
' archive run (an offline PC would otherwise wait for timeouts once per folder).
Private Const API_DOWN_AFTER_FAILURES As Long = 3

' ---------------------------------------------------------------------------
'  New mail
' ---------------------------------------------------------------------------

' Moves a ticket mail into its ticket folder (reusing an active folder, restoring an archived
' one or creating a new one) and aligns the folder name with the ticket subject.
' Returns True if the mail was a ticket mail and has been moved.
Public Function ProcessMail(ByVal mail As IIncomingMail, ByVal store As ITicketStore) As Boolean
    Dim ticketNumber As String
    Dim remainingText As String
    Dim folder As ITicketFolder
    Dim ticketId As Long

    On Error GoTo Failed

    If mail.SenderName <> TicketConfig.AllowedSender() Then Exit Function
    If Not TicketParsing.ParseSubject(mail.Subject, ticketNumber, remainingText) Then Exit Function
    If Len(ticketNumber) > 10 Then Exit Function   ' not a real ticket number (would overflow a Long)
    ticketId = CLng(Mid$(ticketNumber, 2))

    Set folder = GetOrCreateTicketFolder(store, ticketNumber, remainingText)

    store.MoveMailToFolder mail, folder
    RenameToSubject folder, ticketId, NinjaApi.GetTicketSubject(ticketId)
    ProcessMail = True
    Exit Function

Failed:
    TicketLog.Fail "Processing a ticket mail failed: " & Err.Description
End Function

' ---------------------------------------------------------------------------
'  Foreign mails (M6)
' ---------------------------------------------------------------------------

' Tries to assign a mail that does NOT come from the ticket sender to an existing ticket
' (TicketMatcher) and logs the decision. What happens next depends on the safety level of the rule that
' matched (MatchLevels): log only, queue a proposal (MatchProposals) or move the mail. Only an
' unambiguous match with an open ticket is ever moved or proposed; otherwise the mail stays in the
' inbox. allowActions = False (dry run) never proposes or moves, whatever the level. Returns the match
' outcome (MatchNone for mails of the ticket sender, which are none of this function's business).
Public Function MatchForeignMail(ByVal mail As IIncomingMail, ByVal index As ITicketIndex, ByVal store As ITicketStore, _
                                 ByVal allowActions As Boolean, Optional ByVal source As String = "watcher") As MatchOutcome
    Dim ticketId As Long
    Dim rule As MatchRule
    Dim outcome As MatchOutcome
    Dim line As String
    Dim action As String
    Dim level As MatchLevel

    On Error GoTo Failed

    If mail.SenderName = TicketConfig.AllowedSender() Then Exit Function
    If Not MatchLevels.AnyRuleActive() Then Exit Function

    outcome = TicketMatcher.MatchMail(mail.Subject, mail.BodyText, mail.ConversationId, mail.SenderDomain, _
                                      mail.HasAutoHeader, index, ticketId, rule)
    line = TicketMatcher.DescribeMatch(outcome, rule, ticketId, mail.SenderDomain, mail.Subject)

    If outcome <> MatchFound Then
        TicketLog.Info line
        TicketAudit.RecordMatchDecision mail, outcome, rule, ticketId, "none", source
        MatchForeignMail = outcome
        Exit Function
    End If

    ' Later mails of the thread follow, but only after a decisive rule (not after a subject guess).
    If rule = RuleNinjaId Or rule = RuleExternalNumber Then index.RememberConversation mail.ConversationId, ticketId

    level = MatchLevels.LevelFor(rule)
    If Not allowActions Or level <= LevelLogOnly Then
        action = "log_only"
        TicketLog.Info line & " (log only)"
    ElseIf level = LevelConfirm Then
        If QueueProposal(mail, index, ticketId, rule) Then
            action = "proposed"
            TicketLog.Info line & " (proposed, waits for confirmation)"
        Else
            action = "stays"
            TicketLog.Info line & " (stays in the inbox)"
        End If
    ElseIf MoveToTicket(mail, store, ticketId) Then
        action = "moved"
        TicketLog.Info line & " (moved)"
    Else
        action = "stays"
        TicketLog.Info line & " (stays in the inbox)"
    End If
    TicketAudit.RecordMatchDecision mail, outcome, rule, ticketId, action, source
    MatchForeignMail = outcome
    Exit Function

Failed:
    TicketLog.Fail "Matching a mail failed: " & Err.Description
End Function

' The user accepted a proposal: moves the mail into the folder of the (still open) ticket and writes
' the decision to the audit log. Returns True if the mail was moved. A closed ticket, an unreadable
' ticket or a mail that cannot be moved any more (moved or deleted meanwhile) leave the mail alone.
Public Function ConfirmProposal(ByVal proposal As MatchProposal, ByVal store As ITicketStore) As Boolean
    Dim moved As Boolean

    On Error GoTo Failed

    moved = MoveToTicket(proposal.Mail, store, proposal.TicketId)
    TicketAudit.RecordUserDecision proposal.Mail, proposal.Rule, proposal.TicketId, True, IIf(moved, "moved", "stays")
    If moved Then
        TicketLog.Info "Review: mail moved to ticket #" & proposal.TicketId
    Else
        TicketLog.Info "Review: mail not moved to ticket #" & proposal.TicketId & " (ticket not open or not readable)"
    End If
    ConfirmProposal = moved
    Exit Function

Failed:
    TicketLog.Fail "Review: the mail could not be moved to ticket #" & proposal.TicketId & ": " & Err.Description
    TicketAudit.RecordUserDecision proposal.Mail, proposal.Rule, proposal.TicketId, True, "stays"
End Function

' The user rejected a proposal: remembers the mail (it is not proposed again) and writes the
' decision to the audit log.
Public Sub RejectProposal(ByVal proposal As MatchProposal)
    On Error GoTo Failed

    MatchRejections.Reject proposal.Mail.MessageKey
    TicketAudit.RecordUserDecision proposal.Mail, proposal.Rule, proposal.TicketId, False, "rejected"
    TicketLog.Info "Review: proposal for ticket #" & proposal.TicketId & " rejected"
    Exit Sub

Failed:
    TicketLog.Fail "Review: rejecting a proposal failed: " & Err.Description
End Sub

' Adds the mail to the pending proposals if the ticket is open and readable. False otherwise
' (the mail then simply stays where it is). A mail that is already waiting counts as proposed.
Private Function QueueProposal(ByVal mail As IIncomingMail, ByVal index As ITicketIndex, ByVal ticketId As Long, _
                               ByVal rule As MatchRule) As Boolean
    Dim ticket As MatchTicket

    If MatchRejections.IsRejected(mail.MessageKey) Then
        TicketLog.Info "Ticket " & ticketId & ": the user rejected this mail before, no proposal"
        Exit Function
    End If

    Set ticket = index.GetTicket(ticketId)
    If ticket Is Nothing Then
        TicketLog.Info "Ticket " & ticketId & ": not readable, no proposal"
        Exit Function
    End If
    If Not ticket.IsOpen Then
        TicketLog.Info "Ticket " & ticketId & ": not open, no proposal"
        Exit Function
    End If

    MatchProposals.Add mail, ticketId, ticket.Summary, rule
    QueueProposal = True
End Function

' Moves the mail into the folder of an OPEN ticket (existing, restored or new). A closed or
' unreachable ticket leaves the mail where it is. Returns True if the mail was moved.
Private Function MoveToTicket(ByVal mail As IIncomingMail, ByVal store As ITicketStore, ByVal ticketId As Long) As Boolean
    Dim subject As String
    Dim state As NinjaTicketState
    Dim folder As ITicketFolder
    Dim ticketNumber As String
    Dim title As String

    state = NinjaApi.GetTicketInfo(ticketId, subject)
    If state <> NinjaTicketOpen Then
        TicketLog.Info "Ticket " & ticketId & ": not open or not readable, the mail is not moved"
        Exit Function
    End If

    ticketNumber = "#" & CStr(ticketId)
    title = subject
    If Len(Trim$(title)) = 0 Then title = TicketMatcher.NormalizeSubject(mail.Subject)

    Set folder = GetOrCreateTicketFolder(store, ticketNumber, " " & title)
    store.MoveMailToFolder mail, folder
    RenameToSubject folder, ticketId, subject
    MoveToTicket = True
End Function

' Active folder of the ticket; otherwise the archived folder is restored; otherwise a new one is
' created. remainingText is the folder title text (see TicketParsing.BuildFolderName).
Private Function GetOrCreateTicketFolder(ByVal store As ITicketStore, ByVal ticketNumber As String, _
                                         ByVal remainingText As String) As ITicketFolder
    Dim folder As ITicketFolder
    Dim archived As ITicketFolder

    Set folder = store.FindActiveFolder(ticketNumber)

    If folder Is Nothing Then
        Set archived = store.FindArchivedFolder(ticketNumber)
        If Not archived Is Nothing Then
            store.RestoreFromArchive archived
            Set folder = archived
            TicketLog.Info "Ticket " & ticketNumber & ": folder restored from the archive"
        End If
    End If

    If folder Is Nothing Then
        Set folder = store.CreateActiveFolder(TicketParsing.BuildFolderName(ticketNumber, remainingText))
        TicketLog.Info "Ticket " & ticketNumber & ": new folder '" & folder.Name & "'"
    End If

    Set GetOrCreateTicketFolder = folder
End Function

' ---------------------------------------------------------------------------
'  Archive
' ---------------------------------------------------------------------------

' Checks every active ticket folder and moves the closed ones into the archive.
' Returns the number of archived folders. A problem with one folder never stops the run.
' progress (optional) is told how far the run is; Finish is always called.
Public Function ArchiveClosedTickets(ByVal store As ITicketStore, ByVal fallbackEnabled As Boolean, _
        Optional ByVal progress As IProgressSink = Nothing) As Long
    Dim folders As Collection
    Dim i As Long
    Dim unavailableInARow As Long
    Dim apiDown As Boolean
    Dim archived As Long

    On Error GoTo Failed

    Set folders = store.ListActiveFolders()
    If Not progress Is Nothing Then progress.Begin folders.Count

    For i = 1 To folders.Count
        If Not progress Is Nothing Then progress.Advance "Ticket folder " & i & " of " & folders.Count & ": " & SafeName(folders(i))
        If ArchiveFolder(store, folders(i), fallbackEnabled, apiDown, unavailableInARow) Then archived = archived + 1
    Next i

    TicketLog.Info "Archive run finished: " & archived & " of " & folders.Count & " folders archived"
    If Not progress Is Nothing Then progress.Finish
    ArchiveClosedTickets = archived
    Exit Function

Failed:
    TicketLog.Fail "Archive run aborted: " & Err.Description
    If Not progress Is Nothing Then progress.Finish
    ArchiveClosedTickets = archived
End Function

' Returns True if the folder was archived.
Private Function ArchiveFolder(ByVal store As ITicketStore, ByVal folder As ITicketFolder, ByVal fallbackEnabled As Boolean, _
        ByRef apiDown As Boolean, ByRef unavailableInARow As Long) As Boolean
    Dim ticketId As Long
    Dim subject As String
    Dim state As NinjaTicketState
    Dim mailSaysClosed As Boolean
    Dim mailDate As Date
    Dim closedDate As Date
    Dim decision As ArchiveDecision
    Dim folderName As String

    On Error GoTo Failed

    folderName = folder.Name
    ticketId = TicketParsing.TicketIdFromFolderName(folderName)
    If ticketId <= 0 Then Exit Function   ' not a ticket folder

    If apiDown Then
        state = NinjaTicketUnavailable
    Else
        state = NinjaApi.GetTicketInfo(ticketId, subject)
        If state = NinjaTicketUnavailable Then
            unavailableInARow = unavailableInARow + 1
            TicketLog.Warn "Ticket " & ticketId & ": NinjaOne not available (" & NinjaApi.LastError & ")"
            If unavailableInARow >= API_DOWN_AFTER_FAILURES Then
                apiDown = True
                TicketLog.Warn "NinjaOne not available " & unavailableInARow & " times in a row: no more API calls in this run"
            End If
        Else
            unavailableInARow = 0
            RenameToSubject folder, ticketId, subject
        End If
    End If

    If TicketDecision.NeedsMailCheck(state, fallbackEnabled) Then
        mailSaysClosed = folder.FindClosedStatusMail(mailDate)
    End If

    decision = TicketDecision.DecideArchive(state, mailSaysClosed, fallbackEnabled)

    Select Case decision
        Case ArchiveByApi
            If state = NinjaTicketClosed Then
                closedDate = TimeUtil.UtcToLocal(NinjaApi.GetTicketClosedDate(ticketId))
            End If
            store.MoveToArchive folder, closedDate
            TicketLog.Info "Ticket " & ticketId & ": archived (" & IIf(state = NinjaTicketNotFound, "ticket no longer exists", "closed") & ")"
            ArchiveFolder = True

        Case ArchiveByMailFallback
            store.MoveToArchive folder, mailDate
            TicketLog.Warn "Ticket " & ticketId & ": archived by OFFLINE FALLBACK (mail says closed, date of that mail)"
            ArchiveFolder = True
    End Select
    Exit Function

Failed:
    TicketLog.Fail "Folder '" & folderName & "': " & Err.Description
End Function

' ---------------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------------

' Folder name for messages; never raises.
Private Function SafeName(ByVal folder As ITicketFolder) As String
    On Error Resume Next
    SafeName = folder.Name
    If Err.Number <> 0 Then SafeName = "(unknown folder)"
End Function

' Renames the folder to "#id (subject)" if the ticket subject changed.
' An empty subject (e.g. API not available) never renames.
Private Sub RenameToSubject(ByVal folder As ITicketFolder, ByVal ticketId As Long, ByVal subject As String)
    Dim expectedName As String

    If Len(Trim$(subject)) = 0 Then Exit Sub

    expectedName = "#" & CStr(ticketId) & " (" & subject & ")"
    If folder.Name <> expectedName Then
        On Error GoTo Failed
        folder.Name = expectedName
    End If
    Exit Sub

Failed:
    TicketLog.Warn "Ticket " & ticketId & ": folder could not be renamed (" & Err.Description & ")"
End Sub
