Attribute VB_Name = "ManualScan"
' ManualScan v0.1.0
'
' Finds mails the user moved into a ticket folder by hand (M7) and writes one audit line per mail
' (event manual_assignment). The line says what the matcher would have done with the mail:
'
'   agree      it would have chosen the same ticket
'   differ     it would have chosen another ticket
'   missed     it found no ticket, or no unambiguous one
'   excluded   the mail is of a kind that is never matched (automatic mail, excluded domain)
'   unchecked  all matching rules are switched off
'
' Mails of the ticket sender (the ticket system's own mails) and mails already known to the audit log
' (moved by the automation, confirmed in the review, or seen by an earlier scan) are skipped, so every
' manual assignment is logged once. A manual assignment is also a safe label: the conversation is
' remembered for the matcher's conversation rule. No Outlook types.

Option Explicit

' Scans the active ticket folders for mails received since the given time. seenKeys (Dictionary,
' message key -> True) is updated. Returns the number of manual assignments recorded.
Public Function ScanTicketFolders(ByVal store As ITicketStore, ByVal index As ITicketIndex, ByVal seenKeys As Object, _
                                  ByVal since As Date) As Long
    Dim folders As Collection
    Dim i As Long

    On Error GoTo Failed
    Set folders = store.ListActiveFolders()
    For i = 1 To folders.Count
        ScanTicketFolders = ScanTicketFolders + ScanFolder(folders(i), index, seenKeys, since)
    Next i
    Exit Function

Failed:
    TicketLog.Fail "Scan for manual assignments aborted: " & Err.Description
End Function

' The message keys the audit log already knows (Dictionary, key -> True): mails the automation
' moved, mails confirmed in the review and manual assignments seen before.
Public Function SeenKeysFromRecords(ByVal records As Collection) As Object
    Dim seen As Object
    Dim rec As Object
    Dim ev As String
    Dim key As String

    Set seen = CreateObject("Scripting.Dictionary")
    For Each rec In records
        ev = JsonLite.GetString(rec, "event")
        key = JsonLite.GetString(rec, "msg")
        If Len(key) > 0 Then
            If ev = "user_confirmed" Or ev = "manual_assignment" Or _
               (ev = "auto_decision" And JsonLite.GetString(rec, "action") = "moved") Then
                If Not seen.Exists(key) Then seen.Add key, True
            End If
        End If
    Next rec
    Set SeenKeysFromRecords = seen
End Function

' ---------------------------------------------------------------------------
'  Implementation
' ---------------------------------------------------------------------------

' Returns the number of manual assignments found in this folder. A problem with the folder or one
' mail never stops the scan.
Private Function ScanFolder(ByVal folder As ITicketFolder, ByVal index As ITicketIndex, ByVal seenKeys As Object, _
                            ByVal since As Date) As Long
    Dim folderName As String
    Dim ticketId As Long
    Dim mails As Collection
    Dim mail As IIncomingMail

    On Error GoTo Failed
    folderName = folder.Name
    ticketId = TicketParsing.TicketIdFromFolderName(folderName)
    If ticketId <= 0 Then Exit Function

    Set mails = folder.RecentMails(since)
    For Each mail In mails
        If ScanMail(mail, ticketId, index, seenKeys) Then ScanFolder = ScanFolder + 1
    Next mail
    Exit Function

Failed:
    TicketLog.Fail "Scan of folder '" & folderName & "': " & Err.Description
End Function

Private Function ScanMail(ByVal mail As IIncomingMail, ByVal ticketId As Long, ByVal index As ITicketIndex, _
                          ByVal seenKeys As Object) As Boolean
    Dim key As String
    Dim outcome As MatchOutcome
    Dim rule As MatchRule
    Dim matcherTicket As Long
    Dim action As String

    On Error GoTo Failed
    If mail.SenderName = TicketConfig.AllowedSender() Then Exit Function

    key = mail.MessageKey
    If Len(key) = 0 Then Exit Function              ' cannot be found again later
    If seenKeys.Exists(key) Then Exit Function

    If TicketMatcher.IsExcluded(mail.SenderDomain, mail.HasAutoHeader) Then
        action = "excluded"
    ElseIf Not MatchLevels.AnyRuleActive() Then
        action = "unchecked"
    Else
        outcome = TicketMatcher.MatchMail(mail.Subject, mail.BodyText, mail.ConversationId, mail.SenderDomain, _
                                          mail.HasAutoHeader, index, matcherTicket, rule)
        If outcome <> MatchFound Then matcherTicket = 0
        If outcome = MatchFound And matcherTicket = ticketId Then
            action = "agree"
        ElseIf outcome = MatchFound Then
            action = "differ"
        Else
            action = "missed"
        End If
    End If

    TicketAudit.RecordManualAssignment mail, outcome, rule, ticketId, matcherTicket, action
    seenKeys.Add key, True
    If Not index Is Nothing Then index.RememberConversation mail.ConversationId, ticketId

    TicketLog.Info "Manual assignment to ticket #" & ticketId & ": " & action
    ScanMail = True
    Exit Function

Failed:
    TicketLog.Fail "Scan of a mail in ticket #" & ticketId & ": " & Err.Description
End Function
