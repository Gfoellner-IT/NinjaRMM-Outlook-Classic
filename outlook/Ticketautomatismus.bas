Attribute VB_Name = "Ticketautomatismus"
' Ticketautomatismus v2.5.0
' @author NeumaC
' https://github.com/NeumaC/NinjaRMM_VBA
'
' Outlook entry points of the ticket automation (rules, ribbon buttons):
'   RunEmailRule        sorts all ticket mails of the inbox into ticket folders
'   RunEmailRuleScript  for the rule action "Run a script": sorts the one mail the rule hands over
'   RunArchiveRule      moves folders of closed tickets into the archive
'   RunManualAssignmentScan  logs mails the user moved into ticket folders by hand (M7), silent
'   RunMatchRulesReview (button) writes the report on the audit log and copies the start sentence for a review with Claude
'   RunMatchReview      (button) collects the proposals of the last days and shows them for review (level 2)
'   RunMatchDryRun      logs which existing ticket each of the newest foreign mails would be assigned
'                       to (M6); never moves a mail. Mails are only moved by the sweep and the watcher
'                       if the safety level of the matching rule says so (TicketConfig.MATCH_LEVEL_*).
'
'   RunArchiveRuleWithProgress  like RunArchiveRule, with a non-modal progress window (for buttons)
'   StartMailWatcher / StopMailWatcher  sort new mails as soon as they arrive (event ItemAdd);
'                       call StartMailWatcher from Application_Startup in ThisOutlookSession
'
' This module only connects Outlook to the logic in TicketService. It shows no dialogs
' (the rules run unattended); problems are written to TicketLog.
' Configuration: TicketConfig. Required VBA reference: the Outlook object library only.

Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private pWatcher As TicketMailWatcher
Private pIndex As NinjaTicketIndex     ' lives as long as Outlook runs, so its cache survives
Private pPrompt As IMatchPrompt

' ---------------------------------------------------------------------------
'  Entry points
' ---------------------------------------------------------------------------

Public Sub RunEmailRule()
    SweepInbox True
End Sub

' Sorts ticket mails and matches the foreign mails of the last MATCH_SWEEP_DAYS days. offerReview = True
' asks once at the end whether to review new proposals; the start-up sweep passes False (no dialog while
' Outlook is starting). It has a parameter and is therefore not a macro itself; RunEmailRule is the one.
Private Sub SweepInbox(ByVal offerReview As Boolean)
    Dim inbox As Outlook.Folder
    Dim ticketsFolder As Outlook.Folder
    Dim store As ITicketStore
    Dim items As Outlook.Items
    Dim item As Object
    Dim incoming As IIncomingMail
    Dim i As Long
    Dim moved As Long
    Dim matched As Long

    LearnedRules.Reload                           ' pick up edits of the rules file
    If Not InitFolders(inbox, ticketsFolder) Then Exit Sub
    Set store = NewStore(ticketsFolder)

    Set items = inbox.Items
    For i = items.Count To 1 Step -1
        Set item = items(i)
        If TypeOf item Is Outlook.MailItem Then
            Set incoming = NewMail(item)
            If TicketService.ProcessMail(incoming, store) Then
                moved = moved + 1
            ElseIf incoming.ReceivedTime >= Now - TicketConfig.MATCH_SWEEP_DAYS Then
                If TicketService.MatchForeignMail(incoming, MatchIndex(), store, True, "sweep") = MatchFound Then matched = matched + 1
            End If
        End If
    Next i

    TicketLog.Info "RunEmailRule finished: " & moved & " ticket mails moved, " & matched & " foreign mails matched"

    If offerReview Then MatchReview.OfferReview ReviewPrompt(), store
End Sub

' For a menu button: collects the proposals of the last days (a sweep without the hint question), then
' shows them for review. Only says so if no rule is on level 2 (confirm) or there is nothing to review.
Public Sub RunMatchReview()
    Dim inbox As Outlook.Folder
    Dim ticketsFolder As Outlook.Folder
    Dim store As ITicketStore

    If Not InitFolders(inbox, ticketsFolder) Then Exit Sub
    Set store = NewStore(ticketsFolder)

    If Not MatchLevels.AnyRuleAtLevel(LevelConfirm) And MatchProposals.Count() = 0 Then
        MsgBox MatchPromptText.NoConfirmLevelText(), vbInformation, MatchPromptText.DialogTitle()
        Exit Sub
    End If

    SweepInbox False

    If MatchProposals.Count() = 0 Then
        MsgBox MatchPromptText.NoProposalsText(), vbInformation, MatchPromptText.DialogTitle()
        Exit Sub
    End If

    MatchReview.ReviewPending ReviewPrompt(), store
End Sub

' Looks for mails the user moved into a ticket folder by hand and logs them in the audit log (ManualScan).
' No dialog; also run silently at start and before the rules report. Does nothing if the audit log is off.
Public Sub RunManualAssignmentScan()
    Dim inbox As Outlook.Folder
    Dim ticketsFolder As Outlook.Folder
    Dim sink As FileAuditSink
    Dim seen As Object
    Dim found As Long

    On Error GoTo Failed
    If Not TicketConfig.ENABLE_AUDIT_LOG Then Exit Sub
    If Not InitFolders(inbox, ticketsFolder) Then Exit Sub

    Set sink = New FileAuditSink
    Set seen = ManualScan.SeenKeysFromRecords(AuditReport.LoadRecords(sink.Folder(), TicketConfig.REPORT_MONTHS_BACK, Date))
    found = ManualScan.ScanTicketFolders(NewStore(ticketsFolder), MatchIndex(), seen, Now - TicketConfig.MANUAL_SCAN_DAYS)
    TicketLog.Info "Scan for manual assignments finished: " & found & " new"
    Exit Sub

Failed:
    TicketLog.Fail "RunManualAssignmentScan failed: " & Err.Description
End Sub

' For a menu button: evaluates the audit log (AuditReport), writes the report as a Markdown file and puts the
' sentence that starts the review of the rules with Claude on the clipboard. Read-only, changes no rule.
Public Sub RunMatchRulesReview()
    Dim sink As FileAuditSink
    Dim records As Collection
    Dim reportText As String
    Dim path As String
    Dim copied As Boolean
    Dim message As String

    On Error GoTo Failed

    RunManualAssignmentScan                       ' bring the manual assignments up to date first
    Set sink = New FileAuditSink
    Set records = AuditReport.LoadRecords(sink.Folder(), TicketConfig.REPORT_MONTHS_BACK, Date)
    reportText = AuditReport.BuildReport(records, Now)
    path = AuditReport.WriteReport(reportText, AuditReport.ReportFolder(), Date)
    copied = ClipboardHelper.CopyText(AuditReport.ReviewPromptText(path))

    message = GermanText.De("Der Bericht wurde geschrieben:") & vbCrLf & path & vbCrLf & vbCrLf
    If copied Then
        message = message & GermanText.De("Der Einleitungssatz liegt in der Zwischenablage. F[ue]ge ihn in eine Claude-Sitzung ein (Strg+V) oder tippe dort /review-rules.")
    Else
        message = message & GermanText.De("Der Einleitungssatz konnte nicht kopiert werden. Tippe in der Claude-Sitzung /review-rules.")
    End If
    MsgBox message, vbInformation, MatchPromptText.DialogTitle()
    Exit Sub

Failed:
    TicketLog.Fail "RunMatchRulesReview failed: " & Err.Description
    MsgBox GermanText.De("Der Bericht konnte nicht erstellt werden: ") & Err.Description, vbExclamation, MatchPromptText.DialogTitle()
End Sub

' Dry run of the matching for the newest inbox mails (default 100): logs one decision per mail and
' a summary, never moves anything (independent of the safety levels).
' Read the result in %TEMP%\ninja_ticket_log.txt.
Public Sub RunMatchDryRun(Optional ByVal count As Long = 100)
    Dim inbox As Outlook.Folder
    Dim ticketsFolder As Outlook.Folder
    Dim items As Outlook.Items
    Dim item As Object
    Dim incoming As IIncomingMail
    Dim store As ITicketStore
    Dim outcome As MatchOutcome
    Dim checked As Long
    Dim found As Long
    Dim ambiguous As Long

    LearnedRules.Reload
    If Not InitFolders(inbox, ticketsFolder) Then Exit Sub
    Set store = NewStore(ticketsFolder)

    Set items = inbox.Items
    items.Sort "[ReceivedTime]", True

    For Each item In items
        If checked >= count Then Exit For
        If TypeOf item Is Outlook.MailItem Then
            Set incoming = NewMail(item)
            If incoming.SenderName <> TicketConfig.AllowedSender() Then
                checked = checked + 1
                outcome = TicketService.MatchForeignMail(incoming, MatchIndex(), store, False, "dry_run")
                If outcome = MatchFound Then found = found + 1
                If outcome = MatchAmbiguous Then ambiguous = ambiguous + 1
            End If
        End If
    Next item

    TicketLog.Info "Match dry run finished: " & checked & " foreign mails, " & found & " matched, " & ambiguous & _
                   " ambiguous, " & (checked - found - ambiguous) & " without match"
End Sub

' Rule action "Run a script": processes only the mail handed over by the rule, and only if it
' is still in the inbox after all other rules ran (a rule may have moved it elsewhere).
Public Sub RunEmailRuleScript(ByVal mail As Outlook.MailItem)
    Dim inbox As Outlook.Folder
    Dim ticketsFolder As Outlook.Folder

    Sleep 1000   ' give the other rules time to finish

    On Error GoTo ExitSub
    If mail Is Nothing Then Exit Sub
    If Not InitFolders(inbox, ticketsFolder) Then Exit Sub

    ' Fails if the mail was moved or deleted in the meantime -> leave quietly.
    If mail.Parent.EntryID <> inbox.EntryID Then Exit Sub

    TicketService.ProcessMail NewMail(mail), NewStore(ticketsFolder)

ExitSub:
End Sub

Public Sub RunArchiveRule()
    Dim inbox As Outlook.Folder
    Dim ticketsFolder As Outlook.Folder

    If Not InitFolders(inbox, ticketsFolder) Then Exit Sub

    TicketService.ArchiveClosedTickets NewStore(ticketsFolder), TicketConfig.ENABLE_OFFLINE_FALLBACK
End Sub

' For buttons: the same as RunArchiveRule, but shows a non-modal progress window.
Public Sub RunArchiveRuleWithProgress()
    Dim inbox As Outlook.Folder
    Dim ticketsFolder As Outlook.Folder
    Dim progress As IProgressSink

    If Not InitFolders(inbox, ticketsFolder) Then Exit Sub

    Set progress = New ProgressSinkForm
    TicketService.ArchiveClosedTickets NewStore(ticketsFolder), TicketConfig.ENABLE_OFFLINE_FALLBACK, progress
End Sub

' Starts the event-driven sorting of new mails (call it from Application_Startup):
'   Private Sub Application_Startup()
'       Ticketautomatismus.StartMailWatcher
'   End Sub
' ItemAdd does not run when many mails arrive at once, so one sweep over the inbox follows.
Public Sub StartMailWatcher()
    Dim inbox As Outlook.Folder
    Dim ticketsFolder As Outlook.Folder

    If Not InitFolders(inbox, ticketsFolder) Then Exit Sub

    StopMailWatcher
    Set pWatcher = New TicketMailWatcher
    pWatcher.StartWatching inbox, NewStore(ticketsFolder), TicketConfig.EVENT_DELAY_MS, MatchIndex(), ReviewPrompt()
    TicketLog.Info "Mail watcher started"

    SweepInbox False   ' no dialog while Outlook starts; proposals are offered with the next new mail
    RunManualAssignmentScan
End Sub

Public Sub StopMailWatcher()
    If Not pWatcher Is Nothing Then
        pWatcher.StopWatching
        Set pWatcher = Nothing
        TicketLog.Info "Mail watcher stopped"
    End If
End Sub

' Kept for callers outside this project (e.g. ThisOutlookSession).
Public Sub ProcessEmail(ByVal mailItem As Outlook.MailItem, ByVal tasksFolder As Outlook.Folder)
    TicketService.ProcessMail NewMail(mailItem), NewStore(tasksFolder)
End Sub

' ---------------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------------

Private Function InitFolders(ByRef inbox As Outlook.Folder, ByRef ticketsFolder As Outlook.Folder) As Boolean
    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)

    On Error Resume Next
    Set ticketsFolder = inbox.Folders(TicketConfig.FOLDER_TICKETS)
    On Error GoTo 0

    If ticketsFolder Is Nothing Then
        TicketLog.Fail "The folder '" & TicketConfig.FOLDER_TICKETS & "' does not exist below the inbox."
        Exit Function
    End If
    InitFolders = True
End Function

' The dialogs of the confirmation level: the review form with a list (variant B, falls back to message
' boxes if the form is not imported) or plain message boxes (variant A), see TicketConfig.REVIEW_DIALOG.
Private Function ReviewPrompt() As IMatchPrompt
    If pPrompt Is Nothing Then
        If TicketConfig.REVIEW_DIALOG = 2 Then
            Set pPrompt = New FormMatchPrompt
        Else
            Set pPrompt = New MsgBoxMatchPrompt
        End If
    End If
    Set ReviewPrompt = pPrompt
End Function

Private Function MatchIndex() As ITicketIndex
    If pIndex Is Nothing Then Set pIndex = New NinjaTicketIndex
    Set MatchIndex = pIndex
End Function

Private Function NewStore(ByVal ticketsFolder As Outlook.Folder) As ITicketStore
    Dim store As OutlookTicketStore

    Set store = New OutlookTicketStore
    store.Init ticketsFolder
    Set NewStore = store
End Function

Private Function NewMail(ByVal mail As Outlook.MailItem) As IIncomingMail
    Dim adapter As OutlookIncomingMail

    Set adapter = New OutlookIncomingMail
    adapter.Init mail
    Set NewMail = adapter
End Function
