Attribute VB_Name = "ManualScanTests"
' ManualScanTests v0.1.0
'
' Tests of the detection of manual assignments (ManualScan) and of their place in the report.
' In-memory store, folders and index; no Outlook, no network, no real audit folder.
' Ticket 20101 and 30001 are open tickets, 30001 is found by its subject.

Option Explicit

Private pSink As FakeAuditSink
Private pStore As FakeTicketStore
Private pIndex As FakeTicketIndex
Private pFolder20101 As FakeTicketFolder
Private pFolder30001 As FakeTicketFolder
Private pSeen As Object

Private Const NOW_TIME As Double = 46300.5     ' a fixed day; mails are "received" relative to it

Private Sub Setup()
    TicketLog.DisableFileLogging
    TicketLog.Clear
    TicketAudit.ResetState
    MatchLevels.ResetState

    Set pSink = New FakeAuditSink
    TicketAudit.UseSink pSink

    Set pStore = New FakeTicketStore
    Set pFolder20101 = pStore.AddActive("#20101 (Example open ticket)")
    Set pFolder30001 = pStore.AddActive("#30001 (Example order total differs from expected value)")

    Set pIndex = New FakeTicketIndex
    pIndex.AddTicket 20101, "Example open ticket", True
    pIndex.AddTicket 30001, "Example order total differs from expected value", True

    Set pSeen = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Teardown()
    MatchLevels.ResetState
    TicketAudit.ResetState
    Set pSink = Nothing
    Set pStore = Nothing
    Set pIndex = Nothing
    Set pFolder20101 = Nothing
    Set pFolder30001 = Nothing
    Set pSeen = Nothing
End Sub

Private Function MailIn(ByVal folder As FakeTicketFolder, ByVal subject As String, ByVal key As String, _
                        Optional ByVal daysAgo As Long = 1) As FakeMail
    Dim mail As FakeMail

    Set mail = New FakeMail
    mail.SenderName = "Some Customer"
    mail.SenderDomain = "example.com"
    mail.Subject = subject
    mail.MessageKey = key
    mail.ConversationId = "conv-" & key
    mail.ReceivedTime = NOW_TIME - daysAgo
    folder.AddMail mail
    Set MailIn = mail
End Function

Private Function Scan() As Long
    Scan = ManualScan.ScanTicketFolders(pStore, pIndex, pSeen, NOW_TIME - 30)
End Function

' --- the four results ------------------------------------------------------------------

Public Sub Test_Scan_AgreeWhenTheMatcherWouldHaveChosenTheSameTicket()
    Setup
    MailIn pFolder20101, "Ticket 20101 bitte pruefen", "aa11aa11"

    TestAssert.AssertEqual 1, Scan()
    TestAssert.AssertEqual 1, pSink.Count
    TestAssert.AssertTrue InStr(pSink.LastLine(), """event"":""manual_assignment""") > 0, pSink.LastLine()
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""agree""") > 0
    TestAssert.AssertTrue InStr(pSink.LastLine(), """ticket"":20101") > 0
    TestAssert.AssertTrue InStr(pSink.LastLine(), """matcher_ticket"":20101") > 0
    TestAssert.AssertTrue InStr(pSink.LastLine(), """rule"":""NinjaId""") > 0
    TestAssert.AssertTrue InStr(pSink.LastLine(), """source"":""scan""") > 0
    TestAssert.AssertTrue InStr(pSink.LastLine(), """msg"":""aa11aa11""") > 0
    Teardown
End Sub

Public Sub Test_Scan_DifferWhenTheMatcherWouldHaveChosenAnotherTicket()
    Setup
    MailIn pFolder20101, "AW: Example order total differs from expected value", "bb22bb22"

    TestAssert.AssertEqual 1, Scan()
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""differ""") > 0, pSink.LastLine()
    TestAssert.AssertTrue InStr(pSink.LastLine(), """ticket"":20101") > 0, "the ticket of the folder"
    TestAssert.AssertTrue InStr(pSink.LastLine(), """matcher_ticket"":30001") > 0, "what the matcher would have chosen"
    Teardown
End Sub

Public Sub Test_Scan_MissedWhenTheMatcherFindsNothing()
    Setup
    MailIn pFolder20101, "Rueckfrage zum Drucker", "cc33cc33"

    TestAssert.AssertEqual 1, Scan()
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""missed""") > 0, pSink.LastLine()
    TestAssert.AssertTrue InStr(pSink.LastLine(), """outcome"":""none""") > 0
    TestAssert.AssertEqual 0, InStr(pSink.LastLine(), "matcher_ticket"), "no matcher ticket"
    TestAssert.AssertTrue InStr(pSink.LastLine(), """subject"":""rueckfrage zum drucker""") > 0, "the subject of a ticket mail is logged"
    Teardown
End Sub

Public Sub Test_Scan_ExcludedMailsAreMarked()
    Dim mail As FakeMail

    Setup
    Set mail = MailIn(pFolder20101, "Ticket 20101", "dd44dd44")
    mail.HasAutoHeader = True

    TestAssert.AssertEqual 1, Scan()
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""excluded""") > 0, pSink.LastLine()
    Teardown
End Sub

Public Sub Test_Scan_AllRulesOffMeansUnchecked()
    Setup
    MatchLevels.UseLevels 0, 0, 0, 0
    MailIn pFolder20101, "Ticket 20101", "ee55ee55"

    TestAssert.AssertEqual 1, Scan()
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""unchecked""") > 0, pSink.LastLine()
    Teardown
End Sub

' --- what is skipped ---------------------------------------------------------------------

Public Sub Test_Scan_EveryMailIsLoggedOnce()
    Setup
    MailIn pFolder20101, "Ticket 20101", "aa11aa11"

    TestAssert.AssertEqual 1, Scan()
    TestAssert.AssertEqual 0, Scan(), "the second scan finds nothing new"
    TestAssert.AssertEqual 1, pSink.Count
    TestAssert.AssertTrue pSeen.Exists("aa11aa11")
    Teardown
End Sub

Public Sub Test_Scan_KnownMailsAreSkipped()
    Setup
    MailIn pFolder20101, "Ticket 20101", "aa11aa11"
    pSeen.Add "aa11aa11", True

    TestAssert.AssertEqual 0, Scan()
    TestAssert.AssertEqual 0, pSink.Count
    Teardown
End Sub

Public Sub Test_Scan_TicketSenderMailsAreNotManual()
    Dim mail As FakeMail

    Setup
    Set mail = MailIn(pFolder20101, "[example] (#20101) Example", "ff66ff66")
    mail.SenderName = TicketConfig.AllowedSender()

    TestAssert.AssertEqual 0, Scan()
    TestAssert.AssertEqual 0, pSink.Count
    Teardown
End Sub

Public Sub Test_Scan_MailsWithoutMessageKeyAreSkipped()
    Setup
    MailIn pFolder20101, "Ticket 20101", ""

    TestAssert.AssertEqual 0, Scan()
    Teardown
End Sub

Public Sub Test_Scan_OnlyRecentMails()
    Setup
    MailIn pFolder20101, "Ticket 20101", "aa11aa11", 5
    MailIn pFolder20101, "Ticket 20101 alt", "bb22bb22", 45

    TestAssert.AssertEqual 1, Scan(), "the mail received 45 days ago is out of range"
    Teardown
End Sub

Public Sub Test_Scan_FoldersThatAreNoTicketFoldersAreSkipped()
    Dim other As FakeTicketFolder

    Setup
    Set other = pStore.AddActive("Sonstiges")
    MailIn other, "Ticket 20101", "aa11aa11"

    TestAssert.AssertEqual 0, Scan()
    TestAssert.AssertEqual 0, other.RecentMailsRequests, "such a folder is not even read"
    Teardown
End Sub

Public Sub Test_Scan_ABrokenFolderDoesNotStopTheScan()
    Dim broken As FakeTicketFolder

    Setup
    Set broken = pStore.AddActive("#999 (broken)")
    broken.FailOnNameRead = True
    MailIn pFolder30001, "Hallo", "aa11aa11"

    TestAssert.AssertEqual 1, Scan()
    TestAssert.AssertTrue InStr(TicketLog.Text(), "Scan of folder") > 0, TicketLog.Text()
    Teardown
End Sub

' --- side effects and bookkeeping ----------------------------------------------------------

Public Sub Test_Scan_RemembersTheConversationForTheMatcher()
    Dim idx As ITicketIndex

    Setup
    MailIn pFolder30001, "Hallo", "aa11aa11"

    Scan
    Set idx = pIndex
    TestAssert.AssertEqual 30001, idx.FindByConversation("conv-aa11aa11"), "a manual assignment is a safe label for its thread"
    Teardown
End Sub

Public Sub Test_Scan_SeenKeysFromRecords()
    Dim records As Collection
    Dim seen As Object

    Set records = AuditReport.ParseLines( _
        "{""event"":""auto_decision"",""action"":""moved"",""msg"":""a1a1a1a1""}" & vbLf & _
        "{""event"":""auto_decision"",""action"":""proposed"",""msg"":""b2b2b2b2""}" & vbLf & _
        "{""event"":""auto_decision"",""action"":""log_only"",""msg"":""c3c3c3c3""}" & vbLf & _
        "{""event"":""user_confirmed"",""action"":""moved"",""msg"":""d4d4d4d4""}" & vbLf & _
        "{""event"":""user_rejected"",""action"":""rejected"",""msg"":""e5e5e5e5""}" & vbLf & _
        "{""event"":""manual_assignment"",""action"":""agree"",""msg"":""f6f6f6f6""}" & vbLf & _
        "{""event"":""manual_assignment"",""action"":""agree"",""msg"":""""}")
    Set seen = ManualScan.SeenKeysFromRecords(records)

    TestAssert.AssertTrue seen.Exists("a1a1a1a1"), "moved by the automation"
    TestAssert.AssertFalse seen.Exists("b2b2b2b2"), "only proposed"
    TestAssert.AssertFalse seen.Exists("c3c3c3c3"), "only logged"
    TestAssert.AssertTrue seen.Exists("d4d4d4d4"), "confirmed in the review"
    TestAssert.AssertFalse seen.Exists("e5e5e5e5"), "rejected, a later manual move still counts"
    TestAssert.AssertTrue seen.Exists("f6f6f6f6"), "seen by an earlier scan"
    TestAssert.AssertEqual 3, seen.Count
End Sub

' --- the report ----------------------------------------------------------------------------

Public Sub Test_Report_ShowsManualAssignments()
    Dim lines As String
    Dim i As Long
    Dim text As String

    For i = 1 To 7
        lines = lines & AuditLogic.BuildMatchEvent(DateSerial(2026, 9, i), "manual_assignment", "scan", IIf(i <= 4, "found", "none"), IIf(i <= 4, "NinjaId", "None"), _
            20108 + i, IIf(i <= 3, "agree", IIf(i = 4, "differ", "missed")), IIf(i <= 5, "example.com", "vendor.example"), "item235 - error " & i, "x", "", "", "v" & i, "m" & i, False, IIf(i = 4, 30001, 0)) & vbLf
    Next i
    text = AuditReport.BuildReport(AuditReport.ParseLines(lines), DateSerial(2026, 10, 9))

    TestAssert.AssertTrue InStr(text, "## Manuelle Zuordnungen") > 0, text
    TestAssert.AssertTrue InStr(text, "7 Mails wurden von Hand in einen Ticketordner verschoben") > 0, text
    TestAssert.AssertTrue InStr(text, "3 (42 %)") > 0, "agreement rate"
    TestAssert.AssertTrue InStr(text, "| anderes Ticket (differ) | 1 |") > 0
    TestAssert.AssertTrue InStr(text, "| verpasst (missed) | 3 |") > 0
    TestAssert.AssertTrue InStr(text, "- vendor.example: 2") > 0, "missed assignments by domain"
    TestAssert.AssertTrue InStr(text, "manuell 7)") > 0, "overview counts them"
    TestAssert.AssertTrue InStr(text, "| example.com | item# - error # | 5 |") > 0, "manual assignments count as confirmations in the pattern table"
End Sub

Public Sub Test_Report_NoManualAssignmentsYet()
    Dim text As String

    text = AuditReport.BuildReport(New Collection, DateSerial(2026, 10, 9))
    TestAssert.AssertTrue InStr(text, "## Manuelle Zuordnungen") > 0
    TestAssert.AssertTrue InStr(text, "Noch keine von Hand") > 0, text
End Sub
