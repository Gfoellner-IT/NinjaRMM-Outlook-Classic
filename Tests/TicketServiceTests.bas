Attribute VB_Name = "TicketServiceTests"
' TicketServiceTests v0.1.0
'
' Tests of the ticket automation with an in-memory store (FakeTicketStore) and a fake
' NinjaOne transport. No Outlook, no network.

Option Explicit

Private pStore As FakeTicketStore
Private pTransport As FakeNinjaTransport

' ---------------------------------------------------------------------------
'  Setup helpers
' ---------------------------------------------------------------------------

Private Sub Setup()
    TicketLog.DisableFileLogging
    TicketLog.Clear
    NinjaApi.ResetState

    Set pStore = New FakeTicketStore
    Set pTransport = New FakeNinjaTransport
    NinjaApi.UseTransport pTransport
End Sub

Private Sub Teardown()
    NinjaApi.ResetState
    Set pStore = Nothing
    Set pTransport = Nothing
End Sub

Private Function NewMail(ByVal subject As String, Optional ByVal sender As String = "") As FakeMail
    Dim mail As FakeMail

    Set mail = New FakeMail
    mail.Subject = subject
    If Len(sender) = 0 Then mail.SenderName = TicketConfig.AllowedSender() Else mail.SenderName = sender
    Set NewMail = mail
End Function

Private Function Process(ByVal mail As FakeMail) As Boolean
    Process = TicketService.ProcessMail(mail, pStore)
End Function

' Ticket 20101 is open, 20102 is closed (log: closed manually 2026-10-06 14:01:28 UTC).
Private Sub AnswerOpenAndClosedTickets()
    pTransport.AddResponse "ticketing/ticket/20101", 200, Fixtures.TicketOpen()
    pTransport.AddResponse "ticketing/ticket/20102", 200, Fixtures.TicketClosed()
    pTransport.AddResponse "ticketing/ticket/20102/log-entry?type=SAVE", 200, Fixtures.LogClosedManual()
End Sub

' ---------------------------------------------------------------------------
'  New mails
' ---------------------------------------------------------------------------

Public Sub Test_Mail_NewTicketCreatesFolderAndMovesMail()
    Dim mail As FakeMail

    Setup
    Set mail = NewMail("[example] (#3621) Printer broken")

    TestAssert.AssertTrue Process(mail)
    TestAssert.AssertEqual 1, pStore.CreatedCount
    TestAssert.AssertEqual 1, pStore.ActiveCount
    TestAssert.AssertEqual "#3621 (Printer broken)", pStore.ActiveAt(1).FolderName, "API not reachable: name from the subject"
    TestAssert.AssertTrue mail.MovedTo Is pStore.ActiveAt(1), "mail is in the new folder"
    Teardown
End Sub

Public Sub Test_Mail_FolderIsRenamedToCurrentTicketSubject()
    Dim mail As FakeMail

    Setup
    pTransport.AddResponse "ticketing/ticket/20101", 200, Fixtures.TicketOpen()
    Set mail = NewMail("[example] (#20101) Old mail subject")

    TestAssert.AssertTrue Process(mail)
    TestAssert.AssertEqual "#20101 (Example open ticket)", pStore.ActiveAt(1).FolderName
    Teardown
End Sub

Public Sub Test_Mail_ExistingFolderIsReused()
    Dim existing As FakeTicketFolder
    Dim mail As FakeMail

    Setup
    Set existing = pStore.AddActive("#3621 (Printer broken)")
    Set mail = NewMail("[example] (#3621) Re: Printer broken")

    TestAssert.AssertTrue Process(mail)
    TestAssert.AssertEqual 0, pStore.CreatedCount
    TestAssert.AssertTrue mail.MovedTo Is existing
    TestAssert.AssertEqual 1, existing.MailCount
    Teardown
End Sub

Public Sub Test_Mail_ShorterTicketNumberDoesNotMatchLongerFolder()
    Dim other As FakeTicketFolder
    Dim mail As FakeMail

    Setup
    Set other = pStore.AddActive("#3621 (Other ticket)")
    Set mail = NewMail("[example] (#36) Short number")

    TestAssert.AssertTrue Process(mail)
    TestAssert.AssertEqual 1, pStore.CreatedCount, "new folder for #36"
    TestAssert.AssertEqual 0, other.MailCount, "the folder of #3621 stays untouched"
    Teardown
End Sub

Public Sub Test_Mail_ArchivedFolderIsRestored()
    Dim archived As FakeTicketFolder
    Dim mail As FakeMail

    Setup
    Set archived = pStore.AddArchived("#3621 (Printer broken)")
    Set mail = NewMail("[example] (#3621) Reopened")

    TestAssert.AssertTrue Process(mail)
    TestAssert.AssertEqual 1, pStore.RestoredCount
    TestAssert.AssertEqual 0, pStore.CreatedCount
    TestAssert.AssertEqual 0, pStore.ArchivedCount
    TestAssert.AssertEqual 1, pStore.ActiveCount
    TestAssert.AssertTrue mail.MovedTo Is archived
    TestAssert.AssertFalse archived.IsArchived
    Teardown
End Sub

Public Sub Test_Mail_OtherSenderIsIgnored()
    Dim mail As FakeMail

    Setup
    Set mail = NewMail("[example] (#3621) x", "someone else")

    TestAssert.AssertFalse Process(mail)
    TestAssert.AssertEqual 0, pStore.ActiveCount
    TestAssert.AssertTrue mail.MovedTo Is Nothing
    Teardown
End Sub

Public Sub Test_Mail_NonTicketSubjectIsIgnored()
    Dim mail As FakeMail

    Setup
    Set mail = NewMail("Newsletter October")

    TestAssert.AssertFalse Process(mail)
    TestAssert.AssertEqual 0, pStore.ActiveCount
    Teardown
End Sub

Public Sub Test_Mail_AbsurdTicketNumberIsIgnored()
    Setup

    TestAssert.AssertFalse Process(NewMail("[example] (#12345678901234) x"))
    TestAssert.AssertEqual 0, pStore.ActiveCount
    Teardown
End Sub

Public Sub Test_Mail_FolderNameStaysIfSubjectIsAlreadyCurrent()
    Dim existing As FakeTicketFolder

    Setup
    pTransport.AddResponse "ticketing/ticket/20101", 200, Fixtures.TicketOpen()
    Set existing = pStore.AddActive("#20101 (Example open ticket)")

    TestAssert.AssertTrue Process(NewMail("[example] (#20101) x"))
    TestAssert.AssertEqual 0, existing.RenameCount, "no needless rename"
    Teardown
End Sub

' ---------------------------------------------------------------------------
'  Archive
' ---------------------------------------------------------------------------

Public Sub Test_Archive_ClosedTicketIsArchivedWithClosingDate()
    Dim folder As FakeTicketFolder

    Setup
    AnswerOpenAndClosedTickets
    Set folder = pStore.AddActive("#20102 (Example)")

    TestAssert.AssertEqual 1, TicketService.ArchiveClosedTickets(pStore, True)
    TestAssert.AssertEqual 0, pStore.ActiveCount
    TestAssert.AssertEqual 1, pStore.ArchivedCount
    TestAssert.AssertDateNear TimeUtil.UtcToLocal(DateSerial(2026, 10, 6) + TimeSerial(14, 1, 28)), folder.ArchivedFor, 0, "closing date from the ticket log"
    TestAssert.AssertEqual 0, folder.MailChecks, "mails are not read when the API answered"
    Teardown
End Sub

Public Sub Test_Archive_OpenTicketStays()
    Dim folder As FakeTicketFolder

    Setup
    AnswerOpenAndClosedTickets
    Set folder = pStore.AddActive("#20101 (Example open ticket)")

    TestAssert.AssertEqual 0, TicketService.ArchiveClosedTickets(pStore, True)
    TestAssert.AssertEqual 1, pStore.ActiveCount
    TestAssert.AssertFalse folder.IsArchived
    Teardown
End Sub

Public Sub Test_Archive_OpenTicketStaysEvenIfOldMailSaysClosed()
    Dim folder As FakeTicketFolder

    Setup
    AnswerOpenAndClosedTickets
    Set folder = pStore.AddActive("#20101 (Example open ticket)")
    folder.HasClosedStatusMail = True
    folder.ClosedStatusMailDate = DateSerial(2026, 1, 5)

    TestAssert.AssertEqual 0, TicketService.ArchiveClosedTickets(pStore, True)
    TestAssert.AssertEqual 0, folder.MailChecks, "the API decides"
    Teardown
End Sub

Public Sub Test_Archive_DeletedTicketIsArchivedWithTodaysDate()
    Dim folder As FakeTicketFolder

    Setup
    pTransport.AddResponse "ticketing/ticket/99999999", 404, ""
    Set folder = pStore.AddActive("#99999999 (Gone)")

    TestAssert.AssertEqual 1, TicketService.ArchiveClosedTickets(pStore, True)
    TestAssert.AssertEqual 0#, CDbl(folder.ArchivedFor), "date 0 = the store uses today"
    TestAssert.AssertEqual 1, pTransport.RequestCount, "no log request for a ticket that does not exist"
    Teardown
End Sub

Public Sub Test_Archive_ApiUnavailable_FallbackArchivesWithMailDate()
    Dim folder As FakeTicketFolder

    Setup
    Set folder = pStore.AddActive("#20102 (Example)")
    folder.HasClosedStatusMail = True
    folder.ClosedStatusMailDate = DateSerial(2026, 9, 15) + TimeSerial(10, 0, 0)

    TestAssert.AssertEqual 1, TicketService.ArchiveClosedTickets(pStore, True)
    TestAssert.AssertDateNear DateSerial(2026, 9, 15) + TimeSerial(10, 0, 0), folder.ArchivedFor, 0, "date of the mail, not today"
    TestAssert.AssertTrue InStr(TicketLog.Text(), "OFFLINE FALLBACK") > 0, "the fallback is visible in the log"
    Teardown
End Sub

Public Sub Test_Archive_ApiUnavailable_FallbackDisabledDoesNothing()
    Dim folder As FakeTicketFolder

    Setup
    Set folder = pStore.AddActive("#20102 (Example)")
    folder.HasClosedStatusMail = True
    folder.ClosedStatusMailDate = DateSerial(2026, 9, 15)

    TestAssert.AssertEqual 0, TicketService.ArchiveClosedTickets(pStore, False)
    TestAssert.AssertEqual 1, pStore.ActiveCount
    TestAssert.AssertEqual 0, folder.MailChecks, "mails are not even read"
    Teardown
End Sub

Public Sub Test_Archive_ApiUnavailable_MailNotClosedStays()
    Dim folder As FakeTicketFolder

    Setup
    Set folder = pStore.AddActive("#20102 (Example)")

    TestAssert.AssertEqual 0, TicketService.ArchiveClosedTickets(pStore, True)
    TestAssert.AssertEqual 1, folder.MailChecks, "mails were checked"
    TestAssert.AssertFalse folder.IsArchived
    Teardown
End Sub

Public Sub Test_Archive_ServerErrorCountsAsUnavailable()
    Dim folder As FakeTicketFolder

    Setup
    pTransport.AddResponse "ticketing/ticket/20102", 503, ""
    Set folder = pStore.AddActive("#20102 (Example)")
    folder.HasClosedStatusMail = True
    folder.ClosedStatusMailDate = DateSerial(2026, 9, 15)

    TestAssert.AssertEqual 1, TicketService.ArchiveClosedTickets(pStore, True)
    TestAssert.AssertDateNear DateSerial(2026, 9, 15), folder.ArchivedFor, 0
    Teardown
End Sub

Public Sub Test_Archive_NonTicketFolderIsSkipped()
    Setup
    pStore.AddActive "Misc"
    pStore.AddActive "#abc (x)"

    TestAssert.AssertEqual 0, TicketService.ArchiveClosedTickets(pStore, True)
    TestAssert.AssertEqual 0, pTransport.RequestCount, "no API call for folders that are not ticket folders"
    TestAssert.AssertEqual 2, pStore.ActiveCount
    Teardown
End Sub

Public Sub Test_Archive_FolderIsRenamedToCurrentSubject()
    Dim folder As FakeTicketFolder

    Setup
    AnswerOpenAndClosedTickets
    Set folder = pStore.AddActive("#20101 (Old subject)")

    TicketService.ArchiveClosedTickets pStore, True
    TestAssert.AssertEqual "#20101 (Example open ticket)", folder.FolderName
    Teardown
End Sub

Public Sub Test_Archive_OneRequestPerOpenTicket()
    Setup
    AnswerOpenAndClosedTickets
    pStore.AddActive "#20101 (x)"

    TicketService.ArchiveClosedTickets pStore, True
    TestAssert.AssertEqual 1, pTransport.RequestCount, "state and subject come from one request"
    Teardown
End Sub

Public Sub Test_Archive_StopsAskingTheApiWhenItIsDown()
    Dim i As Long
    Dim fifth As FakeTicketFolder

    Setup
    For i = 1 To 6
        If i = 5 Then
            Set fifth = pStore.AddActive("#" & (1000 + i) & " (x)")
            fifth.HasClosedStatusMail = True
            fifth.ClosedStatusMailDate = DateSerial(2026, 8, 20)
        Else
            pStore.AddActive "#" & (1000 + i) & " (x)"
        End If
    Next i

    TestAssert.AssertEqual 1, TicketService.ArchiveClosedTickets(pStore, True), "only the folder with a closed mail"
    TestAssert.AssertEqual 3, pTransport.RequestCount, "API asked 3 times, then given up for this run"
    TestAssert.AssertTrue fifth.IsArchived, "the fallback still works for the remaining folders"
    TestAssert.AssertDateNear DateSerial(2026, 8, 20), fifth.ArchivedFor, 0
    Teardown
End Sub

Public Sub Test_Archive_ApiRecoveringResetsTheFailureCount()
    Setup
    AnswerOpenAndClosedTickets
    ' 2 unavailable, 1 answered, 2 unavailable: never 3 in a row
    pStore.AddActive "#1001 (x)"
    pStore.AddActive "#1002 (x)"
    pStore.AddActive "#20101 (x)"
    pStore.AddActive "#1003 (x)"
    pStore.AddActive "#1004 (x)"

    TicketService.ArchiveClosedTickets pStore, False
    TestAssert.AssertEqual 5, pTransport.RequestCount, "all five folders were asked"
    Teardown
End Sub

Public Sub Test_Archive_ErrorInOneFolderDoesNotStopTheRun()
    Dim broken As FakeTicketFolder
    Dim good As FakeTicketFolder

    Setup
    AnswerOpenAndClosedTickets
    Set broken = pStore.AddActive("#20101 (x)")
    broken.FailOnNameRead = True
    Set good = pStore.AddActive("#20102 (Example)")

    TestAssert.AssertEqual 1, TicketService.ArchiveClosedTickets(pStore, True)
    TestAssert.AssertTrue good.IsArchived
    TestAssert.AssertTrue InStr(TicketLog.Text(), "simulated folder error") > 0, "the problem is logged"
    Teardown
End Sub

Public Sub Test_Archive_ResultIsLogged()
    Setup
    AnswerOpenAndClosedTickets
    pStore.AddActive "#20102 (Example)"
    pStore.AddActive "#20101 (x)"

    TicketService.ArchiveClosedTickets pStore, True
    TestAssert.AssertTrue InStr(TicketLog.Text(), "1 of 2 folders archived") > 0
    Teardown
End Sub

Public Sub Test_Archive_ReportsProgress()
    Dim sink As FakeProgressSink

    Setup
    AnswerOpenAndClosedTickets
    pStore.AddActive "#20102 (Example)"
    pStore.AddActive "#20101 (x)"
    pStore.AddActive "Misc"
    Set sink = New FakeProgressSink

    TestAssert.AssertEqual 1, TicketService.ArchiveClosedTickets(pStore, True, sink)
    TestAssert.AssertEqual 1, sink.BeginCount
    TestAssert.AssertEqual 3, sink.TotalReported, "total = number of active folders"
    TestAssert.AssertEqual 3, sink.Messages.Count, "one message per folder"
    TestAssert.AssertTrue InStr(sink.Messages(1), "1 of 3") > 0, "message names the position"
    TestAssert.AssertEqual 1, sink.FinishCount, "finished exactly once"
    Teardown
End Sub

Public Sub Test_Archive_ProgressFinishesEvenWhenAFolderFails()
    Dim sink As FakeProgressSink
    Dim broken As FakeTicketFolder

    Setup
    AnswerOpenAndClosedTickets
    Set broken = pStore.AddActive("#20101 (x)")
    broken.FailOnNameRead = True
    pStore.AddActive "#20102 (Example)"
    Set sink = New FakeProgressSink

    TicketService.ArchiveClosedTickets pStore, True, sink
    TestAssert.AssertEqual 2, sink.Messages.Count, "both folders were reported"
    TestAssert.AssertTrue InStr(sink.Messages(1), "(unknown folder)") > 0, "an unreadable folder name does not break the report"
    TestAssert.AssertEqual 1, sink.FinishCount
    Teardown
End Sub

Public Sub Test_Archive_ProgressIsReportedForAnEmptyStore()
    Dim sink As FakeProgressSink

    Setup
    Set sink = New FakeProgressSink

    TestAssert.AssertEqual 0, TicketService.ArchiveClosedTickets(pStore, True, sink)
    TestAssert.AssertEqual 0, sink.TotalReported
    TestAssert.AssertEqual 1, sink.FinishCount
    Teardown
End Sub
