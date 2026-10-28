Attribute VB_Name = "TicketMatchServiceTests"
' TicketMatchServiceTests v0.1.0
'
' Tests of TicketService.MatchForeignMail (log-only mode, moving, conversation memory) with an
' in-memory store, FakeTicketIndex and a fake NinjaOne transport. No Outlook, no network.
' Ticket 20101 is open ("Example open ticket"), 20102 is closed.

Option Explicit

Private pStore As FakeTicketStore
Private pIndex As FakeTicketIndex
Private pTransport As FakeNinjaTransport

Private Sub Setup()
    TicketLog.DisableFileLogging
    MatchLevels.UseAll LevelAuto
    MatchProposals.Clear
    TicketLog.Clear
    NinjaApi.ResetState

    Set pStore = New FakeTicketStore
    Set pTransport = New FakeNinjaTransport
    NinjaApi.UseTransport pTransport
    pTransport.AddResponse "ticketing/ticket/20101", 200, Fixtures.TicketOpen()
    pTransport.AddResponse "ticketing/ticket/20102", 200, Fixtures.TicketClosed()

    Set pIndex = New FakeTicketIndex
    pIndex.AddTicket 20101, "Example open ticket", True
    pIndex.AddTicket 20102, "Example closed ticket", False
    pIndex.AddTicket 20001, "CSM | [TicketNr. 900000101] - Vendor problem", True
End Sub

Private Sub Teardown()
    MatchLevels.ResetState
    MatchProposals.Clear
    NinjaApi.ResetState
    Set pStore = Nothing
    Set pIndex = Nothing
    Set pTransport = Nothing
End Sub

Private Function Foreign(ByVal subject As String) As FakeMail
    Dim mail As FakeMail

    Set mail = New FakeMail
    mail.SenderName = "Some Customer"
    mail.SenderDomain = "example.com"
    mail.Subject = subject
    Set Foreign = mail
End Function

Private Function Match(ByVal mail As FakeMail, ByVal moveEnabled As Boolean) As MatchOutcome
    Match = TicketService.MatchForeignMail(mail, pIndex, pStore, moveEnabled)
End Function

' --- ignored and log-only -----------------------------------------------------

Public Sub Test_MatchMail_TicketSenderIsIgnored()
    Dim mail As FakeMail

    Setup
    Set mail = Foreign("Ticket 20101")
    mail.SenderName = TicketConfig.AllowedSender()

    TestAssert.AssertEqual MatchNone, Match(mail, True)
    TestAssert.AssertEqual 0, pIndex.LookupCount, "not even looked up"
    TestAssert.AssertEqual "", TicketLog.Text()
    Teardown
End Sub

Public Sub Test_MatchMail_LogOnlyNeverMoves()
    Dim mail As FakeMail

    Setup
    Set mail = Foreign("Rueckfrage zu Ticket 20101")

    TestAssert.AssertEqual MatchFound, Match(mail, False)
    TestAssert.AssertTrue mail.MovedTo Is Nothing, "log-only mode leaves the mail alone"
    TestAssert.AssertEqual 0, pStore.ActiveCount, "no folder is created either"
    TestAssert.AssertTrue InStr(TicketLog.Text(), "Match found rule=NinjaId ticket=#20101") > 0, TicketLog.Text()
    TestAssert.AssertTrue InStr(TicketLog.Text(), "(log only)") > 0
    Teardown
End Sub

Public Sub Test_MatchMail_LogLineNamesRuleDomainAndSubject()
    Setup
    Match Foreign("Ticket 20101 bitte pruefen"), False

    TestAssert.AssertTrue InStr(TicketLog.Text(), "domain=example.com") > 0
    TestAssert.AssertTrue InStr(TicketLog.Text(), "subject='Ticket 20101 bitte pruefen'") > 0
    Teardown
End Sub

' --- moving -------------------------------------------------------------------

Public Sub Test_MatchMail_MovesIntoExistingFolder()
    Dim mail As FakeMail
    Dim folder As FakeTicketFolder

    Setup
    Set folder = pStore.AddActive("#20101 (Example open ticket)")
    Set mail = Foreign("Re: Ticket 20101")

    TestAssert.AssertEqual MatchFound, Match(mail, True)
    TestAssert.AssertTrue mail.MovedTo Is folder
    TestAssert.AssertEqual 0, pStore.CreatedCount
    TestAssert.AssertTrue InStr(TicketLog.Text(), "(moved)") > 0
    Teardown
End Sub

Public Sub Test_MatchMail_CreatesFolderNamedAfterTicketSubject()
    Dim mail As FakeMail

    Setup
    Set mail = Foreign("Ticket 20101")

    TestAssert.AssertEqual MatchFound, Match(mail, True)
    TestAssert.AssertEqual 1, pStore.CreatedCount
    TestAssert.AssertEqual "#20101 (Example open ticket)", pStore.ActiveAt(1).FolderName
    TestAssert.AssertTrue mail.MovedTo Is pStore.ActiveAt(1)
    Teardown
End Sub

Public Sub Test_MatchMail_RestoresArchivedFolderOfOpenTicket()
    Dim mail As FakeMail
    Dim archived As FakeTicketFolder

    Setup
    Set archived = pStore.AddArchived("#20101 (Example open ticket)")
    Set mail = Foreign("Ticket 20101")

    TestAssert.AssertEqual MatchFound, Match(mail, True)
    TestAssert.AssertEqual 1, pStore.RestoredCount
    TestAssert.AssertTrue mail.MovedTo Is archived
    Teardown
End Sub

Public Sub Test_MatchMail_ClosedTicketStaysInInbox()
    Dim mail As FakeMail

    Setup
    Set mail = Foreign("Ticket 20102")

    TestAssert.AssertEqual MatchFound, Match(mail, True)
    TestAssert.AssertTrue mail.MovedTo Is Nothing
    TestAssert.AssertEqual 0, pStore.ActiveCount
    TestAssert.AssertTrue InStr(TicketLog.Text(), "(stays in the inbox)") > 0
    Teardown
End Sub

Public Sub Test_MatchMail_UnreachableApiStaysInInbox()
    Dim mail As FakeMail

    Setup
    pIndex.AddTicket 12345, "Not answered by the fake API", True
    Set mail = Foreign("Ticket 12345")

    TestAssert.AssertEqual MatchFound, Match(mail, True)
    TestAssert.AssertTrue mail.MovedTo Is Nothing, "state of the ticket unknown: do not move"
    Teardown
End Sub

Public Sub Test_MatchMail_NoMatchOrAmbiguousNeverMoves()
    Dim mail As FakeMail

    Setup
    Set mail = Foreign("Hallo zusammen")
    TestAssert.AssertEqual MatchNone, Match(mail, True)
    TestAssert.AssertTrue mail.MovedTo Is Nothing

    Set mail = Foreign("Ticket 20101 und Ticket 20102")
    TestAssert.AssertEqual MatchAmbiguous, Match(mail, True)
    TestAssert.AssertTrue mail.MovedTo Is Nothing
    TestAssert.AssertTrue InStr(TicketLog.Text(), "Match ambiguous") > 0
    TestAssert.AssertEqual 0, pStore.ActiveCount
    Teardown
End Sub

Public Sub Test_MatchMail_UnknownVendorNumberStaysInInbox()
    Dim mail As FakeMail

    Setup
    Set mail = Foreign("AW: [TicketNr. 399999999] Etwas ohne internes Ticket")
    mail.SenderDomain = "vendor.example"

    TestAssert.AssertEqual MatchNone, Match(mail, True)
    TestAssert.AssertTrue mail.MovedTo Is Nothing
    TestAssert.AssertTrue InStr(TicketLog.Text(), "rule=ExternalNumber") > 0
    Teardown
End Sub

Public Sub Test_MatchMail_ExcludedMailIsNotMoved()
    Dim mail As FakeMail

    Setup
    Set mail = Foreign("Ticket 20101")
    mail.HasAutoHeader = True

    TestAssert.AssertEqual MatchNone, Match(mail, True)
    TestAssert.AssertTrue mail.MovedTo Is Nothing
    Teardown
End Sub

' --- conversation memory ------------------------------------------------------

Public Sub Test_MatchMail_ThreadFollowsAfterVendorNumberMatch()
    Dim first As FakeMail
    Dim reply As FakeMail

    Setup
    Set first = Foreign("[TicketNr. 900000101] Vendor problem")
    first.SenderDomain = "vendor.example"
    first.ConversationId = "thread-1"
    Set reply = Foreign("AW: Vendor problem ohne Nummer")
    reply.SenderDomain = "vendor.example"
    reply.ConversationId = "thread-1"

    TestAssert.AssertEqual MatchFound, Match(first, False)
    TestAssert.AssertEqual MatchFound, Match(reply, False), "the reply has no number but the same conversation"
    TestAssert.AssertTrue InStr(TicketLog.Text(), "rule=Conversation ticket=#20001") > 0, TicketLog.Text()
    Teardown
End Sub

Public Sub Test_MatchMail_SubjectGuessIsNotRemembered()
    Setup
    pIndex.AddTicket 30001, "Example order total differs from expected value", True
    Dim mail As FakeMail
    Set mail = Foreign("AW: Example order total differs from expected value")
    mail.ConversationId = "thread-2"

    TestAssert.AssertEqual MatchFound, Match(mail, False)
    TestAssert.AssertTrue InStr(TicketLog.Text(), "rule=Subject") > 0
    TestAssert.AssertEqual 0, CLng(TicketMatcherLookup("thread-2")), "only decisive rules assign the conversation"
    Teardown
End Sub

Private Function TicketMatcherLookup(ByVal conversationId As String) As Long
    Dim idx As ITicketIndex

    Set idx = pIndex
    TicketMatcherLookup = idx.FindByConversation(conversationId)
End Function
