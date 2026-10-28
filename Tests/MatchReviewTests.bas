Attribute VB_Name = "MatchReviewTests"
' MatchReviewTests v0.1.0
'
' Tests of the confirmation level: reviewing proposals (accept / reject / skip / cancel), the hint
' question and the memory of rejected mails. FakeMatchPrompt stands in for the dialogs; no Outlook,
' no network. Ticket 20101 is open, 20102 is closed, 30001 is found by its subject.

Option Explicit

Private pSink As FakeAuditSink
Private pStore As FakeTicketStore
Private pIndex As FakeTicketIndex
Private pTransport As FakeNinjaTransport
Private pPrompt As FakeMatchPrompt
Private pRejections As FakeRejectionStore

Private Sub Setup()
    TicketLog.DisableFileLogging
    TicketLog.Clear
    NinjaApi.ResetState
    TicketAudit.ResetState
    MatchLevels.UseAll LevelConfirm
    MatchProposals.Clear
    MatchReview.ResetState

    Set pRejections = New FakeRejectionStore
    MatchRejections.UseStore pRejections

    Set pSink = New FakeAuditSink
    TicketAudit.UseSink pSink

    Set pStore = New FakeTicketStore
    Set pTransport = New FakeNinjaTransport
    NinjaApi.UseTransport pTransport
    pTransport.AddResponse "ticketing/ticket/20101", 200, Fixtures.TicketOpen()
    pTransport.AddResponse "ticketing/ticket/20102", 200, Fixtures.TicketClosed()
    pTransport.AddResponse "ticketing/ticket/30001", 200, Fixtures.TicketOpen()

    Set pIndex = New FakeTicketIndex
    pIndex.AddTicket 20101, "Example open ticket", True
    pIndex.AddTicket 20102, "Example closed ticket", False
    pIndex.AddTicket 30001, "Example order total differs from expected value", True

    Set pPrompt = New FakeMatchPrompt
End Sub

Private Sub Teardown()
    MatchLevels.ResetState
    MatchProposals.Clear
    MatchReview.ResetState
    MatchRejections.ResetState
    TicketAudit.ResetState
    NinjaApi.ResetState
    Set pSink = Nothing
    Set pStore = Nothing
    Set pIndex = Nothing
    Set pTransport = Nothing
    Set pPrompt = Nothing
    Set pRejections = Nothing
End Sub

Private Function Foreign(ByVal subject As String, ByVal messageKey As String) As FakeMail
    Dim mail As FakeMail

    Set mail = New FakeMail
    mail.SenderName = "Some Customer"
    mail.SenderDomain = "example.com"
    mail.Subject = subject
    mail.ConversationId = "conv-" & messageKey
    mail.MessageKey = messageKey
    Set Foreign = mail
End Function

' Queues a proposal for ticket 20101 and returns the mail.
Private Function Propose(ByVal messageKey As String) As FakeMail
    Dim mail As FakeMail

    Set mail = Foreign("Ticket 20101", messageKey)
    TicketService.MatchForeignMail mail, pIndex, pStore, True
    Set Propose = mail
End Function

Private Function Review() As Long
    Review = MatchReview.ReviewPending(pPrompt, pStore)
End Function

' --- accept ---------------------------------------------------------------------------

Public Sub Test_Review_AcceptMovesTheMailAndDropsTheProposal()
    Dim mail As FakeMail

    Setup
    Set mail = Propose("aa11aa11")
    pPrompt.SetDecisions DecisionAccept

    TestAssert.AssertEqual 1, Review()
    TestAssert.AssertFalse mail.MovedTo Is Nothing, "mail is in the ticket folder"
    TestAssert.AssertEqual "#20101 (Example open ticket)", pStore.ActiveAt(1).FolderName
    TestAssert.AssertEqual 0, MatchProposals.Count()
    TestAssert.AssertEqual 1, pPrompt.Shown.Count
    TestAssert.AssertTrue InStr(pSink.LastLine(), """event"":""user_confirmed""") > 0, pSink.LastLine()
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""moved""") > 0
    TestAssert.AssertTrue InStr(pSink.LastLine(), """source"":""review""") > 0
    TestAssert.AssertTrue InStr(pSink.LastLine(), """subject"":""ticket 20101""") > 0, "the subject of a decided mail is logged"
    Teardown
End Sub

Public Sub Test_Review_AcceptUsesTheExistingFolder()
    Dim mail As FakeMail
    Dim folder As FakeTicketFolder

    Setup
    Set folder = pStore.AddActive("#20101 (Example open ticket)")
    Set mail = Propose("aa11aa11")
    pPrompt.SetDecisions DecisionAccept

    TestAssert.AssertEqual 1, Review()
    TestAssert.AssertTrue mail.MovedTo Is folder
    TestAssert.AssertEqual 0, pStore.CreatedCount
    Teardown
End Sub

Public Sub Test_Review_AcceptWhenTheTicketWasClosedMeanwhile()
    Dim mail As FakeMail

    Setup
    Set mail = Propose("aa11aa11")
    pTransport.AddResponse "ticketing/ticket/20101", 200, Fixtures.TicketClosed()
    pPrompt.SetDecisions DecisionAccept

    TestAssert.AssertEqual 0, Review()
    TestAssert.AssertTrue mail.MovedTo Is Nothing, "closed in the meantime: not moved"
    TestAssert.AssertEqual 0, MatchProposals.Count(), "the stale proposal is dropped"
    TestAssert.AssertTrue InStr(pSink.LastLine(), """event"":""user_confirmed""") > 0
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""stays""") > 0, pSink.LastLine()
    Teardown
End Sub

Public Sub Test_Review_AcceptWhenTheMailCannotBeMovedAnyMore()
    Dim mail As FakeMail

    Setup
    Set mail = Propose("aa11aa11")
    pStore.FailMoves = True
    pPrompt.SetDecisions DecisionAccept

    TestAssert.AssertEqual 0, Review()
    TestAssert.AssertEqual 0, MatchProposals.Count()
    TestAssert.AssertTrue InStr(TicketLog.Text(), "could not be moved") > 0, TicketLog.Text()
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""stays""") > 0
    Teardown
End Sub

' --- reject ----------------------------------------------------------------------------

Public Sub Test_Review_RejectRemembersTheMail()
    Dim mail As FakeMail

    Setup
    Set mail = Propose("aa11aa11")
    pPrompt.SetDecisions DecisionReject

    TestAssert.AssertEqual 0, Review()
    TestAssert.AssertTrue mail.MovedTo Is Nothing
    TestAssert.AssertEqual 0, MatchProposals.Count()
    TestAssert.AssertTrue MatchRejections.IsRejected("aa11aa11")
    TestAssert.AssertEqual 1, pRejections.Keys.Count, "written to the store"
    TestAssert.AssertTrue InStr(pSink.LastLine(), """event"":""user_rejected""") > 0, pSink.LastLine()
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""rejected""") > 0
    Teardown
End Sub

Public Sub Test_Review_RejectedMailIsNotProposedAgain()
    Setup
    Propose "aa11aa11"
    pPrompt.SetDecisions DecisionReject
    Review

    TestAssert.AssertEqual MatchFound, TicketService.MatchForeignMail(Foreign("Ticket 20101", "aa11aa11"), pIndex, pStore, True)
    TestAssert.AssertEqual 0, MatchProposals.Count(), "same mail: no new proposal"
    TestAssert.AssertTrue InStr(TicketLog.Text(), "rejected this mail before") > 0

    TicketService.MatchForeignMail Foreign("Ticket 20101", "bb22bb22"), pIndex, pStore, True
    TestAssert.AssertEqual 1, MatchProposals.Count(), "another mail is still proposed"
    Teardown
End Sub

Public Sub Test_Rejections_SurviveARestart()
    Dim saved As FakeRejectionStore

    Setup
    MatchRejections.Reject "aa11aa11"
    Set saved = pRejections

    MatchRejections.ResetState            ' simulates a new Outlook session
    TestAssert.AssertFalse MatchRejections.IsRejected("aa11aa11"), "no store: nothing known"
    MatchRejections.UseStore saved
    TestAssert.AssertTrue MatchRejections.IsRejected("aa11aa11"), "loaded from the store"
    TestAssert.AssertFalse MatchRejections.IsRejected("cc33cc33")
    Teardown
End Sub

Public Sub Test_Rejections_EmptyKeyAndBrokenStore()
    Setup
    MatchRejections.Reject ""
    TestAssert.AssertEqual 0, pRejections.Keys.Count
    TestAssert.AssertFalse MatchRejections.IsRejected("")

    pRejections.FailOnAdd = True
    MatchRejections.Reject "dd44dd44"
    TestAssert.AssertTrue MatchRejections.IsRejected("dd44dd44"), "holds in memory even if the file fails"
    TestAssert.AssertTrue InStr(TicketLog.Text(), "could not be saved") > 0
    Teardown
End Sub

Public Sub Test_Rejections_FileStoreRoundTrip()
    Dim store As FileRejectionStore
    Dim iface As IRejectionStore
    Dim fso As Object
    Dim base As String
    Dim keys As Collection

    base = Environ$("TEMP") & "\ninja_rej_test_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd() * 100000))
    Set store = New FileRejectionStore
    store.BaseFolder = base & "\sub"
    Set iface = store

    iface.AddKey "AA11AA11"
    iface.AddKey "bb22bb22"
    iface.AddKey "not a key"
    iface.AddKey "123"

    Set keys = iface.LoadKeys()
    TestAssert.AssertEqual 2, keys.Count, "only valid keys are stored"
    TestAssert.AssertEqual "aa11aa11", keys(1), "keys are lower case"
    TestAssert.AssertEqual "bb22bb22", keys(2)

    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    fso.DeleteFolder base, True
End Sub

' --- skip, cancel, mixed ---------------------------------------------------------------------

Public Sub Test_Review_SkipLeavesTheProposal()
    Dim mail As FakeMail

    Setup
    Set mail = Propose("aa11aa11")
    pPrompt.SetDecisions DecisionSkip

    TestAssert.AssertEqual 0, Review()
    TestAssert.AssertTrue mail.MovedTo Is Nothing
    TestAssert.AssertEqual 1, MatchProposals.Count()
    TestAssert.AssertEqual 0, pSink.Count - 1, "only the proposal itself was logged"
    Teardown
End Sub

Public Sub Test_Review_CancelDecidesNothing()
    Setup
    Propose "aa11aa11"
    Propose "bb22bb22"
    pPrompt.Cancel = True

    TestAssert.AssertEqual 0, Review()
    TestAssert.AssertEqual 2, MatchProposals.Count()
    TestAssert.AssertEqual 0, pRejections.Keys.Count
    Teardown
End Sub

Public Sub Test_Review_MixedDecisionsInOrder()
    Dim m1 As FakeMail, m2 As FakeMail, m3 As FakeMail

    Setup
    Set m1 = Propose("aa11aa11")
    Set m2 = Propose("bb22bb22")
    Set m3 = Propose("cc33cc33")
    pPrompt.SetDecisions DecisionAccept, DecisionReject, DecisionSkip

    TestAssert.AssertEqual 1, Review()
    TestAssert.AssertFalse m1.MovedTo Is Nothing, "first: accepted"
    TestAssert.AssertTrue m2.MovedTo Is Nothing, "second: rejected"
    TestAssert.AssertTrue MatchRejections.IsRejected("bb22bb22")
    TestAssert.AssertEqual 1, MatchProposals.Count(), "third: skipped, still waiting"
    TestAssert.AssertEqual "cc33cc33", MatchProposals.Item(1).Mail.MessageKey
    TestAssert.AssertEqual "aa11aa11", pPrompt.Shown(1).Mail.MessageKey, "shown oldest first"
    TestAssert.AssertTrue InStr(TicketLog.Text(), "1 moved, 0 not moved, 1 rejected, 1 skipped") > 0, TicketLog.Text()
    Teardown
End Sub

Public Sub Test_Review_MissingOrInvalidDecisionsMeanSkip()
    Setup
    Propose "aa11aa11"
    Propose "bb22bb22"
    Propose "cc33cc33"
    pPrompt.SetDecisions DecisionAccept, 7

    TestAssert.AssertEqual 1, Review()
    TestAssert.AssertEqual 2, MatchProposals.Count(), "7 is no decision, the third has none"
    Teardown
End Sub

Public Sub Test_Review_NothingPendingDoesNotAsk()
    Setup
    TestAssert.AssertEqual 0, Review()
    TestAssert.AssertEqual 0, pPrompt.ReviewCount
    TestAssert.AssertEqual 0, MatchReview.ReviewPending(Nothing, pStore)
    Teardown
End Sub

' --- the hint question --------------------------------------------------------------------

Public Sub Test_Offer_AsksOnceForNewProposals()
    Setup
    pPrompt.AnswerAskNow = False
    Propose "aa11aa11"

    MatchReview.OfferReview pPrompt, pStore
    TestAssert.AssertEqual 1, pPrompt.AskCount
    TestAssert.AssertEqual 1, pPrompt.LastPendingCount
    TestAssert.AssertEqual 0, pPrompt.ReviewCount, "no: no review"
    TestAssert.AssertEqual 1, MatchProposals.Count()

    MatchReview.OfferReview pPrompt, pStore
    TestAssert.AssertEqual 1, pPrompt.AskCount, "nothing new: not asked again"
    Teardown
End Sub

Public Sub Test_Offer_YesStartsTheReview()
    Dim mail As FakeMail

    Setup
    pPrompt.AnswerAskNow = True
    Set mail = Propose("aa11aa11")
    pPrompt.SetDecisions DecisionAccept

    MatchReview.OfferReview pPrompt, pStore
    TestAssert.AssertEqual 1, pPrompt.ReviewCount
    TestAssert.AssertFalse mail.MovedTo Is Nothing
    Teardown
End Sub

Public Sub Test_Offer_NothingPendingNeverAsks()
    Setup
    MatchReview.OfferReview pPrompt, pStore
    MatchReview.OfferReview Nothing, pStore
    TestAssert.AssertEqual 0, pPrompt.AskCount
    Teardown
End Sub

Public Sub Test_Offer_NewProposalsWaitForTheCooldown()
    Setup
    pPrompt.AnswerAskNow = False
    MatchReview.TestNow = DateSerial(2026, 10, 9) + TimeSerial(10, 0, 0)

    Propose "aa11aa11"
    MatchReview.OfferReview pPrompt, pStore
    TestAssert.AssertEqual 1, pPrompt.AskCount

    Propose "bb22bb22"
    MatchReview.TestNow = MatchReview.TestNow + 5 / 1440#
    MatchReview.OfferReview pPrompt, pStore
    TestAssert.AssertEqual 1, pPrompt.AskCount, "5 minutes later: still cooling down"

    MatchReview.TestNow = MatchReview.TestNow + 11 / 1440#
    MatchReview.OfferReview pPrompt, pStore
    TestAssert.AssertEqual 2, pPrompt.AskCount, "16 minutes later: asks about the new proposal"
    TestAssert.AssertEqual 2, pPrompt.LastPendingCount
    Teardown
End Sub

Public Sub Test_Offer_NoSecondDialogWhileTheFirstIsOpen()
    Setup
    pPrompt.AnswerAskNow = True
    Propose "aa11aa11"
    pPrompt.SetDecisions DecisionSkip
    Set pPrompt.ReentryStore = pStore

    MatchReview.OfferReview pPrompt, pStore
    TestAssert.AssertEqual 1, pPrompt.AskCount, "the offer made from inside the dialog is ignored"
    TestAssert.AssertEqual 1, pPrompt.ReviewCount
    Teardown
End Sub

' --- end to end ----------------------------------------------------------------------------

Public Sub Test_EndToEnd_Level2ThenAccept()
    Dim bySubject As FakeMail
    Dim byId As FakeMail

    Setup
    Set bySubject = Foreign("AW: Example order total differs from expected value", "ee55ee55")
    Set byId = Foreign("Ticket 20101", "ff66ff66")

    TestAssert.AssertEqual MatchFound, TicketService.MatchForeignMail(bySubject, pIndex, pStore, True)
    TestAssert.AssertEqual MatchFound, TicketService.MatchForeignMail(byId, pIndex, pStore, True)
    TestAssert.AssertEqual 2, MatchProposals.Count()
    TestAssert.AssertTrue bySubject.MovedTo Is Nothing, "nothing moves before the review"

    pPrompt.SetDecisions DecisionAccept, DecisionAccept
    TestAssert.AssertEqual 2, Review()
    TestAssert.AssertEqual 2, pStore.ActiveCount, "one folder per ticket"
    TestAssert.AssertFalse bySubject.MovedTo Is Nothing
    TestAssert.AssertFalse byId.MovedTo Is Nothing
    TestAssert.AssertEqual 0, MatchProposals.Count()
    Teardown
End Sub
