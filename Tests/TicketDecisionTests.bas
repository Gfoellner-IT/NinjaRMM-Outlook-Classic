Attribute VB_Name = "TicketDecisionTests"
' TicketDecisionTests v0.1.0
'
' Truth table of TicketDecision.DecideArchive (4 API states x mail text x fallback flag).

Option Explicit

Private Sub Check(ByVal state As NinjaTicketState, ByVal mailClosed As Boolean, ByVal fallback As Boolean, ByVal expected As ArchiveDecision)
    TestAssert.AssertEqual expected, TicketDecision.DecideArchive(state, mailClosed, fallback), _
        "state=" & state & " mailClosed=" & mailClosed & " fallback=" & fallback
End Sub

Public Sub Test_Decision_ClosedIsArchivedByApi()
    Check NinjaTicketClosed, False, False, ArchiveByApi
    Check NinjaTicketClosed, False, True, ArchiveByApi
    Check NinjaTicketClosed, True, False, ArchiveByApi
    Check NinjaTicketClosed, True, True, ArchiveByApi
End Sub

Public Sub Test_Decision_NotFoundIsArchivedByApi()
    Check NinjaTicketNotFound, False, False, ArchiveByApi
    Check NinjaTicketNotFound, False, True, ArchiveByApi
    Check NinjaTicketNotFound, True, False, ArchiveByApi
    Check NinjaTicketNotFound, True, True, ArchiveByApi
End Sub

Public Sub Test_Decision_OpenIsNeverArchived()
    ' Even if an old mail says "Geschlossen" (ticket reopened later): the API wins.
    Check NinjaTicketOpen, False, False, ArchiveNo
    Check NinjaTicketOpen, False, True, ArchiveNo
    Check NinjaTicketOpen, True, False, ArchiveNo
    Check NinjaTicketOpen, True, True, ArchiveNo
End Sub

Public Sub Test_Decision_UnavailableUsesFallbackOnlyWhenEnabledAndMailSaysClosed()
    Check NinjaTicketUnavailable, True, True, ArchiveByMailFallback
    Check NinjaTicketUnavailable, True, False, ArchiveNo
    Check NinjaTicketUnavailable, False, True, ArchiveNo
    Check NinjaTicketUnavailable, False, False, ArchiveNo
End Sub

Public Sub Test_Decision_NeedsMailCheckOnlyWhenApiUnavailableAndFallbackOn()
    TestAssert.AssertTrue TicketDecision.NeedsMailCheck(NinjaTicketUnavailable, True)
    TestAssert.AssertFalse TicketDecision.NeedsMailCheck(NinjaTicketUnavailable, False)
    TestAssert.AssertFalse TicketDecision.NeedsMailCheck(NinjaTicketOpen, True)
    TestAssert.AssertFalse TicketDecision.NeedsMailCheck(NinjaTicketClosed, True)
    TestAssert.AssertFalse TicketDecision.NeedsMailCheck(NinjaTicketNotFound, True)
End Sub

Public Sub Test_Decision_FallbackIsOnByDefault()
    TestAssert.AssertTrue TicketConfig.ENABLE_OFFLINE_FALLBACK, "agreed default: fallback on"
End Sub
