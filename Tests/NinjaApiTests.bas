Attribute VB_Name = "NinjaApiTests"
' NinjaApiTests v0.1.0
'
' Tests of NinjaApi with a fake transport. No network access.

Option Explicit

Private Const TICKET_OPEN As Long = 20101
Private Const TICKET_CLOSED As Long = 20102

Private pFake As FakeNinjaTransport

Private Sub Setup()
    NinjaApi.ResetState
    Set pFake = New FakeNinjaTransport
    NinjaApi.UseTransport pFake
End Sub

Private Sub Teardown()
    NinjaApi.ResetState
    Set pFake = Nothing
End Sub

Public Sub Test_Api_StateOpen()
    Setup
    pFake.AddResponse "ticketing/ticket/20101", 200, Fixtures.TicketOpen()

    TestAssert.AssertEqual NinjaTicketOpen, NinjaApi.GetTicketState(TICKET_OPEN)
    TestAssert.AssertFalse NinjaApi.IsTicketClosed(TICKET_OPEN)
    Teardown
End Sub

Public Sub Test_Api_StateClosed()
    Setup
    pFake.AddResponse "ticketing/ticket/20102", 200, Fixtures.TicketClosed()

    TestAssert.AssertEqual NinjaTicketClosed, NinjaApi.GetTicketState(TICKET_CLOSED)
    TestAssert.AssertTrue NinjaApi.IsTicketClosed(TICKET_CLOSED)
    Teardown
End Sub

Public Sub Test_Api_NotFoundCountsAsClosedButIsDistinguishable()
    Setup
    pFake.AddResponse "ticketing/ticket/99999999", 404, ""

    TestAssert.AssertEqual NinjaTicketNotFound, NinjaApi.GetTicketState(99999999)
    TestAssert.AssertTrue NinjaApi.IsTicketClosed(99999999), "deleted tickets are archived like before"
    Teardown
End Sub

Public Sub Test_Api_UnavailableIsNeverClosed()
    Setup
    pFake.AddFailure "ticketing/ticket/20101", "Token request failed: no response from the token endpoint"
    pFake.AddResponse "ticketing/ticket/20102", 503, ""

    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaApi.GetTicketState(TICKET_OPEN), "offline"
    TestAssert.AssertTrue InStr(NinjaApi.LastError, "No response") > 0, "last error is kept for the log"
    TestAssert.AssertFalse NinjaApi.IsTicketClosed(TICKET_OPEN), "offline is not closed"
    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaApi.GetTicketState(TICKET_CLOSED), "server error"
    TestAssert.AssertFalse NinjaApi.IsTicketClosed(TICKET_CLOSED), "5xx is not closed"
    Teardown
End Sub

Public Sub Test_Api_UnknownPathIsUnavailable()
    Setup

    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaApi.GetTicketState(123)
    Teardown
End Sub

Public Sub Test_Api_InvalidTicketIdMakesNoRequest()
    Setup

    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaApi.GetTicketState(0)
    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaApi.GetTicketState(-5)
    TestAssert.AssertEqual "", NinjaApi.GetTicketSubject(0)
    TestAssert.AssertEqual 0, pFake.RequestCount
    Teardown
End Sub

Public Sub Test_Api_Subject()
    Setup
    pFake.AddResponse "ticketing/ticket/20101", 200, Fixtures.TicketOpen()
    pFake.AddResponse "ticketing/ticket/20102", 500, ""

    TestAssert.AssertEqual "Example open ticket", NinjaApi.GetTicketSubject(TICKET_OPEN)
    TestAssert.AssertEqual "", NinjaApi.GetTicketSubject(TICKET_CLOSED), "server error"
    TestAssert.AssertEqual "", NinjaApi.GetTicketSubject(555), "no response"
    Teardown
End Sub

Public Sub Test_Api_ClosedDate()
    Setup
    pFake.AddResponse "ticketing/ticket/20102/log-entry?type=SAVE", 200, Fixtures.LogClosedManual()

    TestAssert.AssertDateNear DateSerial(2026, 10, 6) + TimeSerial(14, 1, 28), NinjaApi.GetTicketClosedDate(TICKET_CLOSED), 0
    TestAssert.AssertTrue pFake.WasRequested("ticketing/ticket/20102/log-entry?type=SAVE"), "log endpoint"
    Teardown
End Sub

Public Sub Test_Api_ClosedDateUnknown()
    Setup
    pFake.AddResponse "ticketing/ticket/20101/log-entry?type=SAVE", 200, Fixtures.LogOpenOnly()
    pFake.AddResponse "ticketing/ticket/20102/log-entry?type=SAVE", 404, ""

    TestAssert.AssertEqual 0#, CDbl(NinjaApi.GetTicketClosedDate(TICKET_OPEN)), "no closing entry"
    TestAssert.AssertEqual 0#, CDbl(NinjaApi.GetTicketClosedDate(TICKET_CLOSED)), "log not found"
    TestAssert.AssertEqual 0#, CDbl(NinjaApi.GetTicketClosedDate(777)), "no response"
    Teardown
End Sub

Public Sub Test_Api_OneRequestPerQuery()
    Setup
    pFake.AddResponse "ticketing/ticket/20101", 200, Fixtures.TicketOpen()

    NinjaApi.GetTicketState TICKET_OPEN
    TestAssert.AssertEqual 1, pFake.RequestCount
    TestAssert.AssertEqual "ticketing/ticket/20101", pFake.LastPath
    Teardown
End Sub
