Attribute VB_Name = "NinjaTicketLogicTests"
' NinjaTicketLogicTests v0.1.0
'
' Unit tests for the ticket state, subject and closing date logic.

Option Explicit

' ---------------------------------------------------------------------------
'  Ticket state
' ---------------------------------------------------------------------------

Public Sub Test_TicketState_Open()
    TestAssert.AssertEqual NinjaTicketOpen, NinjaTicketLogic.ParseTicketState(200, Fixtures.TicketOpen())
End Sub

Public Sub Test_TicketState_Closed()
    TestAssert.AssertEqual NinjaTicketClosed, NinjaTicketLogic.ParseTicketState(200, Fixtures.TicketClosed())
End Sub

Public Sub Test_TicketState_DeletedCountsAsClosed()
    TestAssert.AssertEqual NinjaTicketClosed, NinjaTicketLogic.ParseTicketState(200, Fixtures.TicketDeleted())
End Sub

Public Sub Test_TicketState_NotFound()
    TestAssert.AssertEqual NinjaTicketNotFound, NinjaTicketLogic.ParseTicketState(404, ""), "empty body"
    TestAssert.AssertEqual NinjaTicketNotFound, NinjaTicketLogic.ParseTicketState(404, Fixtures.TicketNotFound()), "TICKET_NOT_FOUND"
    TestAssert.AssertEqual NinjaTicketNotFound, NinjaTicketLogic.ParseTicketState(404, "<html>404</html>"), "non-JSON body"
End Sub

Public Sub Test_TicketState_Other404IsUnavailable()
    TestAssert.AssertEqual NinjaTicketUnavailable, _
        NinjaTicketLogic.ParseTicketState(404, "{""resultCode"":""ENDPOINT_NOT_FOUND""}"), "a different problem than a missing ticket"
End Sub

Public Sub Test_TicketState_ErrorsAreUnavailable()
    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaTicketLogic.ParseTicketState(0, ""), "no response"
    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaTicketLogic.ParseTicketState(401, ""), "unauthorized"
    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaTicketLogic.ParseTicketState(403, ""), "forbidden"
    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaTicketLogic.ParseTicketState(500, ""), "server error"
    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaTicketLogic.ParseTicketState(503, ""), "service unavailable"
End Sub

Public Sub Test_TicketState_UnreadableAnswerIsUnavailable()
    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaTicketLogic.ParseTicketState(200, ""), "empty"
    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaTicketLogic.ParseTicketState(200, "not json"), "garbage"
    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaTicketLogic.ParseTicketState(200, "[]"), "array instead of object"
    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaTicketLogic.ParseTicketState(200, "{""id"":1}"), "no status"
    TestAssert.AssertEqual NinjaTicketUnavailable, NinjaTicketLogic.ParseTicketState(200, "{""status"":{""statusId"":""6000""}}"), "statusId is not a number"
End Sub

' ---------------------------------------------------------------------------
'  Subject
' ---------------------------------------------------------------------------

Public Sub Test_TicketSubject()
    TestAssert.AssertEqual "Example open ticket", NinjaTicketLogic.ParseTicketSubject(Fixtures.TicketOpen())
    TestAssert.AssertEqual "Example", NinjaTicketLogic.ParseTicketSubject(Fixtures.TicketClosed())
    TestAssert.AssertEqual "", NinjaTicketLogic.ParseTicketSubject(""), "empty"
    TestAssert.AssertEqual "", NinjaTicketLogic.ParseTicketSubject("not json"), "garbage"
    TestAssert.AssertEqual "", NinjaTicketLogic.ParseTicketSubject("{""id"":1}"), "no subject"
    TestAssert.AssertEqual "", NinjaTicketLogic.ParseTicketSubject("{""subject"":null}"), "null subject"
End Sub

' ---------------------------------------------------------------------------
'  Closing date (all times UTC)
' ---------------------------------------------------------------------------

Public Sub Test_UnixToDate()
    TestAssert.AssertDateNear DateSerial(1970, 1, 1), NinjaTicketLogic.UnixToDate(0), 0, "epoch"
    TestAssert.AssertDateNear DateSerial(2026, 10, 6) + TimeSerial(14, 1, 28), NinjaTicketLogic.UnixToDate(1791295288), 0, "known ticket close time"
    TestAssert.AssertDateNear DateSerial(2026, 10, 6) + TimeSerial(14, 1, 28), NinjaTicketLogic.UnixToDate(1791295288.695002), 0, "fraction is cut off"
    TestAssert.AssertDateNear DateSerial(2026, 9, 1) + TimeSerial(13, 55, 6), NinjaTicketLogic.UnixToDate(1788270906), 0, "second known time"
End Sub

Public Sub Test_ClosedDate_ClosedManually()
    ' Ticket 20102: closed by a technician, no automation entry in the log.
    TestAssert.AssertDateNear DateSerial(2026, 10, 6) + TimeSerial(14, 1, 28), NinjaTicketLogic.ParseClosedDate(Fixtures.LogClosedManual()), 0
End Sub

Public Sub Test_ClosedDate_ClosedByAutomation()
    TestAssert.AssertDateNear DateSerial(2026, 9, 1) + TimeSerial(13, 55, 6), NinjaTicketLogic.ParseClosedDate(Fixtures.LogClosedAutomation()), 0
End Sub

Public Sub Test_ClosedDate_EnglishStatusName()
    TestAssert.AssertDateNear DateSerial(2026, 10, 6) + TimeSerial(14, 1, 28), NinjaTicketLogic.ParseClosedDate(Fixtures.LogClosedEnglish()), 0
End Sub

Public Sub Test_ClosedDate_ReopenedAndClosedAgainUsesNewest()
    ' Closed 2026-10-05, reopened, closed again 2026-10-07 19:06:40.
    TestAssert.AssertDateNear DateSerial(2026, 10, 7) + TimeSerial(19, 6, 40), NinjaTicketLogic.ParseClosedDate(Fixtures.LogReopened()), 0
End Sub

Public Sub Test_ClosedDate_DeletedTicketUsesDeleteEntry()
    TestAssert.AssertDateNear NinjaTicketLogic.UnixToDate(1791500000), NinjaTicketLogic.ParseClosedDate(Fixtures.LogDeleted()), 0
End Sub

Public Sub Test_ClosedDate_NoClosingEntryGivesZero()
    TestAssert.AssertEqual 0#, CDbl(NinjaTicketLogic.ParseClosedDate(Fixtures.LogOpenOnly())), "only open and new entries"
    TestAssert.AssertEqual 0#, CDbl(NinjaTicketLogic.ParseClosedDate("[]")), "empty log"
    TestAssert.AssertEqual 0#, CDbl(NinjaTicketLogic.ParseClosedDate("")), "empty text"
    TestAssert.AssertEqual 0#, CDbl(NinjaTicketLogic.ParseClosedDate("not json")), "garbage"
    TestAssert.AssertEqual 0#, CDbl(NinjaTicketLogic.ParseClosedDate("{}")), "object instead of array"
End Sub

Public Sub Test_ClosedDate_ToleratesStrangeEntries()
    Dim log As String

    ' entries that are not objects, have no changeDiff, a null changeDiff or no createTime must not break the result
    log = "[1,""x"",null,{""type"":""SAVE""},{""changeDiff"":null},{""changeDiff"":{""status"":null}}," & _
          "{""changeDiff"":{""status"":{""new"":""Geschlossen""}}}," & _
          "{""changeDiff"":{""status"":{""new"":""Geschlossen""}},""createTime"":1791295288}]"
    TestAssert.AssertDateNear DateSerial(2026, 10, 6) + TimeSerial(14, 1, 28), NinjaTicketLogic.ParseClosedDate(log), 0
End Sub

Public Sub Test_IsClosedStatusName()
    TestAssert.AssertTrue NinjaTicketLogic.IsClosedStatusName("Geschlossen")
    TestAssert.AssertTrue NinjaTicketLogic.IsClosedStatusName("Closed")
    TestAssert.AssertTrue NinjaTicketLogic.IsClosedStatusName("closed"), "case-insensitive"
    TestAssert.AssertFalse NinjaTicketLogic.IsClosedStatusName("Offen")
    TestAssert.AssertFalse NinjaTicketLogic.IsClosedStatusName(""), "empty"
    TestAssert.AssertFalse NinjaTicketLogic.IsClosedStatusName("Geschl"), "no partial matches"
End Sub
