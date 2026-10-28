Attribute VB_Name = "NinjaTicketIndexTests"
' NinjaTicketIndexTests v0.1.0
'
' Tests of the board query logic (NinjaTicketLogic) and of NinjaTicketIndex with a fake
' transport. No network access.

Option Explicit

Private Const BOARD_PATH As String = "ticketing/trigger/board/2/run"

Private pFake As FakeNinjaTransport
Private pIndex As NinjaTicketIndex

Private Sub Setup()
    Set pFake = New FakeNinjaTransport
    Set pIndex = New NinjaTicketIndex
    pIndex.UseTransport pFake
    pIndex.TestNow = DateSerial(2026, 10, 7) + TimeSerial(10, 0, 0)
End Sub

Private Function Idx() As ITicketIndex
    Set Idx = pIndex
End Function

' Request bodies exactly as the index builds them.
Private Function OpenBody(ByVal size As Long, ByVal cursor As Long) As String
    OpenBody = NinjaTicketLogic.BuildBoardRunBody(size, cursor, "status", "not_equals", "6000")
End Function

Private Function ExternalBody(ByVal number As String) As String
    ExternalBody = NinjaTicketLogic.BuildBoardRunBody(50, 0, "summary", "contains", number)
End Function

' --- NinjaTicketLogic ---------------------------------------------------------

Public Sub Test_BoardBody()
    TestAssert.AssertEqual "{""pageSize"":200}", NinjaTicketLogic.BuildBoardRunBody(200, 0, "", "", "")
    TestAssert.AssertEqual "{""pageSize"":200,""lastCursorId"":400}", NinjaTicketLogic.BuildBoardRunBody(200, 400, "", "", "")
    TestAssert.AssertEqual "{""pageSize"":50,""filters"":[{""field"":""summary"",""operator"":""contains"",""value"":""900000105""}]}", _
        NinjaTicketLogic.BuildBoardRunBody(50, 0, "summary", "contains", "900000105")
    TestAssert.AssertEqual "{""pageSize"":10,""lastCursorId"":10,""filters"":[{""field"":""status"",""operator"":""not_equals"",""value"":""6000""}]}", _
        NinjaTicketLogic.BuildBoardRunBody(10, 10, "status", "not_equals", "6000")
End Sub

Public Sub Test_BoardBody_EscapesValue()
    TestAssert.AssertEqual "{""pageSize"":5,""filters"":[{""field"":""summary"",""operator"":""contains"",""value"":""a\""b\\c""}]}", _
        NinjaTicketLogic.BuildBoardRunBody(5, 0, "summary", "contains", "a""b\c")
End Sub

Public Sub Test_ParseBoardRun_Page()
    Dim tickets As Collection
    Dim cursor As Long
    Dim rowCount As Long

    TestAssert.AssertTrue NinjaTicketLogic.ParseBoardRun(Fixtures.BoardRunPage1(), tickets, cursor, rowCount)
    TestAssert.AssertEqual 2, tickets.Count
    TestAssert.AssertEqual 2, rowCount
    TestAssert.AssertEqual 2, cursor
    TestAssert.AssertEqual 301, tickets(1).Id
    TestAssert.AssertEqual "Example printer problem in the warehouse", tickets(1).Summary
    TestAssert.AssertTrue tickets(1).IsOpen
    TestAssert.AssertTrue tickets(2).IsOpen, "status Neu is open"
End Sub

Public Sub Test_ParseBoardRun_ClosedAndResolved()
    Dim tickets As Collection
    Dim cursor As Long
    Dim rowCount As Long

    TestAssert.AssertTrue NinjaTicketLogic.ParseBoardRun(Fixtures.BoardRunVendor(), tickets, cursor, rowCount)
    TestAssert.AssertEqual 3, tickets.Count
    TestAssert.AssertTrue tickets(1).IsOpen
    TestAssert.AssertFalse tickets(2).IsOpen, "statusId 6000 is closed"

    TestAssert.AssertTrue NinjaTicketLogic.ParseBoardRun(Fixtures.BoardRunPage2(), tickets, cursor, rowCount)
    TestAssert.AssertTrue tickets(1).IsOpen, "resolved (5000) is not closed"
End Sub

Public Sub Test_ParseBoardRun_EmptyAndInvalid()
    Dim tickets As Collection
    Dim cursor As Long
    Dim rowCount As Long

    TestAssert.AssertTrue NinjaTicketLogic.ParseBoardRun(Fixtures.BoardRunEmpty(), tickets, cursor, rowCount)
    TestAssert.AssertEqual 0, tickets.Count
    TestAssert.AssertEqual 0, rowCount

    TestAssert.AssertFalse NinjaTicketLogic.ParseBoardRun("", tickets, cursor, rowCount)
    TestAssert.AssertFalse NinjaTicketLogic.ParseBoardRun("not json", tickets, cursor, rowCount)
    TestAssert.AssertFalse NinjaTicketLogic.ParseBoardRun("[]", tickets, cursor, rowCount)
    TestAssert.AssertFalse NinjaTicketLogic.ParseBoardRun("{""other"":1}", tickets, cursor, rowCount)
    TestAssert.AssertEqual 0, tickets.Count
End Sub

' --- GetTicket ----------------------------------------------------------------

Public Sub Test_Index_GetTicket()
    Dim t As MatchTicket

    Setup
    pFake.AddResponse "ticketing/ticket/20101", 200, Fixtures.TicketOpen()
    pFake.AddResponse "ticketing/ticket/20102", 200, Fixtures.TicketClosed()
    pFake.AddResponse "ticketing/ticket/99999", 404, ""
    pFake.AddResponse "ticketing/ticket/50000", 500, ""

    Set t = Idx().GetTicket(20101)
    TestAssert.AssertTrue Not t Is Nothing
    TestAssert.AssertEqual 20101, t.Id
    TestAssert.AssertEqual "Example open ticket", t.Summary
    TestAssert.AssertTrue t.IsOpen

    Set t = Idx().GetTicket(20102)
    TestAssert.AssertTrue Not t Is Nothing
    TestAssert.AssertFalse t.IsOpen

    TestAssert.AssertTrue Idx().GetTicket(99999) Is Nothing, "unknown ticket"
    TestAssert.AssertTrue Idx().GetTicket(50000) Is Nothing, "server error"
    TestAssert.AssertTrue Len(pIndex.LastError) > 0
    TestAssert.AssertTrue Idx().GetTicket(0) Is Nothing, "invalid id"
    TestAssert.AssertTrue Idx().GetTicket(12345) Is Nothing, "no response"
End Sub

' --- FindByExternalNumber -----------------------------------------------------

Public Sub Test_Index_ExternalNumber_FiltersAndCaches()
    Dim found As Collection

    Setup
    pFake.AddPostResponse BOARD_PATH, ExternalBody("900000101"), 200, Fixtures.BoardRunVendor()

    Set found = Idx().FindByExternalNumber("900000101")
    TestAssert.AssertEqual 2, found.Count, "the unrelated row is dropped locally"
    TestAssert.AssertEqual 401, found(1).Id
    TestAssert.AssertTrue found(1).IsOpen
    TestAssert.AssertFalse found(2).IsOpen
    TestAssert.AssertEqual 1, pFake.PostCount

    Set found = Idx().FindByExternalNumber("900000101")
    TestAssert.AssertEqual 2, found.Count
    TestAssert.AssertEqual 1, pFake.PostCount, "second lookup comes from the cache"
End Sub

Public Sub Test_Index_ExternalNumber_CacheExpires()
    Setup
    pFake.AddPostResponse BOARD_PATH, ExternalBody("900000101"), 200, Fixtures.BoardRunVendor()

    Idx().FindByExternalNumber "900000101"
    pIndex.TestNow = pIndex.TestNow + 9 / 1440#
    Idx().FindByExternalNumber "900000101"
    TestAssert.AssertEqual 1, pFake.PostCount, "9 minutes: still cached"

    pIndex.TestNow = pIndex.TestNow + 2 / 1440#
    Idx().FindByExternalNumber "900000101"
    TestAssert.AssertEqual 2, pFake.PostCount, "11 minutes: queried again"
End Sub

Public Sub Test_Index_ExternalNumber_NoTicketIsCachedToo()
    Setup
    pFake.AddPostResponse BOARD_PATH, ExternalBody("900000110"), 200, Fixtures.BoardRunEmpty()

    TestAssert.AssertEqual 0, Idx().FindByExternalNumber("900000110").Count
    TestAssert.AssertEqual 0, Idx().FindByExternalNumber("900000110").Count
    TestAssert.AssertEqual 1, pFake.PostCount
End Sub

Public Sub Test_Index_ExternalNumber_FailureIsNotCached()
    Setup
    pFake.AddPostFailure BOARD_PATH, ExternalBody("900000101"), "timeout"

    TestAssert.AssertEqual 0, Idx().FindByExternalNumber("900000101").Count
    TestAssert.AssertTrue InStr(pIndex.LastError, "timeout") > 0
    pFake.AddPostResponse BOARD_PATH, ExternalBody("900000101"), 200, Fixtures.BoardRunVendor()
    TestAssert.AssertEqual 2, Idx().FindByExternalNumber("900000101").Count, "retried after the failure"
End Sub

Public Sub Test_Index_ExternalNumber_HttpErrorAndGarbage()
    Setup
    pFake.AddPostResponse BOARD_PATH, ExternalBody("1"), 500, ""
    pFake.AddPostResponse BOARD_PATH, ExternalBody("2"), 200, "garbage"

    TestAssert.AssertEqual 0, Idx().FindByExternalNumber("1").Count
    TestAssert.AssertTrue InStr(pIndex.LastError, "500") > 0
    TestAssert.AssertEqual 0, Idx().FindByExternalNumber("2").Count
    TestAssert.AssertTrue Len(pIndex.LastError) > 0
End Sub

' --- FindOpenBySubject --------------------------------------------------------

Private Sub AddOpenPages(ByVal size As Long)
    pFake.AddPostResponse BOARD_PATH, OpenBody(size, 0), 200, Fixtures.BoardRunPage1()
    pFake.AddPostResponse BOARD_PATH, OpenBody(size, 2), 200, Fixtures.BoardRunPage2()
End Sub

Public Sub Test_Index_OpenBySubject_PagesAndGroups()
    Dim found As Collection

    Setup
    pIndex.PageSize = 2
    AddOpenPages 2

    Set found = Idx().FindOpenBySubject("example subject with prefix")
    TestAssert.AssertEqual 1, found.Count, "TICKET prefix and AW: are normalized away"
    TestAssert.AssertEqual 302, found(1).Id
    TestAssert.AssertEqual 2, pFake.PostCount, "two pages"

    Set found = Idx().FindOpenBySubject("example vendor ticket")
    TestAssert.AssertEqual 1, found.Count, "third ticket from page 2"
    TestAssert.AssertEqual 303, found(1).Id

    TestAssert.AssertEqual 0, Idx().FindOpenBySubject("something else entirely").Count
    TestAssert.AssertEqual 2, pFake.PostCount, "all lookups served from the cache"
End Sub

Public Sub Test_Index_OpenBySubject_SinglePageStopsEarly()
    Setup
    pIndex.PageSize = 200
    pFake.AddPostResponse BOARD_PATH, OpenBody(200, 0), 200, Fixtures.BoardRunPage1()

    TestAssert.AssertEqual 1, Idx().FindOpenBySubject("example printer problem in the warehouse").Count
    TestAssert.AssertEqual 1, pFake.PostCount, "page shorter than the page size is the last one"
End Sub

Public Sub Test_Index_OpenBySubject_Reloads()
    Setup
    pIndex.PageSize = 2
    AddOpenPages 2

    Idx().FindOpenBySubject "x"
    pIndex.TestNow = pIndex.TestNow + 11 / 1440#
    Idx().FindOpenBySubject "x"
    TestAssert.AssertEqual 4, pFake.PostCount, "reloaded after the cache lifetime"
End Sub

Public Sub Test_Index_OpenBySubject_FailureKeepsOldCache()
    Setup
    pIndex.PageSize = 2
    AddOpenPages 2

    TestAssert.AssertEqual 1, Idx().FindOpenBySubject("example vendor ticket").Count
    pIndex.TestNow = pIndex.TestNow + 11 / 1440#
    pFake.AddPostFailure BOARD_PATH, OpenBody(2, 0), "offline"

    TestAssert.AssertEqual 1, Idx().FindOpenBySubject("example vendor ticket").Count, "stale list is better than none"
    TestAssert.AssertTrue InStr(pIndex.LastError, "offline") > 0
End Sub

Public Sub Test_Index_OpenBySubject_FailureOnFirstLoadGivesNothingAndRetries()
    Setup
    pIndex.PageSize = 2
    pFake.AddPostFailure BOARD_PATH, OpenBody(2, 0), "offline"

    TestAssert.AssertEqual 0, Idx().FindOpenBySubject("example vendor ticket").Count
    AddOpenPages 2
    TestAssert.AssertEqual 1, Idx().FindOpenBySubject("example vendor ticket").Count, "next lookup tries again"
End Sub

Public Sub Test_Index_OpenBySubject_SecondPageFailureDiscardsPartialList()
    Setup
    pIndex.PageSize = 2
    pFake.AddPostResponse BOARD_PATH, OpenBody(2, 0), 200, Fixtures.BoardRunPage1()
    pFake.AddPostFailure BOARD_PATH, OpenBody(2, 2), "offline"

    TestAssert.AssertEqual 0, Idx().FindOpenBySubject("example printer problem in the warehouse").Count, "partial list is not used"
End Sub

' --- conversations ------------------------------------------------------------

Public Sub Test_Index_Conversation()
    Setup
    TestAssert.AssertEqual 0, Idx().FindByConversation("c1")
    pIndex.RememberConversation "c1", 20106
    TestAssert.AssertEqual 20106, Idx().FindByConversation("c1")
    pIndex.RememberConversation "c1", 11200
    TestAssert.AssertEqual 11200, Idx().FindByConversation("c1"), "latest assignment wins"
    pIndex.RememberConversation "", 5
    pIndex.RememberConversation "c2", 0
    TestAssert.AssertEqual 0, Idx().FindByConversation("c2")
    pIndex.ClearCache
    TestAssert.AssertEqual 11200, Idx().FindByConversation("c1"), "ClearCache keeps conversations"
End Sub

' --- together with the matcher ---------------------------------------------------

Public Sub Test_Index_WithMatcher_ExternalNumberOnlyOneOpenTicket()
    Dim id As Long
    Dim rule As MatchRule

    Setup
    pFake.AddPostResponse BOARD_PATH, ExternalBody("900000101"), 200, Fixtures.BoardRunVendor()

    TestAssert.AssertEqual MatchFound, TicketMatcher.MatchMail("AW: [TicketNr. 900000101] Example", "", "", "vendor.example", False, Idx(), id, rule)
    TestAssert.AssertEqual 401, id, "closed duplicate is ignored, the open ticket wins"
    TestAssert.AssertEqual RuleExternalNumber, rule
End Sub

Public Sub Test_Index_WithMatcher_ApiDownMeansNoMatch()
    Dim id As Long
    Dim rule As MatchRule

    Setup
    pFake.AddPostFailure BOARD_PATH, ExternalBody("900000101"), "offline"

    TestAssert.AssertEqual MatchNone, TicketMatcher.MatchMail("AW: [TicketNr. 900000101] Example", "", "", "vendor.example", False, Idx(), id, rule)
    TestAssert.AssertEqual 0, id
End Sub

Public Sub Test_Index_WithMatcher_SubjectRule()
    Dim id As Long
    Dim rule As MatchRule

    Setup
    pIndex.PageSize = 2
    AddOpenPages 2

    TestAssert.AssertEqual MatchFound, TicketMatcher.MatchMail("RE: Example subject with prefix", "", "", "example.com", False, Idx(), id, rule)
    TestAssert.AssertEqual 302, id
    TestAssert.AssertEqual RuleSubject, rule
End Sub
