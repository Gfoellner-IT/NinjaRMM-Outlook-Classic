Attribute VB_Name = "TicketMatcherTests"
' TicketMatcherTests v0.1.0
'
' Unit tests for TicketMatcher (pure logic, FakeTicketIndex instead of NinjaOne).
' Subjects and identifiers are synthetic examples, not copied from a production inbox.

Option Explicit

Private Function NewIndex() As FakeTicketIndex
    Dim idx As FakeTicketIndex

    Set idx = New FakeTicketIndex
    idx.AddTicket 20103, "CSM | [TicketNr. 900000105] - Example invoice processing problem", True
    idx.AddTicket 20106, "WG: Example order total differs from expected value", True
    idx.AddTicket 20107, "TICKET (B" & ChrW$(252) & "ro - AT) / Example interface change request", True
    idx.AddTicket 20108, "Example warehouse configuration request", True
    idx.AddTicket 20104, "CSM|[TicketNr. 900000104] - Example copying problem", True
    idx.AddTicket 20105, "CSM|[TicketNr. 900000104] - B" & ChrW$(252) & "ro: Example copy of delivery data", True
    idx.AddTicket 7001, "[TicketNr. 900000107] alte Anfrage", False
    idx.AddTicket 7002, "[TicketNr. 900000107] alte Anfrage wieder geoeffnet", False
    idx.AddTicket 9001, "Netzwerkausfall", True
    Set NewIndex = idx
End Function

Private Function Run(ByVal idx As FakeTicketIndex, ByVal subjectText As String, ByRef ticketId As Long, ByRef rule As MatchRule, _
                     Optional ByVal bodyText As String = "", Optional ByVal conversationId As String = "", _
                     Optional ByVal senderDomain As String = "example.com", Optional ByVal hasAutoHeader As Boolean = False) As MatchOutcome
    Run = TicketMatcher.MatchMail(subjectText, bodyText, conversationId, senderDomain, hasAutoHeader, idx, ticketId, rule)
End Function

' --- extraction -------------------------------------------------------------

Public Sub Test_ExtractNinjaIds()
    Dim ids As Collection

    Set ids = TicketMatcher.ExtractNinjaIds("Ticket 20108")
    TestAssert.AssertEqual 1, ids.Count
    TestAssert.AssertEqual 20108, ids(1)

    TestAssert.AssertEqual 20108, TicketMatcher.ExtractNinjaIds("Ticket #20108 bitte pruefen")(1)
    TestAssert.AssertEqual 20108, TicketMatcher.ExtractNinjaIds("Re: Ticket-Nr. 20108")(1)
    TestAssert.AssertEqual 20109, TicketMatcher.ExtractNinjaIds("Ticket-Ausloeser: https://x.example/#/ticketing/ticket/20109")(1)
    TestAssert.AssertEqual 3621, TicketMatcher.ExtractNinjaIds("[example] (#3621) x")(1)
    TestAssert.AssertEqual 1, TicketMatcher.ExtractNinjaIds("Ticket 20108 und Ticket 20108").Count, "distinct"
    TestAssert.AssertEqual 2, TicketMatcher.ExtractNinjaIds("Ticket 20108 und Ticket 20103").Count
End Sub

Public Sub Test_ExtractNinjaIds_IgnoresVendorAndOthers()
    TestAssert.AssertEqual 0, TicketMatcher.ExtractNinjaIds("[TicketNr. 900000105] Fehler").Count, "vendor number is 9 digits"
    TestAssert.AssertEqual 0, TicketMatcher.ExtractNinjaIds("Rechnung #38379").Count, "bare # is not a ticket"
    TestAssert.AssertEqual 0, TicketMatcher.ExtractNinjaIds("Ticket 12").Count, "too short"
    TestAssert.AssertEqual 0, TicketMatcher.ExtractNinjaIds("").Count
End Sub

Public Sub Test_ExtractExternalNumbers()
    Dim nums As Collection

    Set nums = TicketMatcher.ExtractExternalNumbers("AW: [TicketNr. 900000106] B" & ChrW$(252) & "ro: item235 - error")
    TestAssert.AssertEqual 1, nums.Count
    TestAssert.AssertEqual "900000106", nums(1)

    TestAssert.AssertEqual "900000103", TicketMatcher.ExtractExternalNumbers("[CSM-TicketNr. 900000103] - Problem")(1)
    TestAssert.AssertEqual 2, TicketMatcher.ExtractExternalNumbers("[TicketNr. 900000101] [TicketNr. 900000102]").Count
    TestAssert.AssertEqual 1, TicketMatcher.ExtractExternalNumbers("[TicketNr. 900000101] [TicketNr. 900000101]").Count, "distinct"
    TestAssert.AssertEqual 0, TicketMatcher.ExtractExternalNumbers("Ticket 20108").Count
    TestAssert.AssertEqual 0, TicketMatcher.ExtractExternalNumbers("[GitHub] Sudo email verification code").Count
End Sub

Public Sub Test_NormalizeSubject()
    TestAssert.AssertEqual "example order total differs from expected value", TicketMatcher.NormalizeSubject("AW: WG: Example order total differs from expected value")
    TestAssert.AssertEqual "example interface change request", _
        TicketMatcher.NormalizeSubject("TICKET (B" & ChrW$(252) & "ro - AT) / Example interface change request")
    TestAssert.AssertEqual "item235 - error", _
        TicketMatcher.NormalizeSubject("AW: [TicketNr. 900000106] B" & ChrW$(252) & "ro: item235 - error")
    TestAssert.AssertEqual "example invoice processing problem", _
        TicketMatcher.NormalizeSubject("CSM | [TicketNr. 900000105] - Example invoice processing problem")
    TestAssert.AssertEqual "a b", TicketMatcher.NormalizeSubject("  RE:   a    b ")
    TestAssert.AssertEqual "", TicketMatcher.NormalizeSubject("")
End Sub

Public Sub Test_IsSpecificSubject()
    TestAssert.AssertTrue TicketMatcher.IsSpecificSubject("example order total differs from expected value")
    TestAssert.AssertFalse TicketMatcher.IsSpecificSubject("netzwerkausfall"), "one word"
    TestAssert.AssertFalse TicketMatcher.IsSpecificSubject("kurz und knapp"), "too short"
    TestAssert.AssertFalse TicketMatcher.IsSpecificSubject("")
End Sub

Public Sub Test_IsExcluded()
    TestAssert.AssertTrue TicketMatcher.IsExcluded("github.com", False)
    TestAssert.AssertTrue TicketMatcher.IsExcluded("GitHub.com", False), "case-insensitive"
    TestAssert.AssertTrue TicketMatcher.IsExcluded("example.com", True), "auto header"
    TestAssert.AssertFalse TicketMatcher.IsExcluded("example.com", False)
    TestAssert.AssertFalse TicketMatcher.IsExcluded("vendor.example", False)
    TestAssert.AssertFalse TicketMatcher.IsExcluded("", False)
End Sub

' --- rule 1: NinjaOne id ------------------------------------------------------

Public Sub Test_Match_NinjaId_InSubject()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchFound, Run(idx, "Ticket 20108", id, rule)
    TestAssert.AssertEqual 20108, id
    TestAssert.AssertEqual RuleNinjaId, rule
End Sub

Public Sub Test_Match_NinjaId_InBody_WhenSubjectHasNone()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchFound, Run(idx, "Frage dazu", id, rule, "siehe Ticket #20103 bitte")
    TestAssert.AssertEqual 20103, id
End Sub

Public Sub Test_Match_NinjaId_Unknown_FallsThrough()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchNone, Run(idx, "Ticket 99999", id, rule)
    TestAssert.AssertEqual 0, id
End Sub

Public Sub Test_Match_NinjaId_Several_IsAmbiguous()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchAmbiguous, Run(idx, "Ticket 20108 und Ticket 20103", id, rule)
    TestAssert.AssertEqual 0, id
End Sub

' --- rule 2: external number --------------------------------------------------

Public Sub Test_Match_ExternalNumber_Found()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchFound, Run(idx, "AW: [TicketNr. 900000105] B" & ChrW$(252) & "ro: Fehlerhafte Verbuchung", id, rule, , , "vendor.example")
    TestAssert.AssertEqual 20103, id
    TestAssert.AssertEqual RuleExternalNumber, rule
End Sub

Public Sub Test_Match_ExternalNumber_NotInAnyTicket_StaysInInbox()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    ' the conversation is known and the subject would match, but an unknown vendor number wins
    idx.AssignConversation "conv1", 20106
    TestAssert.AssertEqual MatchNone, Run(idx, "AW: [TicketNr. 900000106] Example order total differs from expected value", id, rule, , "conv1", "vendor.example")
    TestAssert.AssertEqual 0, id
    TestAssert.AssertEqual RuleExternalNumber, rule
End Sub

Public Sub Test_Match_ExternalNumber_TwoTickets_OneOpen_Wins()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    idx.AddTicket 7003, "[TicketNr. 900000108] geschlossen", False
    idx.AddTicket 7004, "[TicketNr. 900000108] offen", True
    TestAssert.AssertEqual MatchFound, Run(idx, "[TicketNr. 900000108] Rueckfrage", id, rule)
    TestAssert.AssertEqual 7004, id
End Sub

Public Sub Test_Match_ExternalNumber_TwoOpenTickets_IsAmbiguous()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchAmbiguous, Run(idx, "[TicketNr. 900000104] example copying question", id, rule)
    TestAssert.AssertEqual 0, id
End Sub

Public Sub Test_Match_ExternalNumber_TwoClosedTickets_IsAmbiguous()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchAmbiguous, Run(idx, "[TicketNr. 900000107] Rueckfrage", id, rule)
End Sub

Public Sub Test_Match_ExternalNumber_SingleClosedTicket_IsFound()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    idx.AddTicket 7010, "[TicketNr. 900000109] erledigt", False
    TestAssert.AssertEqual MatchFound, Run(idx, "[TicketNr. 900000109] Nachfrage", id, rule)
    TestAssert.AssertEqual 7010, id
End Sub

Public Sub Test_Match_ExternalNumber_TwoNumbers_IsAmbiguous()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchAmbiguous, Run(idx, "[TicketNr. 900000105] [TicketNr. 900000104] x", id, rule)
End Sub

' --- rule 3: conversation -----------------------------------------------------

Public Sub Test_Match_Conversation()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    idx.AssignConversation "conv-42", 20106
    TestAssert.AssertEqual MatchFound, Run(idx, "Kurz", id, rule, , "conv-42")
    TestAssert.AssertEqual 20106, id
    TestAssert.AssertEqual RuleConversation, rule
    TestAssert.AssertEqual MatchNone, Run(idx, "Kurz", id, rule, , "conv-other")
End Sub

' --- rule 4: subject ----------------------------------------------------------

Public Sub Test_Match_Subject_IgnoresPrefixes()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchFound, Run(idx, "AW: Example order total differs from expected value", id, rule)
    TestAssert.AssertEqual 20106, id
    TestAssert.AssertEqual RuleSubject, rule

    TestAssert.AssertEqual MatchFound, Run(idx, "AW: Example interface change request", id, rule)
    TestAssert.AssertEqual 20107, id
End Sub

Public Sub Test_Match_Subject_GenericSubjectNeverMatches()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchNone, Run(idx, "AW: Netzwerkausfall", id, rule)
    TestAssert.AssertEqual 0, id
End Sub

Public Sub Test_Match_Subject_TwoOpenTickets_IsAmbiguous()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    idx.AddTicket 11200, "Example order total differs from expected value", True
    TestAssert.AssertEqual MatchAmbiguous, Run(idx, "AW: Example order total differs from expected value", id, rule)
    TestAssert.AssertEqual 0, id
End Sub

Public Sub Test_Match_Subject_ClosedTicketDoesNotMatch()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = New FakeTicketIndex
    idx.AddTicket 5000, "Drucker im Lager funktioniert nicht", False
    TestAssert.AssertEqual MatchNone, Run(idx, "AW: Drucker im Lager funktioniert nicht", id, rule)
End Sub

Public Sub Test_Match_Subject_NoTicket()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchNone, Run(idx, "Pro Brass Konzert Karten fuer naechste Woche", id, rule)
    TestAssert.AssertEqual RuleNone, rule
End Sub

' --- exclusions and priority --------------------------------------------------

Public Sub Test_Match_ExcludedMail_NeverQueriesIndex()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchNone, Run(idx, "Ticket 20108", id, rule, , , "github.com")
    TestAssert.AssertEqual MatchNone, Run(idx, "Ticket 20108", id, rule, , , "example.com", True)
    TestAssert.AssertEqual 0, idx.LookupCount
End Sub

Public Sub Test_Match_NinjaId_BeatsExternalNumber()
    Dim idx As FakeTicketIndex, id As Long, rule As MatchRule

    Set idx = NewIndex()
    TestAssert.AssertEqual MatchFound, Run(idx, "[TicketNr. 900000104] zu Ticket 20108", id, rule)
    TestAssert.AssertEqual 20108, id
    TestAssert.AssertEqual RuleNinjaId, rule
End Sub

Public Sub Test_Match_NoIndex()
    Dim id As Long, rule As MatchRule

    TestAssert.AssertEqual MatchNone, TicketMatcher.MatchMail("Ticket 20108", "", "", "example.com", False, Nothing, id, rule)
End Sub

' --- header detection and log text -------------------------------------------------

Public Sub Test_HasAutoHeaderLine()
    TestAssert.AssertTrue TicketMatcher.HasAutoHeaderLine("Received: x" & vbCrLf & "List-Unsubscribe: <mailto:a@b.example>" & vbCrLf & "Subject: y")
    TestAssert.AssertTrue TicketMatcher.HasAutoHeaderLine("List-Id: <x.example>")
    TestAssert.AssertTrue TicketMatcher.HasAutoHeaderLine("X-Auto-Response-Suppress: All")
    TestAssert.AssertTrue TicketMatcher.HasAutoHeaderLine("Precedence: bulk")
    TestAssert.AssertTrue TicketMatcher.HasAutoHeaderLine("Auto-Submitted: auto-generated")
    TestAssert.AssertFalse TicketMatcher.HasAutoHeaderLine("Auto-Submitted: no")
    TestAssert.AssertFalse TicketMatcher.HasAutoHeaderLine("Precedence: normal")
    TestAssert.AssertFalse TicketMatcher.HasAutoHeaderLine("Subject: List-Id mentioned in the text")
    TestAssert.AssertFalse TicketMatcher.HasAutoHeaderLine("From: a@b.example" & vbCrLf & "In-Reply-To: <1@x>")
    TestAssert.AssertFalse TicketMatcher.HasAutoHeaderLine("")
End Sub

Public Sub Test_DescribeMatch()
    TestAssert.AssertEqual "Match found rule=ExternalNumber ticket=#20103 domain=vendor.example subject='[TicketNr. 1] x'", _
        TicketMatcher.DescribeMatch(MatchFound, RuleExternalNumber, 20103, "vendor.example", "[TicketNr. 1] x")
    TestAssert.AssertEqual "Match none domain=x.example subject_hash=" & AuditLogic.Hash32Hex("hello"), _
        TicketMatcher.DescribeMatch(MatchNone, RuleNone, 0, "x.example", "hello"), "no match: only the hash of the subject"
    TestAssert.AssertEqual "Match ambiguous rule=Subject subject='a b'", _
        TicketMatcher.DescribeMatch(MatchAmbiguous, RuleSubject, 0, "", "a" & vbCrLf & "b")
    TestAssert.AssertEqual "Match found subject='" & String$(60, "x") & "'", TicketMatcher.DescribeMatch(MatchFound, RuleNone, 0, "", String$(100, "x")), "subject is cut"
    TestAssert.AssertEqual "Match none", TicketMatcher.DescribeMatch(MatchNone, RuleNone, 0, "", ""), "nothing to hash"
End Sub
