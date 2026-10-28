Attribute VB_Name = "LearnedRulesTests"
' LearnedRulesTests v0.1.0
'
' Tests of the learned rules (rules file, rule 5 of the matcher, level, index lookup by stem and the
' rule proposals in the report). The JSON text is injected, the real rules file is never read.

Option Explicit

Private Const RULES_JSON As String = "{""version"":1,""rules"":[" & _
    "{""id"":""a"",""domain"":""vendor.example"",""stem"":""item# - error"",""enabled"":true}," & _
    "{""id"":""b"",""domain"":""*"",""stem"":""monatsabschluss lager"",""enabled"":true}," & _
    "{""id"":""c"",""domain"":""off.example"",""stem"":""abgeschaltete regel"",""enabled"":false}," & _
    "{""id"":""d"",""domain"":"""",""stem"":""ohne domain eintrag""}," & _
    "{""id"":""e"",""domain"":""short.example"",""stem"":""kurz""}," & _
    "{""id"":""f"",""domain"":""VENDOR.EXAMPLE"",""stem"":""ITEM# - Report""}" & _
    "]}"

Private pIndex As FakeTicketIndex

Private Sub Setup()
    TicketLog.Clear
    LearnedRules.ResetState
    MatchLevels.ResetState
    Set pIndex = New FakeTicketIndex
End Sub

Private Sub Teardown()
    LearnedRules.ResetState
    MatchLevels.ResetState
    Set pIndex = Nothing
End Sub

Private Function Run(ByVal subject As String, ByVal domain As String, ByRef ticketId As Long, ByRef rule As MatchRule) As MatchOutcome
    Dim idx As ITicketIndex

    Set idx = pIndex
    Run = TicketMatcher.MatchMail(subject, "", "", domain, False, idx, ticketId, rule)
End Function

' --- the rules file --------------------------------------------------------------------

Public Sub Test_Learned_StemStartsWith()
    TestAssert.AssertTrue LearnedRules.StemStartsWith("item# - error #", "item# - error")
    TestAssert.AssertTrue LearnedRules.StemStartsWith("item# - error #", "item#")
    TestAssert.AssertTrue LearnedRules.StemStartsWith("ITEM# - Error", "item# - error"), "case does not matter"
    TestAssert.AssertTrue LearnedRules.StemStartsWith("item# - error", "item# - error")
    TestAssert.AssertFalse LearnedRules.StemStartsWith("item# - error", "item# - error #"), "the prefix is longer than the stem"
    TestAssert.AssertFalse LearnedRules.StemStartsWith("item# - error #", "item# - err"), "whole words only"
    TestAssert.AssertFalse LearnedRules.StemStartsWith("xitem# - error", "item#")
    TestAssert.AssertFalse LearnedRules.StemStartsWith("", "item#")
    TestAssert.AssertFalse LearnedRules.StemStartsWith("item#", "")
End Sub

Public Sub Test_Learned_LoadsOnlyValidEnabledRules()
    Setup
    LearnedRules.UseJson RULES_JSON

    TestAssert.AssertEqual 3, LearnedRules.Count(), "a, b and f: c is off, d has no domain, e is too short"
    TestAssert.AssertTrue InStr(TicketLog.Text(), "2 rule(s) skipped") > 0, TicketLog.Text()
    Teardown
End Sub

Public Sub Test_Learned_MatchingStem()
    Setup
    LearnedRules.UseJson RULES_JSON

    TestAssert.AssertEqual "item# - error", LearnedRules.MatchingStem("vendor.example", "item# - error #")
    TestAssert.AssertEqual "item# - error", LearnedRules.MatchingStem("VENDOR.EXAMPLE", "item# - error #"), "domain is case-insensitive"
    TestAssert.AssertEqual "", LearnedRules.MatchingStem("other.example", "item# - error #"), "another domain"
    TestAssert.AssertEqual "monatsabschluss lager", LearnedRules.MatchingStem("any.example", "monatsabschluss lager #"), "wildcard domain"
    TestAssert.AssertEqual "", LearnedRules.MatchingStem("off.example", "abgeschaltete regel"), "disabled rule"
    TestAssert.AssertEqual "", LearnedRules.MatchingStem("vendor.example", ""), "empty stem"
    TestAssert.AssertEqual "item# - report", LearnedRules.MatchingStem("vendor.example", "item# - report"), "values of the file are lower-cased"
    Teardown
End Sub

Public Sub Test_Learned_BrokenFileMeansNoRules()
    Setup
    LearnedRules.UseJson "this is not json"
    TestAssert.AssertEqual 0, LearnedRules.Count()
    TestAssert.AssertTrue InStr(TicketLog.Text(), "not valid JSON") > 0

    LearnedRules.UseJson "[1,2,3]"
    TestAssert.AssertEqual 0, LearnedRules.Count()

    LearnedRules.UseJson ""
    TestAssert.AssertEqual 0, LearnedRules.Count()

    LearnedRules.UseJson "{""version"":1}"
    TestAssert.AssertEqual 0, LearnedRules.Count(), "no rules list"
    Teardown
End Sub

Public Sub Test_Learned_NoRulesFileInTests()
    LearnedRules.ResetState
    TestAssert.AssertEqual 0, LearnedRules.Count(), "the real rules file is not read while file logging is disabled"
End Sub

' --- rule 5 of the matcher -------------------------------------------------------------------

Public Sub Test_Learned_MatchesTheSingleOpenTicketOfTheGroup()
    Dim id As Long, rule As MatchRule

    Setup
    LearnedRules.UseJson RULES_JSON
    pIndex.AddTicket 40001, "Item235 - Error bei Verbuchung", True
    pIndex.AddTicket 40002, "Etwas ganz anderes im Lager", True

    TestAssert.AssertEqual MatchFound, Run("AW: item235 - error 4711 in form220", "vendor.example", id, rule)
    TestAssert.AssertEqual 40001, id
    TestAssert.AssertEqual RuleLearned, rule
    Teardown
End Sub

Public Sub Test_Learned_TwoOpenTicketsAreAmbiguous()
    Dim id As Long, rule As MatchRule

    Setup
    LearnedRules.UseJson RULES_JSON
    pIndex.AddTicket 40001, "Item235 - Error bei Verbuchung", True
    pIndex.AddTicket 40003, "Item777 - Error beim Druck", True

    TestAssert.AssertEqual MatchAmbiguous, Run("item235 - error 1", "vendor.example", id, rule)
    TestAssert.AssertEqual RuleLearned, rule
    TestAssert.AssertEqual 0, id
    Teardown
End Sub

Public Sub Test_Learned_NoMatchWithoutTicketDomainOrOpenState()
    Dim id As Long, rule As MatchRule

    Setup
    LearnedRules.UseJson RULES_JSON
    TestAssert.AssertEqual MatchNone, Run("item235 - error 1", "vendor.example", id, rule), "no ticket at all"

    pIndex.AddTicket 40001, "Item235 - Error bei Verbuchung", True
    TestAssert.AssertEqual MatchNone, Run("item235 - error 1", "other.example", id, rule), "domain of the rule is vendor.example"

    Setup
    LearnedRules.UseJson RULES_JSON
    pIndex.AddTicket 40004, "Item235 - Error alt", False
    TestAssert.AssertEqual MatchNone, Run("item235 - error 1", "vendor.example", id, rule), "closed ticket"
    Teardown
End Sub

Public Sub Test_Learned_OtherRulesComeFirst()
    Dim id As Long, rule As MatchRule

    Setup
    LearnedRules.UseJson RULES_JSON
    pIndex.AddTicket 40001, "Item235 - Error bei Verbuchung", True
    pIndex.AddTicket 40005, "Item235 - Error bei Verbuchung im Detail", True

    TestAssert.AssertEqual MatchFound, Run("AW: Ticket 40001 item235 - error", "vendor.example", id, rule)
    TestAssert.AssertEqual RuleNinjaId, rule, "the Ninja id rule wins"
    Teardown
End Sub

Public Sub Test_Learned_SwitchedOffLevelSkipsTheRule()
    Dim id As Long, rule As MatchRule

    Setup
    LearnedRules.UseJson RULES_JSON
    pIndex.AddTicket 40001, "Item235 - Error bei Verbuchung", True
    MatchLevels.UseLevels 1, 1, 1, 1, 3, 0       ' every rule on, the learned rules off

    TestAssert.AssertEqual MatchNone, Run("item235 - error 1", "vendor.example", id, rule)
    TestAssert.AssertEqual 0, pIndex.LookupCount, "no lookup for the rule that is off"

    MatchLevels.UseLevels 1, 1, 1, 1, 3, 1
    TestAssert.AssertEqual MatchFound, Run("item235 - error 1", "vendor.example", id, rule)
    Teardown
End Sub

Public Sub Test_Learned_DefaultLevelIsLogOnly()
    MatchLevels.ResetState
    TestAssert.AssertEqual LevelLogOnly, MatchLevels.LevelFor(RuleLearned)
End Sub

' --- the levels apply: proposal in the confirmation level ------------------------------------

Public Sub Test_Learned_ProposalAtLevel2()
    Dim mail As FakeMail
    Dim store As FakeTicketStore

    Setup
    TicketAudit.ResetState
    MatchProposals.Clear
    LearnedRules.UseJson RULES_JSON
    pIndex.AddTicket 40001, "Item235 - Error bei Verbuchung", True
    MatchLevels.UseLevels 0, 0, 0, 0, 3, 2
    Set store = New FakeTicketStore

    Set mail = New FakeMail
    mail.SenderName = "Some Customer"
    mail.SenderDomain = "vendor.example"
    mail.Subject = "item235 - error 4711"
    mail.MessageKey = "ab12ab12"

    TestAssert.AssertEqual MatchFound, TicketService.MatchForeignMail(mail, pIndex, store, True)
    TestAssert.AssertEqual 1, MatchProposals.Count()
    TestAssert.AssertEqual RuleLearned, MatchProposals.Item(1).Rule
    TestAssert.AssertEqual 40001, MatchProposals.Item(1).TicketId
    TestAssert.AssertEqual "Gelernte Regel", MatchPromptText.RuleLabel(RuleLearned)
    MatchProposals.Clear
    Teardown
End Sub

' --- the index and the report ------------------------------------------------------------------

Public Sub Test_Learned_IndexFindsOpenTicketsByStem()
    Dim fake As FakeNinjaTransport
    Dim impl As NinjaTicketIndex
    Dim idx As ITicketIndex
    Dim path As String

    Set fake = New FakeNinjaTransport
    path = "ticketing/trigger/board/2/run"
    fake.AddPostResponse path, NinjaTicketLogic.BuildBoardRunBody(2, 0, "status", "not_equals", "6000"), 200, Fixtures.BoardRunPage1()
    fake.AddPostResponse path, NinjaTicketLogic.BuildBoardRunBody(2, 2, "status", "not_equals", "6000"), 200, Fixtures.BoardRunPage2()
    Set impl = New NinjaTicketIndex
    impl.UseTransport fake
    impl.PageSize = 2
    impl.TestNow = DateSerial(2026, 10, 9)
    Set idx = impl

    TestAssert.AssertEqual 1, idx.FindOpenByStem("example printer").Count
    TestAssert.AssertEqual 301, idx.FindOpenByStem("example printer")(1).Id
    TestAssert.AssertEqual 3, idx.FindOpenByStem("example").Count, "all three tickets start with the word example"
    TestAssert.AssertEqual 0, idx.FindOpenByStem("nothing like it").Count
    TestAssert.AssertEqual 0, idx.FindOpenByStem("").Count
    TestAssert.AssertEqual 2, fake.PostCount, "one load of the two pages serves every lookup"
End Sub

Public Sub Test_Learned_ReportProposesRulesForTheRulesFile()
    Dim lines As String
    Dim i As Long
    Dim text As String

    For i = 1 To 5
        lines = lines & AuditLogic.BuildMatchEvent(DateSerial(2026, 9, i), "user_confirmed", "review", "found", "Subject", 20108 + i, "moved", _
            "vendor.example", "item235 - error " & i, "x", "", "", "v" & i, "m" & i, False) & vbLf
    Next i
    text = AuditReport.BuildReport(AuditReport.ParseLines(lines), DateSerial(2026, 10, 9))

    TestAssert.AssertTrue InStr(text, "```json") > 0, text
    TestAssert.AssertTrue InStr(text, "{""id"":""vendor-example-item-error"",""domain"":""vendor.example"",""stem"":""item# - error #"",""enabled"":false,") > 0, text
    TestAssert.AssertTrue InStr(text, "| Learned | 1 |") > 0, "the rules table has a row for the learned rules"
End Sub
