Attribute VB_Name = "MatchLevelsTests"
' MatchLevelsTests v0.1.0
'
' Tests of the safety levels (MatchLevels), their effect on the matcher (level 0 = rule off) and on
' TicketService.MatchForeignMail (log only / propose / move) and of the proposal list. No Outlook,
' no network. Ticket 20101 is open, 20102 is closed, 20001 carries the vendor number 900000101,
' 30001 is found by its subject.

Option Explicit

Private pSink As FakeAuditSink
Private pStore As FakeTicketStore
Private pIndex As FakeTicketIndex
Private pTransport As FakeNinjaTransport

Private Sub Setup()
    TicketLog.DisableFileLogging
    TicketLog.Clear
    NinjaApi.ResetState
    TicketAudit.ResetState
    MatchLevels.ResetState
    MatchProposals.Clear

    Set pSink = New FakeAuditSink
    TicketAudit.UseSink pSink

    Set pStore = New FakeTicketStore
    Set pTransport = New FakeNinjaTransport
    NinjaApi.UseTransport pTransport
    pTransport.AddResponse "ticketing/ticket/20101", 200, Fixtures.TicketOpen()
    pTransport.AddResponse "ticketing/ticket/20102", 200, Fixtures.TicketClosed()

    Set pIndex = New FakeTicketIndex
    pIndex.AddTicket 20101, "Example open ticket", True
    pIndex.AddTicket 20102, "Example closed ticket", False
    pIndex.AddTicket 20001, "CSM | [TicketNr. 900000101] - Vendor problem", True
    pIndex.AddTicket 30001, "Example order total differs from expected value", True
End Sub

Private Sub Teardown()
    MatchLevels.ResetState
    MatchProposals.Clear
    TicketAudit.ResetState
    NinjaApi.ResetState
    Set pSink = Nothing
    Set pStore = Nothing
    Set pIndex = Nothing
    Set pTransport = Nothing
End Sub

Private Function Foreign(ByVal subject As String, Optional ByVal conversationId As String = "", Optional ByVal messageKey As String = "") As FakeMail
    Dim mail As FakeMail

    Set mail = New FakeMail
    mail.SenderName = "Some Customer"
    mail.SenderDomain = "example.com"
    mail.Subject = subject
    mail.ConversationId = conversationId
    mail.MessageKey = messageKey
    Set Foreign = mail
End Function

Private Function Match(ByVal mail As FakeMail, Optional ByVal allowActions As Boolean = True) As MatchOutcome
    Match = TicketService.MatchForeignMail(mail, pIndex, pStore, allowActions)
End Function

Private Function RunMatcher(ByVal subject As String, ByVal conversationId As String, ByRef ticketId As Long, ByRef rule As MatchRule) As MatchOutcome
    Dim idx As ITicketIndex

    Set idx = pIndex
    RunMatcher = TicketMatcher.MatchMail(subject, "", conversationId, "example.com", False, idx, ticketId, rule)
End Function

' --- levels ---------------------------------------------------------------------

Public Sub Test_Levels_DefaultsAreLogOnly()
    MatchLevels.ResetState
    TestAssert.AssertEqual LevelLogOnly, MatchLevels.LevelFor(RuleNinjaId)
    TestAssert.AssertEqual LevelLogOnly, MatchLevels.LevelFor(RuleExternalNumber)
    TestAssert.AssertEqual LevelLogOnly, MatchLevels.LevelFor(RuleConversation)
    TestAssert.AssertEqual LevelLogOnly, MatchLevels.LevelFor(RuleSubject)
    TestAssert.AssertTrue MatchLevels.AnyRuleActive()
End Sub

Public Sub Test_Levels_PerRuleAndCap()
    MatchLevels.UseLevels 3, 2, 1, 0
    TestAssert.AssertEqual LevelAuto, MatchLevels.LevelFor(RuleNinjaId)
    TestAssert.AssertEqual LevelConfirm, MatchLevels.LevelFor(RuleExternalNumber)
    TestAssert.AssertEqual LevelLogOnly, MatchLevels.LevelFor(RuleConversation)
    TestAssert.AssertEqual LevelOff, MatchLevels.LevelFor(RuleSubject)

    MatchLevels.UseLevels 3, 3, 3, 3, 2
    TestAssert.AssertEqual LevelConfirm, MatchLevels.LevelFor(RuleNinjaId), "cap 2 lowers level 3"
    MatchLevels.UseLevels 1, 1, 1, 1, 2
    TestAssert.AssertEqual LevelLogOnly, MatchLevels.LevelFor(RuleSubject), "the cap never raises a level"

    MatchLevels.UseLevels 3, 3, 3, 3, 0
    TestAssert.AssertFalse MatchLevels.AnyRuleActive(), "cap 0 switches everything off"
    MatchLevels.ResetState
End Sub

Public Sub Test_Levels_OutOfRangeValuesAreClamped()
    MatchLevels.UseLevels 7, -1, 2, 9, 3
    TestAssert.AssertEqual LevelAuto, MatchLevels.LevelFor(RuleNinjaId)
    TestAssert.AssertEqual LevelOff, MatchLevels.LevelFor(RuleExternalNumber)
    TestAssert.AssertEqual LevelAuto, MatchLevels.LevelFor(RuleSubject)
    TestAssert.AssertEqual LevelOff, MatchLevels.LevelFor(RuleNone)
    MatchLevels.ResetState
End Sub

Public Sub Test_Levels_Names()
    TestAssert.AssertEqual "off", MatchLevels.LevelName(LevelOff)
    TestAssert.AssertEqual "log_only", MatchLevels.LevelName(LevelLogOnly)
    TestAssert.AssertEqual "confirm", MatchLevels.LevelName(LevelConfirm)
    TestAssert.AssertEqual "auto", MatchLevels.LevelName(LevelAuto)
End Sub

' --- level 0: the rule is not used -----------------------------------------------

Public Sub Test_Off_NinjaIdRuleIsSkipped()
    Dim id As Long, rule As MatchRule

    Setup
    MatchLevels.UseLevels 0, 1, 1, 1
    TestAssert.AssertEqual MatchNone, RunMatcher("Ticket 20101", "", id, rule)
    TestAssert.AssertEqual 0, pIndex.LookupCount, "no lookup for a switched-off rule"

    MatchLevels.UseLevels 1, 1, 1, 1
    TestAssert.AssertEqual MatchFound, RunMatcher("Ticket 20101", "", id, rule)
    TestAssert.AssertEqual 20101, id
    Teardown
End Sub

Public Sub Test_Off_ExternalNumberRuleDoesNotStopTheOthers()
    Dim id As Long, rule As MatchRule

    Setup
    pIndex.AssignConversation "c1", 20101

    MatchLevels.UseLevels 1, 1, 1, 1
    TestAssert.AssertEqual MatchNone, RunMatcher("[TicketNr. 399999999] unknown vendor ticket", "c1", id, rule), "rule on: unknown number stops"

    MatchLevels.UseLevels 1, 0, 1, 1
    TestAssert.AssertEqual MatchFound, RunMatcher("[TicketNr. 399999999] unknown vendor ticket", "c1", id, rule), "rule off: conversation decides"
    TestAssert.AssertEqual RuleConversation, rule
    TestAssert.AssertEqual 20101, id
    Teardown
End Sub

Public Sub Test_Off_ConversationAndSubjectRules()
    Dim id As Long, rule As MatchRule

    Setup
    pIndex.AssignConversation "c1", 20101

    MatchLevels.UseLevels 1, 1, 0, 1
    TestAssert.AssertEqual MatchNone, RunMatcher("Kurz", "c1", id, rule), "conversation rule off"

    MatchLevels.UseLevels 1, 1, 1, 0
    TestAssert.AssertEqual MatchNone, RunMatcher("AW: Example order total differs from expected value", "", id, rule), "subject rule off"

    MatchLevels.UseLevels 1, 1, 1, 1
    TestAssert.AssertEqual MatchFound, RunMatcher("AW: Example order total differs from expected value", "", id, rule)
    TestAssert.AssertEqual 30001, id
    Teardown
End Sub

Public Sub Test_Off_AllRulesOffDoesNothingAtAll()
    Setup
    MatchLevels.UseLevels 0, 0, 0, 0

    TestAssert.AssertEqual MatchNone, Match(Foreign("Ticket 20101"))
    TestAssert.AssertEqual 0, pIndex.LookupCount
    TestAssert.AssertEqual "", TicketLog.Text(), "not even logged"
    TestAssert.AssertEqual 0, pSink.Count, "no audit line"
    Teardown
End Sub

' --- what the level does -----------------------------------------------------------

Public Sub Test_Level1_OnlyLogs()
    Dim mail As FakeMail

    Setup
    MatchLevels.UseAll LevelLogOnly
    Set mail = Foreign("Ticket 20101")

    TestAssert.AssertEqual MatchFound, Match(mail)
    TestAssert.AssertTrue mail.MovedTo Is Nothing
    TestAssert.AssertEqual 0, MatchProposals.Count()
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""log_only""") > 0, pSink.LastLine()
    Teardown
End Sub

Public Sub Test_Level2_QueuesAProposal()
    Dim mail As FakeMail
    Dim p As MatchProposal

    Setup
    MatchLevels.UseAll LevelConfirm
    Set mail = Foreign("Ticket 20101", "c1", "k1k1k1k1")

    TestAssert.AssertEqual MatchFound, Match(mail)
    TestAssert.AssertTrue mail.MovedTo Is Nothing, "nothing is moved before the user confirms"
    TestAssert.AssertEqual 0, pStore.ActiveCount, "no folder yet"
    TestAssert.AssertEqual 1, MatchProposals.Count()

    Set p = MatchProposals.Item(1)
    TestAssert.AssertEqual 20101, p.TicketId
    TestAssert.AssertEqual "Example open ticket", p.TicketSummary
    TestAssert.AssertEqual RuleNinjaId, p.Rule
    TestAssert.AssertEqual "k1k1k1k1", p.Mail.MessageKey
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""proposed""") > 0, pSink.LastLine()
    TestAssert.AssertTrue InStr(TicketLog.Text(), "proposed") > 0
    Teardown
End Sub

Public Sub Test_Level3_MovesWithoutAsking()
    Dim mail As FakeMail

    Setup
    MatchLevels.UseAll LevelAuto
    Set mail = Foreign("Ticket 20101")

    TestAssert.AssertEqual MatchFound, Match(mail)
    TestAssert.AssertFalse mail.MovedTo Is Nothing, "moved"
    TestAssert.AssertEqual 0, MatchProposals.Count()
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""moved""") > 0, pSink.LastLine()
    Teardown
End Sub

Public Sub Test_DryRunNeverActs_WhateverTheLevel()
    Dim mail As FakeMail

    Setup
    MatchLevels.UseAll LevelAuto
    Set mail = Foreign("Ticket 20101")
    TestAssert.AssertEqual MatchFound, Match(mail, False)
    TestAssert.AssertTrue mail.MovedTo Is Nothing

    MatchLevels.UseAll LevelConfirm
    TestAssert.AssertEqual MatchFound, Match(Foreign("Ticket 20101"), False)
    TestAssert.AssertEqual 0, MatchProposals.Count(), "a dry run adds no proposals"
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""log_only""") > 0
    Teardown
End Sub

Public Sub Test_Level_IsPerRule()
    Dim byId As FakeMail
    Dim bySubject As FakeMail

    Setup
    MatchLevels.UseLevels 3, 1, 1, 2
    Set byId = Foreign("Ticket 20101", "c1")
    Set bySubject = Foreign("AW: Example order total differs from expected value", "c2")

    TestAssert.AssertEqual MatchFound, Match(byId)
    TestAssert.AssertFalse byId.MovedTo Is Nothing, "Ninja id rule at level 3: moved"
    TestAssert.AssertEqual MatchFound, Match(bySubject)
    TestAssert.AssertTrue bySubject.MovedTo Is Nothing
    TestAssert.AssertEqual 1, MatchProposals.Count(), "subject rule at level 2: proposed"
    TestAssert.AssertEqual 30001, MatchProposals.Item(1).TicketId
    TestAssert.AssertEqual RuleSubject, MatchProposals.Item(1).Rule
    Teardown
End Sub

Public Sub Test_Cap_LowersTheLevel()
    Dim mail As FakeMail

    Setup
    MatchLevels.UseLevels 3, 3, 3, 3, 2
    Set mail = Foreign("Ticket 20101")

    TestAssert.AssertEqual MatchFound, Match(mail)
    TestAssert.AssertTrue mail.MovedTo Is Nothing, "cap 2: only proposed"
    TestAssert.AssertEqual 1, MatchProposals.Count()
    Teardown
End Sub

' --- protection rules hold at every level ---------------------------------------

Public Sub Test_ClosedTicketIsNeverProposedOrMoved()
    Dim mail As FakeMail

    Setup
    MatchLevels.UseAll LevelConfirm
    Set mail = Foreign("Ticket 20102")
    TestAssert.AssertEqual MatchFound, Match(mail)
    TestAssert.AssertEqual 0, MatchProposals.Count(), "closed ticket: no proposal"
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""stays""") > 0, pSink.LastLine()

    MatchLevels.UseAll LevelAuto
    Set mail = Foreign("Ticket 20102")
    TestAssert.AssertEqual MatchFound, Match(mail)
    TestAssert.AssertTrue mail.MovedTo Is Nothing, "closed ticket: not moved"
    Teardown
End Sub

Public Sub Test_AmbiguousIsNeverProposedOrMoved()
    Dim mail As FakeMail

    Setup
    MatchLevels.UseAll LevelAuto
    Set mail = Foreign("Ticket 20101 und Ticket 20102")

    TestAssert.AssertEqual MatchAmbiguous, Match(mail)
    TestAssert.AssertTrue mail.MovedTo Is Nothing
    TestAssert.AssertEqual 0, MatchProposals.Count()
    Teardown
End Sub

Public Sub Test_TicketSenderIsNotProposed()
    Dim mail As FakeMail

    Setup
    MatchLevels.UseAll LevelConfirm
    Set mail = Foreign("Ticket 20101")
    mail.SenderName = TicketConfig.AllowedSender()

    TestAssert.AssertEqual MatchNone, Match(mail)
    TestAssert.AssertEqual 0, MatchProposals.Count()
    Teardown
End Sub

' --- the proposal list ----------------------------------------------------------------

Public Sub Test_Proposals_SameMailIsNotAddedTwice()
    Setup
    MatchLevels.UseAll LevelConfirm

    Match Foreign("Ticket 20101", "c1", "aa11aa11")
    Match Foreign("Ticket 20101", "c1", "aa11aa11")
    TestAssert.AssertEqual 1, MatchProposals.Count(), "same message key"
    Match Foreign("Ticket 20101", "c2", "bb22bb22")
    TestAssert.AssertEqual 2, MatchProposals.Count(), "another mail"
    TestAssert.AssertEqual "aa11aa11", MatchProposals.Item(1).Mail.MessageKey, "oldest first"
    Teardown
End Sub

Public Sub Test_Proposals_RemoveAndClear()
    Setup
    MatchLevels.UseAll LevelConfirm
    Match Foreign("Ticket 20101", "c1", "aa11aa11")
    Match Foreign("Ticket 20101", "c2", "bb22bb22")

    MatchProposals.Remove 1
    TestAssert.AssertEqual 1, MatchProposals.Count()
    TestAssert.AssertEqual "bb22bb22", MatchProposals.Item(1).Mail.MessageKey
    MatchProposals.Clear
    TestAssert.AssertEqual 0, MatchProposals.Count()
    Teardown
End Sub

Public Sub Test_Levels_AnyRuleAtLevel()
    MatchLevels.UseLevels 1, 1, 0, 2
    TestAssert.AssertTrue MatchLevels.AnyRuleAtLevel(LevelConfirm)
    TestAssert.AssertTrue MatchLevels.AnyRuleAtLevel(LevelOff)
    TestAssert.AssertFalse MatchLevels.AnyRuleAtLevel(LevelAuto)

    MatchLevels.UseLevels 3, 3, 3, 3, 1
    TestAssert.AssertFalse MatchLevels.AnyRuleAtLevel(LevelConfirm), "the cap lowers every level to 1"
    TestAssert.AssertTrue MatchLevels.AnyRuleAtLevel(LevelLogOnly)
    MatchLevels.ResetState
End Sub
