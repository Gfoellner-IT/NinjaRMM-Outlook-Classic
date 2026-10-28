Attribute VB_Name = "MatchPromptTextTests"
' MatchPromptTextTests v0.1.0
'
' Tests of the dialog texts and the mapping of message box answers (MatchPromptText). The message
' box dialog itself (MsgBoxMatchPrompt) is tested by hand in Outlook.

Option Explicit

Private Function NewProposal(ByVal subject As String, ByVal domain As String, ByVal ticketId As Long, _
                             ByVal summary As String, ByVal rule As MatchRule) As MatchProposal
    Dim mail As FakeMail
    Dim p As MatchProposal

    Set mail = New FakeMail
    mail.Subject = subject
    mail.SenderDomain = domain

    Set p = New MatchProposal
    Set p.Mail = mail
    p.TicketId = ticketId
    p.TicketSummary = summary
    p.Rule = rule
    Set NewProposal = p
End Function

Public Sub Test_PromptText_Hint()
    Dim one As String
    Dim many As String

    one = MatchPromptText.HintText(1)
    many = MatchPromptText.HintText(3)

    TestAssert.AssertTrue InStr(one, "1 neuen Vorschlag") > 0, one
    TestAssert.AssertTrue InStr(many, "3 neue Vorschl" & ChrW$(228) & "ge") > 0, many
    TestAssert.AssertTrue InStr(many, "jetzt pr" & ChrW$(252) & "fen?") > 0, "ends with the question"
End Sub

Public Sub Test_PromptText_Proposal()
    Dim text As String

    text = MatchPromptText.ProposalText(NewProposal("AW: Example total mismatch", "example.com", 20106, "Example order total issue", RuleSubject), 2, 5)

    TestAssert.AssertTrue InStr(text, "Vorschlag 2 von 5") > 0, text
    TestAssert.AssertTrue InStr(text, "AW: Example total mismatch") > 0
    TestAssert.AssertTrue InStr(text, "example.com") > 0
    TestAssert.AssertTrue InStr(text, "#20106 Example order total issue") > 0
    TestAssert.AssertTrue InStr(text, "Gleicher Betreff") > 0, "the rule in plain words"
    TestAssert.AssertTrue InStr(text, "Ja = " & ChrW$(252) & "bernehmen") > 0
    TestAssert.AssertTrue InStr(text, "Nein = ablehnen") > 0
    TestAssert.AssertTrue InStr(text, "Abbrechen") > 0
End Sub

Public Sub Test_PromptText_ProposalWithoutSummaryAndLongSubject()
    Dim text As String

    text = MatchPromptText.ProposalText(NewProposal(String$(300, "x") & vbCrLf & "tail", "vendor.example", 20103, "", RuleExternalNumber), 1, 1)
    TestAssert.AssertTrue InStr(text, "Ticket:  #20103" & vbCrLf) > 0, "no summary: only the number: " & text
    TestAssert.AssertEqual 0, InStr(text, "tail"), "long subject is cut"
    TestAssert.AssertTrue InStr(text, String$(120, "x") & "...") > 0
    TestAssert.AssertTrue InStr(text, "Nummer des Herstellertickets") > 0
End Sub

Public Sub Test_PromptText_RuleLabels()
    TestAssert.AssertEqual "Ticketnummer in der Mail", MatchPromptText.RuleLabel(RuleNinjaId)
    TestAssert.AssertEqual "Nummer des Herstellertickets", MatchPromptText.RuleLabel(RuleExternalNumber)
    TestAssert.AssertEqual "Gleiche Konversation", MatchPromptText.RuleLabel(RuleConversation)
    TestAssert.AssertEqual "Gleicher Betreff", MatchPromptText.RuleLabel(RuleSubject)
    TestAssert.AssertEqual "Unbekannt", MatchPromptText.RuleLabel(RuleNone)
End Sub

Public Sub Test_PromptText_AnswersToDecisions()
    Dim stopReview As Boolean

    TestAssert.AssertEqual DecisionAccept, MatchPromptText.DecisionForAnswer(vbYes, stopReview)
    TestAssert.AssertFalse stopReview
    TestAssert.AssertEqual DecisionReject, MatchPromptText.DecisionForAnswer(vbNo, stopReview)
    TestAssert.AssertFalse stopReview
    TestAssert.AssertEqual DecisionSkip, MatchPromptText.DecisionForAnswer(vbCancel, stopReview)
    TestAssert.AssertTrue stopReview, "cancel stops the review"
    TestAssert.AssertEqual DecisionSkip, MatchPromptText.DecisionForAnswer(0, stopReview)
    TestAssert.AssertTrue stopReview, "anything unknown is treated as cancel"
End Sub

Public Sub Test_PromptText_OtherTexts()
    TestAssert.AssertTrue Len(MatchPromptText.DialogTitle()) > 0
    TestAssert.AssertTrue InStr(MatchPromptText.NoProposalsText(), "keine offenen") > 0
    TestAssert.AssertTrue InStr(MatchPromptText.NoConfirmLevelText(), "Stufe 2") > 0
    TestAssert.AssertTrue InStr(MatchPromptText.NoConfirmLevelText(), "MATCH_LEVEL_") > 0
End Sub
