Attribute VB_Name = "ReviewFormTests"
' ReviewFormTests v0.1.0
'
' Tests of the review form ReviewForm (dialog variant B): the form is created but never shown, the
' ticking is done through the list box like a user would. forms\ReviewForm.frm must be part of the project.

Option Explicit

Private Function NewProposal(ByVal subject As String, ByVal ticketId As Long, ByVal summary As String) As MatchProposal
    Dim mail As FakeMail
    Dim p As MatchProposal

    Set mail = New FakeMail
    mail.Subject = subject
    mail.SenderDomain = "example.com"
    Set p = New MatchProposal
    Set p.Mail = mail
    p.TicketId = ticketId
    p.TicketSummary = summary
    p.Rule = RuleSubject
    Set NewProposal = p
End Function

Private Function ThreeProposals() As Collection
    Dim list As Collection

    Set list = New Collection
    list.Add NewProposal("Mail A", 101, "Ticket A")
    list.Add NewProposal("Mail B", 102, "Ticket B")
    list.Add NewProposal("Mail C", 103, "Ticket C")
    Set ThreeProposals = list
End Function

Private Function NewForm(ByVal proposals As Collection) As Object
    Dim f As Object

    Set f = VBA.UserForms.Add("ReviewForm")
    f.Setup proposals
    Set NewForm = f
End Function

Private Sub Tick(ByVal f As Object, ByVal listIndex As Long)
    f.Controls("lstProposals").Selected(listIndex) = True
End Sub

Private Function Joined(ByVal decisions As Collection) As String
    Dim v As Variant

    For Each v In decisions
        If Len(Joined) > 0 Then Joined = Joined & ","
        Joined = Joined & CStr(v)
    Next v
End Function

Public Sub Test_Form_ListsTheProposals()
    Dim f As Object
    Dim list As Object

    Set f = NewForm(ThreeProposals())
    Set list = f.Controls("lstProposals")

    TestAssert.AssertEqual 3, list.ListCount
    TestAssert.AssertTrue InStr(list.List(0), "#101 Ticket A") > 0, list.List(0)
    TestAssert.AssertTrue InStr(list.List(2), "Mail C") > 0, list.List(2)
    TestAssert.AssertEqual MatchPromptText.DialogTitle(), f.Caption
    TestAssert.AssertEqual MatchPromptText.FormAcceptCaption(), f.Controls("btnAccept").Caption
    TestAssert.AssertEqual 3, f.OpenCount()
    TestAssert.AssertEqual "0,0,0", Joined(f.Result()), "nothing decided yet"
    Unload f
End Sub

Public Sub Test_Form_AcceptAll()
    Dim f As Object

    Set f = NewForm(ThreeProposals())
    f.SelectAll
    TestAssert.AssertEqual 3, f.SelectedCount()

    f.ApplySelected DecisionAccept
    TestAssert.AssertEqual "1,1,1", Joined(f.Result())
    TestAssert.AssertEqual 0, f.OpenCount()
    TestAssert.AssertEqual 0, f.Controls("lstProposals").ListCount, "decided rows leave the list"
    Unload f
End Sub

Public Sub Test_Form_PartialDecisionsKeepTheOrder()
    Dim f As Object

    Set f = NewForm(ThreeProposals())
    Tick f, 1                                  ' Mail B
    f.ApplySelected DecisionReject
    TestAssert.AssertEqual "0,2,0", Joined(f.Result())
    TestAssert.AssertEqual 2, f.Controls("lstProposals").ListCount
    TestAssert.AssertTrue InStr(f.Controls("lstProposals").List(1), "Mail C") > 0, "row 2 is now Mail C"

    Tick f, 1                                  ' Mail C, the second row of the shortened list
    f.ApplySelected DecisionAccept
    TestAssert.AssertEqual "0,2,1", Joined(f.Result()), "the right proposal gets the decision"
    TestAssert.AssertEqual 1, f.OpenCount()

    f.SelectAll
    f.ApplySelected DecisionAccept
    TestAssert.AssertEqual "1,2,1", Joined(f.Result())
    Unload f
End Sub

Public Sub Test_Form_NothingTickedChangesNothing()
    Dim f As Object

    Set f = NewForm(ThreeProposals())
    f.ApplySelected DecisionAccept
    TestAssert.AssertEqual "0,0,0", Joined(f.Result())
    TestAssert.AssertEqual 3, f.Controls("lstProposals").ListCount
    Unload f
End Sub

Public Sub Test_Form_NoProposals()
    Dim f As Object

    Set f = NewForm(New Collection)
    TestAssert.AssertEqual 0, f.Controls("lstProposals").ListCount
    TestAssert.AssertEqual 0, f.Result().Count
    TestAssert.AssertEqual 0, f.OpenCount()
    Unload f
End Sub

Public Sub Test_Form_ResultMatchesWhatTheReviewExpects()
    ' MatchReview reads the result as a Collection of decisions in the order of the proposals
    Dim f As Object
    Dim result As Collection

    Set f = NewForm(ThreeProposals())
    Tick f, 0
    f.ApplySelected DecisionAccept
    Set result = f.Result()
    TestAssert.AssertEqual 3, result.Count
    TestAssert.AssertEqual CLng(DecisionAccept), CLng(result(1))
    TestAssert.AssertEqual CLng(DecisionSkip), CLng(result(2))
    Unload f
End Sub

Public Sub Test_PromptText_ListLine()
    Dim line As String

    line = MatchPromptText.ListLine(NewProposal("AW: Example total mismatch", 20106, "Example order total issue"), 2, 5)
    TestAssert.AssertTrue Left$(line, 4) = "2/5 ", line
    TestAssert.AssertTrue InStr(line, "#20106 Example order total issue") > 0
    TestAssert.AssertTrue InStr(line, "Von: example.com") > 0
    TestAssert.AssertTrue InStr(line, "Gleicher Betreff") > 0
    TestAssert.AssertTrue InStr(line, "Mail: AW: Example total mismatch") > 0

    line = MatchPromptText.ListLine(NewProposal(String$(200, "x"), 5, ""), 1, 1)
    TestAssert.AssertTrue InStr(line, String$(70, "x") & "...") > 0, "long subject is cut"
    TestAssert.AssertTrue InStr(line, "#5  |  Von:") > 0, "no summary: only the number"
End Sub
