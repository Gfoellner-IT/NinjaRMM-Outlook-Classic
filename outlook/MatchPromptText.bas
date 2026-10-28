Attribute VB_Name = "MatchPromptText"
' MatchPromptText v0.1.0
'
' The texts of the confirmation dialogs and the mapping of the button answers to decisions. Pure
' functions without Outlook types, so the message box dialog stays a thin shell and the texts
' can be unit tested. Dialog texts are German (the users are German speaking); umlauts are built
' with ChrW because source files are ASCII-only.

Option Explicit

Private Const MAX_SUBJECT_CHARS As Long = 120

Public Function DialogTitle() As String
    DialogTitle = "Ticketzuordnung"
End Function

' The hint question.
Public Function HintText(ByVal pendingCount As Long) As String
    If pendingCount = 1 Then
        HintText = "Es gibt 1 neuen Vorschlag f" & ChrW$(252) & "r die Ticketzuordnung."
    Else
        HintText = "Es gibt " & CStr(pendingCount) & " neue Vorschl" & ChrW$(228) & "ge f" & ChrW$(252) & "r die Ticketzuordnung."
    End If
    HintText = HintText & vbCrLf & vbCrLf & "Diese jetzt pr" & ChrW$(252) & "fen?"
End Function

' The question for one proposal; position is 1-based.
Public Function ProposalText(ByVal proposal As MatchProposal, ByVal position As Long, ByVal total As Long) As String
    Dim subjectText As String
    Dim ticketText As String

    subjectText = Replace(Replace(Replace(proposal.Mail.Subject, vbCrLf, " "), vbCr, " "), vbLf, " ")
    If Len(subjectText) > MAX_SUBJECT_CHARS Then subjectText = Left$(subjectText, MAX_SUBJECT_CHARS) & "..."

    ticketText = "#" & CStr(proposal.TicketId)
    If Len(proposal.TicketSummary) > 0 Then ticketText = ticketText & " " & proposal.TicketSummary

    ProposalText = "Vorschlag " & CStr(position) & " von " & CStr(total) & vbCrLf & vbCrLf & _
                   "Mail:    " & subjectText & vbCrLf & _
                   "Von:     " & proposal.Mail.SenderDomain & vbCrLf & _
                   "Ticket:  " & ticketText & vbCrLf & _
                   "Regel:   " & RuleLabel(proposal.Rule) & vbCrLf & vbCrLf & _
                   "In den Ticketordner verschieben?" & vbCrLf & vbCrLf & _
                   "Ja = " & ChrW$(252) & "bernehmen" & vbCrLf & _
                   "Nein = ablehnen (wird nicht wieder vorgeschlagen)" & vbCrLf & _
                   "Abbrechen = Pr" & ChrW$(252) & "fung beenden (die restlichen Vorschl" & ChrW$(228) & "ge bleiben)"
End Function

Public Function NoProposalsText() As String
    NoProposalsText = "Es gibt keine offenen Vorschl" & ChrW$(228) & "ge."
End Function

Public Function NoConfirmLevelText() As String
    NoConfirmLevelText = "Keine Regel steht auf Stufe 2 (Best" & ChrW$(228) & "tigen), deshalb gibt es keine Vorschl" & ChrW$(228) & "ge." & vbCrLf & vbCrLf & _
                         "Die Stufen stehen in TicketConfig (MATCH_LEVEL_*)."
End Function

' German name of a matching rule for the dialog.
Public Function RuleLabel(ByVal rule As MatchRule) As String
    Select Case rule
        Case RuleNinjaId:         RuleLabel = "Ticketnummer in der Mail"
        Case RuleExternalNumber:  RuleLabel = "Nummer des Herstellertickets"
        Case RuleConversation:    RuleLabel = "Gleiche Konversation"
        Case RuleSubject:         RuleLabel = "Gleicher Betreff"
        Case RuleLearned:         RuleLabel = "Gelernte Regel"
        Case Else:                RuleLabel = "Unbekannt"
    End Select
End Function

' Decision for a message box answer (vbYes = 6, vbNo = 7, anything else = cancel).
' stopReview is set when the user cancelled: the proposals not shown yet stay.
Public Function DecisionForAnswer(ByVal answer As Long, ByRef stopReview As Boolean) As MatchDecision
    stopReview = False
    Select Case answer
        Case 6      ' vbYes
            DecisionForAnswer = DecisionAccept
        Case 7      ' vbNo
            DecisionForAnswer = DecisionReject
        Case Else
            stopReview = True
            DecisionForAnswer = DecisionSkip
    End Select
End Function

' ---------------------------------------------------------------------------
'  Review form (variant B)
' ---------------------------------------------------------------------------

' One row of the list in the review form: "#20101 Titel | Von: domain | Regel | Mail: Betreff".
Public Function ListLine(ByVal proposal As MatchProposal, ByVal position As Long, ByVal total As Long) As String
    Dim subjectText As String
    Dim ticketText As String

    subjectText = Replace(Replace(Replace(proposal.Mail.Subject, vbCrLf, " "), vbCr, " "), vbLf, " ")
    If Len(subjectText) > 70 Then subjectText = Left$(subjectText, 70) & "..."

    ticketText = "#" & CStr(proposal.TicketId)
    If Len(proposal.TicketSummary) > 0 Then ticketText = ticketText & " " & Left$(proposal.TicketSummary, 50)

    ListLine = CStr(position) & "/" & CStr(total) & "  " & ticketText & "  |  Von: " & proposal.Mail.SenderDomain & _
               "  |  " & RuleLabel(proposal.Rule) & "  |  Mail: " & subjectText
End Function

Public Function FormInfoText() As String
    FormInfoText = "Vorschl" & ChrW$(228) & "ge f" & ChrW$(252) & "r die Ticketzuordnung. Kreuze Mails an und w" & ChrW$(228) & "hle eine Aktion. " & _
                   "Nicht entschiedene Vorschl" & ChrW$(228) & "ge bleiben beim Schlie" & ChrW$(223) & "en erhalten."
End Function

Public Function FormAcceptCaption() As String
    FormAcceptCaption = "Ausgew" & ChrW$(228) & "hlte " & ChrW$(252) & "bernehmen"
End Function

Public Function FormRejectCaption() As String
    FormRejectCaption = "Ausgew" & ChrW$(228) & "hlte ablehnen"
End Function

Public Function FormSelectAllCaption() As String
    FormSelectAllCaption = "Alle ausw" & ChrW$(228) & "hlen"
End Function

Public Function FormCloseCaption() As String
    FormCloseCaption = "Schlie" & ChrW$(223) & "en"
End Function
