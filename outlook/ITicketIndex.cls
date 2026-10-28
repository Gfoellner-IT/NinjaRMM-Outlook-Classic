VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITicketIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ITicketIndex v0.1.0
'
' Lookup seam of the TicketMatcher: how candidate tickets are found. The matcher decides,
' the index only fetches candidates (and may cache). Unit tests use FakeTicketIndex.
Option Explicit

' The ticket with that NinjaOne id, Nothing if it does not exist.
Public Function GetTicket(ByVal ticketId As Long) As MatchTicket
End Function

' Collection of MatchTicket whose summary contains the external (vendor) ticket number.
' Open and closed tickets.
Public Function FindByExternalNumber(ByVal externalNumber As String) As Collection
End Function

' Collection of MatchTicket (open tickets only) whose normalized summary may equal
' normalizedSubject. The matcher re-checks every candidate.
Public Function FindOpenBySubject(ByVal normalizedSubject As String) As Collection
End Function

' Ticket id already assigned to this mail conversation, 0 if none.
Public Function FindByConversation(ByVal conversationId As String) As Long
End Function

' Remembers that mails of this conversation belong to the ticket (rule 3 for later mails).
Public Sub RememberConversation(ByVal conversationId As String, ByVal ticketId As Long)
End Sub

' Collection of MatchTicket (open tickets only) whose title may start with the stem words of a learned
' rule. The matcher re-checks every candidate.
Public Function FindOpenByStem(ByVal stemPrefix As String) As Collection
End Function
