VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeTicketIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FakeTicketIndex v0.1.0
'
' Test double for ITicketIndex: an in-memory ticket list. Lookups are deliberately loose
' (all open tickets are subject candidates) because the matcher re-checks every candidate.
Implements ITicketIndex
Option Explicit

Private pTickets As Collection
Private pConversations As Collection

Public LookupCount As Long

Private Sub Class_Initialize()
    Set pTickets = New Collection
    Set pConversations = New Collection
End Sub

Public Sub AddTicket(ByVal ticketId As Long, ByVal summary As String, ByVal isOpen As Boolean)
    Dim t As MatchTicket

    Set t = New MatchTicket
    t.Id = ticketId
    t.Summary = summary
    t.IsOpen = isOpen
    pTickets.Add t
End Sub

Public Sub AssignConversation(ByVal conversationId As String, ByVal ticketId As Long)
    On Error Resume Next
    pConversations.Remove "c" & conversationId
    On Error GoTo 0
    pConversations.Add ticketId, "c" & conversationId
End Sub

Private Sub ITicketIndex_RememberConversation(ByVal conversationId As String, ByVal ticketId As Long)
    AssignConversation conversationId, ticketId
End Sub

Private Function ITicketIndex_GetTicket(ByVal ticketId As Long) As MatchTicket
    Dim t As MatchTicket

    LookupCount = LookupCount + 1
    For Each t In pTickets
        If t.Id = ticketId Then
            Set ITicketIndex_GetTicket = t
            Exit Function
        End If
    Next t
End Function

Private Function ITicketIndex_FindByExternalNumber(ByVal externalNumber As String) As Collection
    Dim result As Collection
    Dim t As MatchTicket

    LookupCount = LookupCount + 1
    Set result = New Collection
    For Each t In pTickets
        If InStr(t.Summary, externalNumber) > 0 Then result.Add t
    Next t
    Set ITicketIndex_FindByExternalNumber = result
End Function

Private Function ITicketIndex_FindOpenBySubject(ByVal normalizedSubject As String) As Collection
    Dim result As Collection
    Dim t As MatchTicket

    LookupCount = LookupCount + 1
    Set result = New Collection
    For Each t In pTickets
        If t.IsOpen Then result.Add t
    Next t
    Set ITicketIndex_FindOpenBySubject = result
End Function

Private Function ITicketIndex_FindByConversation(ByVal conversationId As String) As Long
    LookupCount = LookupCount + 1
    On Error Resume Next
    ITicketIndex_FindByConversation = CLng(pConversations("c" & conversationId))
    On Error GoTo 0
End Function

Private Function ITicketIndex_FindOpenByStem(ByVal stemPrefix As String) As Collection
    Dim result As Collection
    Dim t As MatchTicket

    LookupCount = LookupCount + 1
    Set result = New Collection
    For Each t In pTickets
        If t.IsOpen Then result.Add t
    Next t
    Set ITicketIndex_FindOpenByStem = result
End Function
