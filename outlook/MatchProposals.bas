Attribute VB_Name = "MatchProposals"
' MatchProposals v0.1.0
'
' The list of pending proposals (level 2), kept in memory for the running Outlook session.
' A proposal is dropped when it is decided, and a mail that is already in the list (same message
' key) is not added twice. No Outlook types.

Option Explicit

Private pList As Collection

Public Function Count() As Long
    If pList Is Nothing Then Exit Function
    Count = pList.Count
End Function

' Adds a proposal unless the same mail (non-empty message key) is already waiting.
' Returns True if it was added.
Public Function Add(ByVal mail As IIncomingMail, ByVal ticketId As Long, ByVal ticketSummary As String, ByVal rule As MatchRule) As Boolean
    Dim p As MatchProposal
    Dim key As String
    Dim i As Long

    If pList Is Nothing Then Set pList = New Collection

    key = mail.MessageKey
    If Len(key) > 0 Then
        For i = 1 To pList.Count
            If pList(i).Mail.MessageKey = key Then Exit Function
        Next i
    End If

    Set p = New MatchProposal
    Set p.Mail = mail
    p.TicketId = ticketId
    p.TicketSummary = ticketSummary
    p.Rule = rule
    p.CreatedAt = Now
    pList.Add p
    Add = True
End Function

' 1-based, oldest first.
Public Function Item(ByVal index As Long) As MatchProposal
    Set Item = pList(index)
End Function

Public Sub Remove(ByVal index As Long)
    pList.Remove index
End Sub

Public Sub Clear()
    Set pList = New Collection
End Sub

' Removes this proposal (by identity); nothing happens if it is not in the list.
Public Sub RemoveItem(ByVal proposal As MatchProposal)
    Dim i As Long

    If pList Is Nothing Then Exit Sub
    For i = 1 To pList.Count
        If pList(i) Is proposal Then
            pList.Remove i
            Exit Sub
        End If
    Next i
End Sub
