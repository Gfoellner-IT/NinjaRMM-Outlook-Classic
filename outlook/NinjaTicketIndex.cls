VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NinjaTicketIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' NinjaTicketIndex v0.1.0
'
' ITicketIndex on top of the read-only NinjaOne API, with an in-memory cache:
'
'   GetTicket             GET ticketing/ticket/{id}, one request, not cached
'   FindByExternalNumber  one board query "summary contains <number>" per number, cached
'   FindOpenBySubject     all non-closed tickets in pages (filter status <> 6000), grouped by
'                         normalized subject, reloaded when older than INDEX_CACHE_MINUTES
'   FindByConversation    conversation -> ticket table in memory (RememberConversation)
'
' A failed request never fills the cache and leaves an older cache untouched; the lookups then
' return "nothing found", so the mail stays in the inbox. LastError describes the problem for
' the log (never contains secrets). No Outlook types.
Implements ITicketIndex
Option Explicit

Private Const EXTERNAL_PAGE_SIZE As Long = 50
Private Const MAX_PAGES As Long = 40

Private pTransport As INinjaTransport
Private pLastError As String

Private pOpenBySubject As Object      ' normalized subject -> Collection of MatchTicket
Private pOpenByStem As Object         ' subject stem (4 words) -> Collection of MatchTicket
Private pOpenLoadedAt As Date
Private pExternal As Object           ' number -> Array(loadedAt, Collection)
Private pConversations As Object      ' conversation id -> ticket id

' Page size of the open-ticket query (tests use small values to exercise paging).
Public PageSize As Long
' Cache lifetime in minutes.
Public CacheMinutes As Long
' Test seam: fixed "now" (0 = real time).
Public TestNow As Date

Private Sub Class_Initialize()
    PageSize = TicketConfig.INDEX_PAGE_SIZE
    CacheMinutes = TicketConfig.INDEX_CACHE_MINUTES
    Set pExternal = CreateObject("Scripting.Dictionary")
    Set pConversations = CreateObject("Scripting.Dictionary")
End Sub

Public Sub UseTransport(ByVal transport As INinjaTransport)
    Set pTransport = transport
End Sub

Public Property Get LastError() As String
    LastError = pLastError
End Property

' Drops all cached lists (conversation assignments are kept).
Public Sub ClearCache()
    Set pOpenBySubject = Nothing
    Set pOpenByStem = Nothing
    pExternal.RemoveAll
End Sub

' Remembers that mails of this conversation belong to the ticket (rule 3).
Public Sub RememberConversation(ByVal conversationId As String, ByVal ticketId As Long)
    If Len(conversationId) = 0 Or ticketId <= 0 Then Exit Sub
    pConversations.Item(conversationId) = ticketId
End Sub

' ---------------------------------------------------------------------------
'  ITicketIndex
' ---------------------------------------------------------------------------

Private Function ITicketIndex_GetTicket(ByVal ticketId As Long) As MatchTicket
    Dim status As Long
    Dim body As String
    Dim errorText As String
    Dim state As NinjaTicketState
    Dim t As MatchTicket

    pLastError = ""
    If ticketId <= 0 Then Exit Function

    body = Transport().HttpGet("ticketing/ticket/" & CStr(ticketId), status, errorText)
    If status = 0 Then
        pLastError = "No response for ticket " & ticketId & ": " & errorText
        Exit Function
    End If

    state = NinjaTicketLogic.ParseTicketState(status, body)
    Select Case state
        Case NinjaTicketOpen, NinjaTicketClosed
            Set t = New MatchTicket
            t.Id = ticketId
            t.Summary = NinjaTicketLogic.ParseTicketSubject(body)
            t.IsOpen = (state = NinjaTicketOpen)
            Set ITicketIndex_GetTicket = t
        Case NinjaTicketUnavailable
            pLastError = "Unexpected answer for ticket " & ticketId & " (HTTP " & status & ")"
    End Select
End Function

Private Function ITicketIndex_FindByExternalNumber(ByVal externalNumber As String) As Collection
    Dim entry As Variant
    Dim result As Collection
    Dim page As Collection
    Dim cursor As Long
    Dim rowCount As Long
    Dim t As MatchTicket

    pLastError = ""
    Set ITicketIndex_FindByExternalNumber = New Collection
    If Len(externalNumber) = 0 Then Exit Function

    If pExternal.Exists(externalNumber) Then
        entry = pExternal.Item(externalNumber)
        If IsFresh(CDate(entry(0))) Then
            Set ITicketIndex_FindByExternalNumber = entry(1)
            Exit Function
        End If
    End If

    If Not QueryPage(NinjaTicketLogic.BuildBoardRunBody(EXTERNAL_PAGE_SIZE, 0, "summary", "contains", externalNumber), _
                     page, cursor, rowCount) Then Exit Function

    ' Do not trust the server-side filter alone.
    Set result = New Collection
    For Each t In page
        If InStr(t.Summary, externalNumber) > 0 Then result.Add t
    Next t

    pExternal.Item(externalNumber) = Array(Clock(), result)
    Set ITicketIndex_FindByExternalNumber = result
End Function

Private Function ITicketIndex_FindOpenBySubject(ByVal normalizedSubject As String) As Collection
    Set ITicketIndex_FindOpenBySubject = New Collection
    If Len(normalizedSubject) = 0 Then Exit Function

    pLastError = ""
    If Not OpenListIsFresh() Then ReloadOpenTickets
    If pOpenBySubject Is Nothing Then Exit Function

    If pOpenBySubject.Exists(normalizedSubject) Then
        Set ITicketIndex_FindOpenBySubject = pOpenBySubject.Item(normalizedSubject)
    End If
End Function

Private Sub ITicketIndex_RememberConversation(ByVal conversationId As String, ByVal ticketId As Long)
    RememberConversation conversationId, ticketId
End Sub

Private Function ITicketIndex_FindOpenByStem(ByVal stemPrefix As String) As Collection
    Dim result As Collection
    Dim k As Variant
    Dim t As MatchTicket

    Set result = New Collection
    Set ITicketIndex_FindOpenByStem = result
    If Len(stemPrefix) = 0 Then Exit Function

    pLastError = ""
    If Not OpenListIsFresh() Then ReloadOpenTickets
    If pOpenByStem Is Nothing Then Exit Function

    For Each k In pOpenByStem.Keys
        If LearnedRules.StemStartsWith(CStr(k), stemPrefix) Then
            For Each t In pOpenByStem.Item(k)
                result.Add t
            Next t
        End If
    Next k
End Function

Private Function ITicketIndex_FindByConversation(ByVal conversationId As String) As Long
    If Len(conversationId) = 0 Then Exit Function
    If pConversations.Exists(conversationId) Then ITicketIndex_FindByConversation = CLng(pConversations.Item(conversationId))
End Function

' ---------------------------------------------------------------------------
'  Implementation
' ---------------------------------------------------------------------------

Private Function Transport() As INinjaTransport
    If pTransport Is Nothing Then Set pTransport = New NinjaTransport
    Set Transport = pTransport
End Function

Private Function Clock() As Date
    If TestNow <> 0 Then
        Clock = TestNow
    Else
        Clock = Now
    End If
End Function

Private Function IsFresh(ByVal loadedAt As Date) As Boolean
    IsFresh = (Clock() - loadedAt) * 1440# < CacheMinutes
End Function

Private Function OpenListIsFresh() As Boolean
    If pOpenBySubject Is Nothing Then Exit Function
    OpenListIsFresh = IsFresh(pOpenLoadedAt)
End Function

' Posts one board query. False (and LastError) if there is no usable answer.
Private Function QueryPage(ByVal requestBody As String, ByRef tickets As Collection, ByRef cursor As Long, ByRef rowCount As Long) As Boolean
    Dim status As Long
    Dim body As String
    Dim errorText As String

    body = Transport().HttpPostJson("ticketing/trigger/board/" & CStr(TicketConfig.NINJA_ALL_TICKETS_BOARD) & "/run", _
                                    requestBody, status, errorText)
    If status = 0 Then
        pLastError = "No response: " & errorText
        Exit Function
    End If
    If status <> 200 Then
        pLastError = "Ticket query failed (HTTP " & status & ")"
        Exit Function
    End If
    If Not NinjaTicketLogic.ParseBoardRun(body, tickets, cursor, rowCount) Then
        pLastError = "Ticket query: unreadable answer"
        Exit Function
    End If
    QueryPage = True
End Function

' Loads all non-closed tickets page by page. The cache is only replaced when every page was read.
Private Sub ReloadOpenTickets()
    Dim grouped As Object
    Dim groupedStems As Object
    Dim stemKey As String
    Dim page As Collection
    Dim cursor As Long
    Dim rowCount As Long
    Dim pageNo As Long
    Dim size As Long
    Dim t As MatchTicket
    Dim key As String
    Dim list As Collection

    size = PageSize
    If size < 1 Then size = 1
    Set grouped = CreateObject("Scripting.Dictionary")
    Set groupedStems = CreateObject("Scripting.Dictionary")

    Do
        pageNo = pageNo + 1
        If Not QueryPage(NinjaTicketLogic.BuildBoardRunBody(size, cursor, "status", "not_equals", CStr(NinjaTicketLogic.CLOSED_STATUS_ID)), _
                         page, cursor, rowCount) Then Exit Sub

        For Each t In page
            If t.IsOpen Then
                key = TicketMatcher.NormalizeSubject(t.Summary)
                If Len(key) > 0 Then
                    If Not grouped.Exists(key) Then grouped.Add key, New Collection
                    Set list = grouped.Item(key)
                    list.Add t
                    stemKey = TicketMatcher.SubjectStem(key)
                    If Len(stemKey) > 0 Then
                        If Not groupedStems.Exists(stemKey) Then groupedStems.Add stemKey, New Collection
                        Set list = groupedStems.Item(stemKey)
                        list.Add t
                    End If
                End If
            End If
        Next t

        If rowCount < size Then Exit Do
        If cursor <= 0 Then
            pLastError = "Ticket query: no cursor for the next page"
            Exit Sub
        End If
    Loop While pageNo < MAX_PAGES

    Set pOpenBySubject = grouped
    Set pOpenByStem = groupedStems
    pOpenLoadedAt = Clock()
End Sub
