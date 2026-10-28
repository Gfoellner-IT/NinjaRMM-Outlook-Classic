Attribute VB_Name = "NinjaTicketLogic"
' NinjaTicketLogic v0.1.0
'
' Pure logic for NinjaOne ticket responses: ticket state, subject and closing date.
' No Outlook types, no network - fully unit testable. Used by NinjaApi.
'
' The status names in the ticket log ("changeDiff.status") are localized display
' names, so the closing status is recognised by a configurable list of names.

Option Explicit

Public Enum NinjaTicketState
    NinjaTicketUnavailable = 0   ' no usable answer: offline, timeout, 5xx, auth problem, unreadable response
    NinjaTicketOpen = 1
    NinjaTicketClosed = 2        ' status closed or ticket marked as deleted
    NinjaTicketNotFound = 3      ' HTTP 404: the ticket does not exist (any more)
End Enum

' statusId of the closed status in NinjaOne.
Public Const CLOSED_STATUS_ID As Long = 6000

' Display names of the closed status in the ticket log, delimited with "|".
Public Const CLOSED_STATUS_NAMES As String = "|Geschlossen|Closed|"

' ---------------------------------------------------------------------------
'  Ticket state
' ---------------------------------------------------------------------------

Public Function ParseTicketState(ByVal httpStatus As Long, ByVal body As String) As NinjaTicketState
    Dim ticket As Variant
    Dim statusId As Double
    Dim resultCode As String

    Select Case httpStatus
        Case 200
            If Not JsonLite.TryParse(body, ticket) Then Exit Function
            If TypeName(ticket) <> "Dictionary" Then Exit Function

            If JsonLite.GetBoolean(ticket, "deleted", False) Then
                ParseTicketState = NinjaTicketClosed
                Exit Function
            End If

            If Not JsonLite.TryGetNumber(ticket, "status.statusId", statusId) Then Exit Function
            If CLng(statusId) = CLOSED_STATUS_ID Then
                ParseTicketState = NinjaTicketClosed
            Else
                ParseTicketState = NinjaTicketOpen
            End If

        Case 404
            ' The body is usually empty. If it names a different problem than "ticket not found",
            ' this is not a statement about the ticket.
            If JsonLite.TryParse(body, ticket) Then
                resultCode = JsonLite.GetString(ticket, "resultCode", "")
                If Len(resultCode) > 0 And resultCode <> "TICKET_NOT_FOUND" Then Exit Function
            End If
            ParseTicketState = NinjaTicketNotFound
    End Select
End Function

' Subject of a ticket response, "" if the response is not a usable ticket.
Public Function ParseTicketSubject(ByVal body As String) As String
    Dim ticket As Variant

    If Not JsonLite.TryParse(body, ticket) Then Exit Function
    If TypeName(ticket) <> "Dictionary" Then Exit Function
    ParseTicketSubject = JsonLite.GetString(ticket, "subject", "")
End Function

' ---------------------------------------------------------------------------
'  Closing date
' ---------------------------------------------------------------------------

' Time (UTC) of the newest log entry that closed or deleted the ticket.
' Returns 0 if the log contains no such entry or is not readable.
Public Function ParseClosedDate(ByVal logJson As String) As Date
    Dim entries As Variant
    Dim entry As Variant
    Dim newest As Double

    If Not JsonLite.TryParse(logJson, entries) Then Exit Function
    If TypeName(entries) <> "Collection" Then Exit Function

    For Each entry In entries
        newest = NewerClosingTime(entry, newest)
    Next entry

    If newest > 0 Then ParseClosedDate = UnixToDate(newest)
End Function

' Unix seconds (UTC) to a VBA Date. The result is UTC, not local time.
Public Function UnixToDate(ByVal seconds As Double) As Date
    UnixToDate = DateSerial(1970, 1, 1) + Fix(seconds) / 86400#
End Function

Public Function IsClosedStatusName(ByVal name As String) As Boolean
    If Len(name) = 0 Then Exit Function
    IsClosedStatusName = (InStr(1, CLOSED_STATUS_NAMES, "|" & name & "|", vbTextCompare) > 0)
End Function

' ---------------------------------------------------------------------------
'  Board queries (ticket lists)
' ---------------------------------------------------------------------------

' Request body for POST ticketing/trigger/board/{id}/run: page size, optional cursor
' (lastCursorId of the previous page, 0 = first page) and optional single filter
' (filterField "" = no filter). Example: {"pageSize":200,"filters":[{"field":"summary",
' "operator":"contains","value":"900000105"}]}
Public Function BuildBoardRunBody(ByVal pageSize As Long, ByVal cursor As Long, ByVal filterField As String, _
                                  ByVal filterOperator As String, ByVal filterValue As String) As String
    Dim body As String

    body = "{""pageSize"":" & CStr(pageSize)
    If cursor > 0 Then body = body & ",""lastCursorId"":" & CStr(cursor)
    If Len(filterField) > 0 Then
        body = body & ",""filters"":[{""field"":""" & JsonEscape(filterField) & """,""operator"":""" & _
               JsonEscape(filterOperator) & """,""value"":""" & JsonEscape(filterValue) & """}]"
    End If
    BuildBoardRunBody = body & "}"
End Function

' Reads one page of a board query. tickets receives MatchTicket objects (IsOpen = status is not
' closed), cursor the lastCursorId to request the next page with (0 if the answer has none),
' rowCount the number of rows in the page. Returns False if the answer is not a board result.
Public Function ParseBoardRun(ByVal body As String, ByRef tickets As Collection, ByRef cursor As Long, ByRef rowCount As Long) As Boolean
    Dim root As Variant
    Dim rows As Object
    Dim row As Variant
    Dim ticket As MatchTicket
    Dim number As Double
    Dim statusId As Double

    Set tickets = New Collection
    cursor = 0
    rowCount = 0

    If Not JsonLite.TryParse(body, root) Then Exit Function
    If TypeName(root) <> "Dictionary" Then Exit Function
    Set rows = JsonLite.GetNode(root, "data")
    If rows Is Nothing Then Exit Function
    If TypeName(rows) <> "Collection" Then Exit Function

    If JsonLite.TryGetNumber(root, "metadata.lastCursorId", number) Then cursor = CLng(number)
    rowCount = rows.Count

    For Each row In rows
        If IsObject(row) Then
            If TypeName(row) = "Dictionary" Then
                If JsonLite.TryGetNumber(row, "id", number) Then
                    Set ticket = New MatchTicket
                    ticket.Id = CLng(number)
                    ticket.Summary = JsonLite.GetString(row, "summary", "")
                    ticket.IsOpen = JsonLite.TryGetNumber(row, "status.statusId", statusId) And (CLng(statusId) <> CLOSED_STATUS_ID)
                    tickets.Add ticket
                End If
            End If
        End If
    Next row

    ParseBoardRun = True
End Function

' ---------------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------------

Private Function JsonEscape(ByVal text As String) As String
    text = Replace(text, "\", "\\")
    text = Replace(text, """", "\""")
    text = Replace(text, vbCr, "\r")
    text = Replace(text, vbLf, "\n")
    text = Replace(text, vbTab, "\t")
    JsonEscape = text
End Function

' Returns the entry's time if it is a closing/deleting entry and newer than current, else current.
Private Function NewerClosingTime(ByVal entry As Variant, ByVal current As Double) As Double
    Dim entryTime As Double
    Dim isClosing As Boolean

    NewerClosingTime = current
    If Not IsObject(entry) Then Exit Function
    If TypeName(entry) <> "Dictionary" Then Exit Function

    isClosing = IsClosedStatusName(JsonLite.GetString(entry, "changeDiff.status.new", "")) Or _
                (JsonLite.GetString(entry, "type", "") = "DELETE")
    If Not isClosing Then Exit Function

    If Not JsonLite.TryGetNumber(entry, "createTime", entryTime) Then Exit Function
    If entryTime > current Then NewerClosingTime = entryTime
End Function
