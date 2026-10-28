Attribute VB_Name = "NinjaApi"
' NinjaApi v0.2.0
'
' Read-only NinjaOne ticket queries used by the ticket automation:
' state (open/closed/not found/unavailable), subject and closing date.
' No Outlook types. The HTTP layer is replaceable (UseTransport) for unit tests;
' the response logic lives in NinjaTicketLogic.
'
' NinjaTicketState values:
'   NinjaTicketOpen, NinjaTicketClosed (closed or deleted), NinjaTicketNotFound (HTTP 404),
'   NinjaTicketUnavailable (offline, timeout, 5xx, auth problem, unreadable response).
' LastError describes the last problem for the log (never contains secrets).

Option Explicit

Private pTransport As INinjaTransport
Private pLastError As String

' ---------------------------------------------------------------------------
'  Public API
' ---------------------------------------------------------------------------

' State and current subject of a ticket with a single request.
' subject is only filled for open and closed tickets ("" otherwise).
Public Function GetTicketInfo(ByVal ticketId As Long, ByRef subject As String) As NinjaTicketState
    Dim status As Long
    Dim body As String

    subject = ""
    GetTicketInfo = NinjaTicketUnavailable
    If Not Fetch("ticketing/ticket/" & CStr(ticketId), ticketId, status, body) Then Exit Function

    GetTicketInfo = NinjaTicketLogic.ParseTicketState(status, body)
    Select Case GetTicketInfo
        Case NinjaTicketOpen, NinjaTicketClosed
            subject = NinjaTicketLogic.ParseTicketSubject(body)
        Case NinjaTicketUnavailable
            pLastError = "Unexpected answer for ticket " & ticketId & " (HTTP " & status & ")"
    End Select
End Function

Public Function GetTicketState(ByVal ticketId As Long) As NinjaTicketState
    Dim ignoredSubject As String

    GetTicketState = GetTicketInfo(ticketId, ignoredSubject)
End Function

' True if the ticket is closed, deleted or does not exist any more (these folders are archived).
' False for open tickets and whenever the state cannot be determined.
Public Function IsTicketClosed(ByVal ticketId As Long) As Boolean
    Dim state As NinjaTicketState

    state = GetTicketState(ticketId)
    IsTicketClosed = (state = NinjaTicketClosed Or state = NinjaTicketNotFound)
End Function

' Current subject of the ticket, "" if it cannot be read.
Public Function GetTicketSubject(ByVal ticketId As Long) As String
    Dim status As Long
    Dim body As String

    If Not Fetch("ticketing/ticket/" & CStr(ticketId), ticketId, status, body) Then Exit Function
    If status <> 200 Then
        pLastError = "Could not read ticket " & ticketId & " (HTTP " & status & ")"
        Exit Function
    End If
    GetTicketSubject = NinjaTicketLogic.ParseTicketSubject(body)
End Function

' Time (UTC) when the ticket was closed or deleted, 0 if unknown.
Public Function GetTicketClosedDate(ByVal ticketId As Long) As Date
    Dim status As Long
    Dim body As String

    If Not Fetch("ticketing/ticket/" & CStr(ticketId) & "/log-entry?type=SAVE", ticketId, status, body) Then Exit Function
    If status <> 200 Then
        pLastError = "Could not read the log of ticket " & ticketId & " (HTTP " & status & ")"
        Exit Function
    End If
    GetTicketClosedDate = NinjaTicketLogic.ParseClosedDate(body)
End Function

Public Property Get LastError() As String
    LastError = pLastError
End Property

' ---------------------------------------------------------------------------
'  Test seams
' ---------------------------------------------------------------------------

Public Sub UseTransport(ByVal transport As INinjaTransport)
    Set pTransport = transport
End Sub

Public Sub ResetState()
    Set pTransport = Nothing
    pLastError = ""
End Sub

' ---------------------------------------------------------------------------
'  Implementation
' ---------------------------------------------------------------------------

' Runs the GET request. Returns False (and sets LastError) if there was no response at all.
Private Function Fetch(ByVal path As String, ByVal ticketId As Long, ByRef status As Long, ByRef body As String) As Boolean
    Dim errorText As String

    pLastError = ""
    If ticketId <= 0 Then
        pLastError = "Invalid ticket id " & ticketId
        Exit Function
    End If

    If pTransport Is Nothing Then Set pTransport = New NinjaTransport

    body = pTransport.HttpGet(path, status, errorText)
    If status = 0 Then
        pLastError = "No response: " & errorText
        Exit Function
    End If
    Fetch = True
End Function
