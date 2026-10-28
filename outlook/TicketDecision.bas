Attribute VB_Name = "TicketDecision"
' TicketDecision v0.1.0
'
' Pure decision: is a ticket folder archived, and on what basis?
'
'   API state     | decision
'   --------------+--------------------------------------------------------------
'   Closed        | ArchiveByApi
'   NotFound      | ArchiveByApi (the ticket was deleted; same behaviour as before)
'   Open          | ArchiveNo (the mail text is never consulted)
'   Unavailable   | ArchiveByMailFallback if the fallback is enabled AND a mail in the
'                 | folder says "Status: ... Geschlossen", otherwise ArchiveNo

Option Explicit

Public Enum ArchiveDecision
    ArchiveNo = 0
    ArchiveByApi = 1
    ArchiveByMailFallback = 2
End Enum

' True if the mail texts are needed to decide (avoids scanning mails when the API answered).
Public Function NeedsMailCheck(ByVal apiState As NinjaTicketState, ByVal fallbackEnabled As Boolean) As Boolean
    NeedsMailCheck = (apiState = NinjaTicketUnavailable) And fallbackEnabled
End Function

Public Function DecideArchive(ByVal apiState As NinjaTicketState, ByVal mailSaysClosed As Boolean, ByVal fallbackEnabled As Boolean) As ArchiveDecision
    Select Case apiState
        Case NinjaTicketClosed, NinjaTicketNotFound
            DecideArchive = ArchiveByApi
        Case NinjaTicketOpen
            DecideArchive = ArchiveNo
        Case Else
            If fallbackEnabled And mailSaysClosed Then
                DecideArchive = ArchiveByMailFallback
            Else
                DecideArchive = ArchiveNo
            End If
    End Select
End Function
