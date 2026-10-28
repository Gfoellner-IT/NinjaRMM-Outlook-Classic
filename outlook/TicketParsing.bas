Attribute VB_Name = "TicketParsing"
' TicketParsing v0.1.0
'
' Pure functions for ticket numbers, folder names and archive path names.
' Regular expressions are late bound, so no VBA reference is needed.

Option Explicit

' ---------------------------------------------------------------------------
'  Mail subject
' ---------------------------------------------------------------------------

' Splits a ticket mail subject into the ticket number ("#3621") and the rest of the text.
Public Function ParseSubject(ByVal subjectText As String, ByRef ticketNumber As String, ByRef remainingText As String) As Boolean
    Dim re As Object
    Dim matches As Object

    ticketNumber = ""
    remainingText = ""

    Set re = NewRegex(TicketConfig.REGEX_SUBJECT_PATTERN)
    Set matches = re.Execute(subjectText)
    If matches.Count = 0 Then Exit Function

    ticketNumber = "#" & CStr(matches(0).SubMatches(0))
    remainingText = CStr(matches(0).SubMatches(1))
    ParseSubject = True
End Function

' Folder name for a new ticket folder: "#3621 (Text)".
' A leading "TICKET (B..ro - IT) / " is shortened to "IT / ".
Public Function BuildFolderName(ByVal ticketNumber As String, ByVal remainingText As String) As String
    Dim re As Object
    Dim matches As Object
    Dim found As Object

    Set re = NewRegex(TicketConfig.RegexTicketReplace())
    Set matches = re.Execute(remainingText)
    If matches.Count > 0 Then
        Set found = matches(0)
        remainingText = CStr(found.SubMatches(0)) & CStr(found.SubMatches(1)) & _
                        Trim$(Replace(remainingText, CStr(found.Value), ""))
    End If

    BuildFolderName = ticketNumber & " (" & Trim$(remainingText) & ")"
End Function

' ---------------------------------------------------------------------------
'  Folder names
' ---------------------------------------------------------------------------

' Ticket id from a folder name like "#3621 (Text)" or "#3621". 0 if it is not a ticket folder.
Public Function TicketIdFromFolderName(ByVal folderName As String) As Long
    Dim token As String
    Dim spacePos As Long

    If Left$(folderName, 1) <> "#" Then Exit Function

    spacePos = InStr(folderName, " ")
    If spacePos = 0 Then
        token = Mid$(folderName, 2)
    Else
        token = Mid$(folderName, 2, spacePos - 2)
    End If

    If Len(token) = 0 Or Len(token) > 9 Then Exit Function
    If Not IsAllDigits(token) Then Exit Function
    TicketIdFromFolderName = CLng(token)
End Function

' True if folderName belongs to ticketNumber ("#36" matches "#36 (x)" but not "#3621 (x)").
Public Function FolderMatchesTicket(ByVal folderName As String, ByVal ticketNumber As String) As Boolean
    If Len(ticketNumber) = 0 Then Exit Function
    FolderMatchesTicket = (folderName = ticketNumber) Or _
                          (Left$(folderName, Len(ticketNumber) + 1) = ticketNumber & " ")
End Function

' ---------------------------------------------------------------------------
'  Mail body (offline fallback)
' ---------------------------------------------------------------------------

Public Function IsClosedMail(ByVal bodyText As String) As Boolean
    Dim re As Object

    Set re = NewRegex(TicketConfig.REGEX_STATUS_CLOSED)
    IsClosedMail = re.Test(bodyText)
End Function

' ---------------------------------------------------------------------------
'  Archive path: Archiv -> <year> -> <month>
' ---------------------------------------------------------------------------

' Year folder name ("2026"). closedDate = 0 means today.
Public Function ArchiveYearName(ByVal closedDate As Date) As String
    If closedDate = 0 Then closedDate = Date
    ArchiveYearName = CStr(Year(closedDate))
End Function

' Month folder name ("10 - Oktober", the month name follows the system language). closedDate = 0 means today.
Public Function ArchiveMonthName(ByVal closedDate As Date) As String
    If closedDate = 0 Then closedDate = Date
    ArchiveMonthName = Format$(closedDate, "mm") & " - " & Format$(closedDate, "mmmm")
End Function

' ---------------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------------

Private Function NewRegex(ByVal pattern As String) As Object
    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    Set NewRegex = re
End Function

Private Function IsAllDigits(ByVal text As String) As Boolean
    Dim i As Long

    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        If Not (Mid$(text, i, 1) Like "[0-9]") Then Exit Function
    Next i
    IsAllDigits = True
End Function
