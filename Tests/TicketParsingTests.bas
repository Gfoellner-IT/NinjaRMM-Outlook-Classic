Attribute VB_Name = "TicketParsingTests"
' TicketParsingTests v0.1.0
'
' Unit tests for the pure parsing functions of the ticket automation.

Option Explicit

Public Sub Test_ParseSubject_Valid()
    Dim number As String
    Dim remaining As String

    TestAssert.AssertTrue TicketParsing.ParseSubject("[example] (#3621) Printer broken", number, remaining)
    TestAssert.AssertEqual "#3621", number
    TestAssert.AssertEqual " Printer broken", remaining

    TestAssert.AssertTrue TicketParsing.ParseSubject("[example](#7)x", number, remaining), "no spaces"
    TestAssert.AssertEqual "#7", number
    TestAssert.AssertEqual "x", remaining

    TestAssert.AssertTrue TicketParsing.ParseSubject("[EXAMPLE] (#12) x", number, remaining), "case-insensitive"
    TestAssert.AssertEqual "#12", number
End Sub

Public Sub Test_ParseSubject_Invalid()
    Dim number As String
    Dim remaining As String

    TestAssert.AssertFalse TicketParsing.ParseSubject("Re: [example] (#1) x", number, remaining), "must start with the tag"
    TestAssert.AssertFalse TicketParsing.ParseSubject("[other] (#1) x", number, remaining), "other tag"
    TestAssert.AssertFalse TicketParsing.ParseSubject("[example] (no number)", number, remaining), "no number"
    TestAssert.AssertFalse TicketParsing.ParseSubject("", number, remaining), "empty"
    TestAssert.AssertEqual "", number, "outputs are cleared"
End Sub

Public Sub Test_BuildFolderName()
    TestAssert.AssertEqual "#3621 (Printer broken)", TicketParsing.BuildFolderName("#3621", " Printer broken")
    TestAssert.AssertEqual "#3621 ()", TicketParsing.BuildFolderName("#3621", "")
    TestAssert.AssertEqual "#3621 (a b)", TicketParsing.BuildFolderName("#3621", "   a b   ")
End Sub

Public Sub Test_BuildFolderName_ShortensTicketPrefix()
    Dim text As String

    text = " TICKET (B" & ChrW$(252) & "ro - IT) / Printer defekt"
    TestAssert.AssertEqual "#3621 (IT / Printer defekt)", TicketParsing.BuildFolderName("#3621", text)

    text = " TICKET (B" & ChrW$(252) & "ro - HR) / Urlaub"
    TestAssert.AssertEqual "#5 (HR / Urlaub)", TicketParsing.BuildFolderName("#5", text)
End Sub

Public Sub Test_TicketIdFromFolderName()
    TestAssert.AssertEqual 3621, TicketParsing.TicketIdFromFolderName("#3621 (x)")
    TestAssert.AssertEqual 3621, TicketParsing.TicketIdFromFolderName("#3621")
    TestAssert.AssertEqual 12, TicketParsing.TicketIdFromFolderName("#12 (a b c)")
    TestAssert.AssertEqual 0, TicketParsing.TicketIdFromFolderName("3621 (x)"), "no #"
    TestAssert.AssertEqual 0, TicketParsing.TicketIdFromFolderName("#abc (x)"), "no digits"
    TestAssert.AssertEqual 0, TicketParsing.TicketIdFromFolderName("#36x (x)"), "digits and letters"
    TestAssert.AssertEqual 0, TicketParsing.TicketIdFromFolderName("#"), "only #"
    TestAssert.AssertEqual 0, TicketParsing.TicketIdFromFolderName(""), "empty"
    TestAssert.AssertEqual 0, TicketParsing.TicketIdFromFolderName("#1234567890 (x)"), "too long for a ticket number"
    TestAssert.AssertEqual 0, TicketParsing.TicketIdFromFolderName("Archiv"), "archive folder"
End Sub

Public Sub Test_FolderMatchesTicket()
    TestAssert.AssertTrue TicketParsing.FolderMatchesTicket("#36 (x)", "#36")
    TestAssert.AssertTrue TicketParsing.FolderMatchesTicket("#36", "#36"), "exact name"
    TestAssert.AssertFalse TicketParsing.FolderMatchesTicket("#3621 (x)", "#36"), "longer ticket number must not match"
    TestAssert.AssertFalse TicketParsing.FolderMatchesTicket("#36(x)", "#36"), "number must be followed by a space"
    TestAssert.AssertFalse TicketParsing.FolderMatchesTicket("Archiv", "#36")
    TestAssert.AssertFalse TicketParsing.FolderMatchesTicket("", "#36")
    TestAssert.AssertFalse TicketParsing.FolderMatchesTicket("#36 (x)", ""), "empty ticket number"
End Sub

Public Sub Test_IsClosedMail()
    TestAssert.AssertTrue TicketParsing.IsClosedMail("Hello" & vbCrLf & "Status: Offen -> Geschlossen" & vbCrLf & "Bye")
    TestAssert.AssertTrue TicketParsing.IsClosedMail("status: geschlossen"), "case-insensitive"
    TestAssert.AssertFalse TicketParsing.IsClosedMail("Status: Offen"), "open"
    TestAssert.AssertFalse TicketParsing.IsClosedMail("Geschlossen"), "word without status"
    TestAssert.AssertFalse TicketParsing.IsClosedMail(""), "empty"
End Sub

Public Sub Test_ArchivePathNames()
    Dim d As Date

    d = DateSerial(2026, 10, 6)
    TestAssert.AssertEqual "2026", TicketParsing.ArchiveYearName(d)
    TestAssert.AssertTrue Left$(TicketParsing.ArchiveMonthName(d), 5) = "10 - ", "month number first"
    TestAssert.AssertTrue Len(TicketParsing.ArchiveMonthName(d)) > 5, "followed by the month name"
    TestAssert.AssertTrue Left$(TicketParsing.ArchiveMonthName(DateSerial(2027, 1, 31)), 5) = "01 - ", "two digit month"
End Sub

Public Sub Test_ArchivePathNames_ZeroMeansToday()
    TestAssert.AssertEqual CStr(Year(Date)), TicketParsing.ArchiveYearName(0)
    TestAssert.AssertEqual TicketParsing.ArchiveMonthName(Date), TicketParsing.ArchiveMonthName(0)
End Sub

' Regression: the synthetic sender name contains a u umlaut. The source must stay ASCII (ChrW), otherwise the
' VBA import turns it into two characters and every comparison fails.
Public Sub Test_AllowedSender_ContainsRealUmlaut()
    TestAssert.AssertEqual "IT Support | B" & ChrW$(252) & "ro", TicketConfig.AllowedSender()
    TestAssert.AssertEqual 252, AscW(Mid$(TicketConfig.AllowedSender(), 15, 1)), "character 15 is u umlaut"
    TestAssert.AssertEqual 17, Len(TicketConfig.AllowedSender()), "no extra characters from a wrong encoding"
End Sub

Public Sub Test_RegexTicketReplace_ContainsRealUmlaut()
    TestAssert.AssertTrue InStr(TicketConfig.RegexTicketReplace(), "B" & ChrW$(252) & "ro") > 0
    TestAssert.AssertEqual 0, InStr(TicketConfig.RegexTicketReplace(), ChrW$(195)), "no A-tilde from a UTF-8 file read as ANSI"
End Sub
