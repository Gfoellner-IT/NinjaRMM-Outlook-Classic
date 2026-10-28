Attribute VB_Name = "AuditReportTests"
' AuditReportTests v0.1.0
'
' Tests of the audit log report (AuditReport) with a synthetic log: rule statistics, level
' recommendation, hot spots, pattern candidates with the check against later decisions, reading
' the monthly files and writing the report. No Outlook, no network, no access to the real audit folder.

Option Explicit

Private pDay As Long

' One audit line; every call advances the time by one day.
Private Function Line(ByVal ev As String, ByVal source As String, ByVal outcome As String, ByVal rule As String, _
                      ByVal ticketId As Long, ByVal action As String, ByVal domain As String, _
                      ByVal subject As String, ByVal conv As String) As String
    pDay = pDay + 1
    Line = AuditLogic.BuildMatchEvent(DateSerial(2026, 8, 1) + pDay + TimeSerial(9, 0, 0), ev, source, outcome, rule, ticketId, action, _
                                      domain, subject, AuditLogic.Hash32Hex(subject), "", "", conv, "m" & CStr(pDay), False) & vbLf
End Function

Private Function SampleLog() As String
    Dim i As Long
    Dim s As String

    pDay = 0
    For i = 1 To 5
        s = s & Line("auto_decision", "sweep", "none", "None", 0, "none", "x.example", "", "n" & i)
    Next i
    For i = 1 To 34       ' the Subject rule: 34 confirmed ...
        s = s & Line("user_confirmed", "review", "found", "Subject", 20106, "moved", "example.com", "example order total issue project " & i, "c" & i)
    Next i
    s = s & Line("user_rejected", "review", "found", "Subject", 20000, "rejected", "newsletter.example", "angebot der woche 12", "r1")   ' ... and 1 rejected
    For i = 1 To 4        ' a domain whose proposals are always rejected
        s = s & Line("user_rejected", "review", "found", "ExternalNumber", 20001, "rejected", "spam.example", "werbung heute " & i, "s" & i)
    Next i
    For i = 1 To 5        ' a vendor pattern, 5 tickets
        s = s & Line("user_confirmed", "review", "found", "ExternalNumber", 20108 + i, "moved", "vendor.example", "item235 - error " & i, "v" & i)
    Next i
    For i = 1 To 2        ' too little support
        s = s & Line("user_confirmed", "review", "found", "Subject", 30000, "moved", "y.example", "kurzes thema " & i, "y" & i)
    Next i
    For i = 1 To 6        ' a pattern that was also rejected once
        s = s & Line("user_confirmed", "review", "found", "Subject", 31000, "moved", "mixed.example", "gemischtes thema " & i, "g" & i)
    Next i
    s = s & Line("user_rejected", "review", "found", "Conversation", 31001, "rejected", "mixed.example", "gemischtes thema 99", "g99")
    s = s & Line("auto_decision", "watcher", "found", "Subject", 20106, "proposed", "example.com", "example order total issue", "p1")
    SampleLog = s
End Function

Private Function Report() As String
    Report = AuditReport.BuildReport(AuditReport.ParseLines(SampleLog()), DateSerial(2026, 10, 9) + TimeSerial(18, 0, 0))
End Function

' --- small helpers ------------------------------------------------------------------

Public Sub Test_GermanText_Markers()
    TestAssert.AssertEqual "Pr" & ChrW$(252) & "fen", GermanText.De("Pr[ue]fen")
    TestAssert.AssertEqual ChrW$(196) & ChrW$(214) & ChrW$(220) & ChrW$(228) & ChrW$(246) & ChrW$(252) & ChrW$(223), GermanText.De("[Ae][Oe][Ue][ae][oe][ue][ss]")
    TestAssert.AssertEqual "plain", GermanText.De("plain")
End Sub

Public Sub Test_SubjectStem()
    TestAssert.AssertEqual "item# - error #", AuditReport.SubjectStem("item235 - error 4711 in form220")
    TestAssert.AssertEqual "example order total issue", AuditReport.SubjectStem("example order total issue project 7")
    TestAssert.AssertEqual "#", AuditReport.SubjectStem("12345")
    TestAssert.AssertEqual "a b c", AuditReport.SubjectStem("a  b   c")
    TestAssert.AssertEqual "a b", AuditReport.SubjectStem("a b c d", 2)
    TestAssert.AssertEqual "", AuditReport.SubjectStem("")
    TestAssert.AssertEqual "", AuditReport.SubjectStem("   ")
End Sub

Public Sub Test_RecommendLevel()
    TestAssert.AssertEqual GermanText.De("noch keine R[ue]ckmeldung"), AuditReport.RecommendLevel(0, 0, 1)
    TestAssert.AssertEqual GermanText.De("Stufe 3 erw[ae]gen"), AuditReport.RecommendLevel(30, 97, 1)
    TestAssert.AssertEqual GermanText.De("Stufe 3 erw[ae]gen"), AuditReport.RecommendLevel(30, 97, 2)
    TestAssert.AssertEqual GermanText.De("l[ae]uft gut, Stufe bleibt"), AuditReport.RecommendLevel(30, 100, 3)
    TestAssert.AssertEqual GermanText.De("Stufe senken oder Regel versch[ae]rfen"), AuditReport.RecommendLevel(10, 79, 2)
    TestAssert.AssertEqual GermanText.De("weiter beobachten"), AuditReport.RecommendLevel(5, 50, 1), "too few decisions to judge"
    TestAssert.AssertEqual GermanText.De("weiter beobachten"), AuditReport.RecommendLevel(29, 100, 1), "not enough decisions for level 3"
    TestAssert.AssertEqual GermanText.De("weiter beobachten"), AuditReport.RecommendLevel(20, 90, 1)
End Sub

Public Sub Test_ParseLines_SkipsBrokenLines()
    Dim records As Collection

    Set records = AuditReport.ParseLines("{""a"":1}" & vbCrLf & "not json" & vbCrLf & vbCrLf & "[1,2]" & vbLf & "{""b"":2}")
    TestAssert.AssertEqual 2, records.Count, "only the two JSON objects"
    TestAssert.AssertEqual 0, AuditReport.ParseLines("").Count
End Sub

' --- the report ------------------------------------------------------------------------

Public Sub Test_Report_RulesTableAndRecommendation()
    Dim text As String

    MatchLevels.UseAll LevelConfirm
    text = Report()
    MatchLevels.ResetState

    TestAssert.AssertTrue InStr(text, "## Regeln") > 0, text
    TestAssert.AssertTrue InStr(text, "| Subject | 2 | 1 | 1 | 0 | 42 | 1 | 97 % | " & GermanText.De("Stufe 3 erw[ae]gen") & " |") > 0, "Subject row: " & text
End Sub

Public Sub Test_Report_Overview()
    Dim text As String

    text = Report()
    TestAssert.AssertTrue InStr(text, "Zeitraum: 2026-08-02 bis 2026-09-29") > 0, text
    TestAssert.AssertTrue InStr(text, GermanText.De("Eintr[ae]ge: 59 (automatische Entscheidungen 6")) > 0, text
End Sub

Public Sub Test_Report_Hotspots()
    Dim text As String

    text = Report()
    TestAssert.AssertTrue InStr(text, "- spam.example (Regel ExternalNumber): 4 abgelehnt, 0 ") > 0, text
    TestAssert.AssertEqual 0, InStr(text, "- newsletter.example"), "a single rejection is no hot spot"
End Sub

Public Sub Test_Report_Candidates()
    Dim text As String

    text = Report()
    TestAssert.AssertTrue InStr(text, "| example.com | example order total issue | 34 | #20106 | 34 |") > 0, "main pattern: " & text
    TestAssert.AssertTrue InStr(text, "| vendor.example | item# - error # | 5 | #20109, #20110, #20111, #20112, #20113 | 5 |") > 0, "vendor pattern: " & text
    TestAssert.AssertEqual 0, InStr(text, "y.example"), "too little support"
    TestAssert.AssertEqual 0, InStr(text, "mixed.example"), "a pattern with a rejection is no candidate"
End Sub

Public Sub Test_Report_CandidateCheckAgainstLaterDecisions()
    Dim text As String

    text = Report()
    ' 53 decided mails, 70 % = 37 for training: the main pattern has 34 confirmations in the training
    ' part and none later; the vendor pattern sits in the later part.
    TestAssert.AssertTrue InStr(text, "(37 von 53)") > 0, text
    TestAssert.AssertTrue InStr(text, "| 34 | 0 ok / 0 abgelehnt | keine Testdaten |") > 0, "no later data for the main pattern: " & text
    TestAssert.AssertTrue InStr(text, "| 0 | 5 ok / 0 abgelehnt | stabil, pr" & ChrW$(252) & "fen |") > 0, "vendor pattern confirmed later: " & text
End Sub

Public Sub Test_Report_MailsWithoutMatch()
    Dim text As String

    text = Report()
    TestAssert.AssertTrue InStr(text, "5 Entscheidungen ohne Treffer") > 0, text
    TestAssert.AssertTrue InStr(text, "- x.example: 5") > 0
End Sub

Public Sub Test_Report_ContainsNoMailTextsOrFullSubjects()
    Dim text As String

    text = Report()
    TestAssert.AssertEqual 0, InStr(text, "projekt"), "only the stem of a subject is shown"
    TestAssert.AssertEqual 0, InStr(text, "werbung"), "subjects of rejected mails are not shown"
    TestAssert.AssertTrue InStr(text, "keine Absendernamen und keine Mailtexte") > 0
End Sub

Public Sub Test_Report_EmptyLog()
    Dim text As String

    text = AuditReport.BuildReport(New Collection, Now)
    TestAssert.AssertTrue InStr(text, GermanText.De("noch keine Eintr[ae]ge")) > 0 Or InStr(text, GermanText.De("enth[ae]lt noch keine Eintr[ae]ge")) > 0, text
    TestAssert.AssertTrue InStr(text, "noch keine Kandidaten") > 0 Or InStr(text, GermanText.De("noch keine Kandidaten")) > 0
End Sub

' --- files ------------------------------------------------------------------------------

Private Function TempFolder() As String
    TempFolder = Environ$("TEMP") & "\ninja_report_test_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd() * 100000))
End Function

Private Sub WriteFile(ByVal path As String, ByVal text As String)
    Dim fso As Object
    Dim f As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(path, True, False)
    f.Write text
    f.Close
End Sub

Public Sub Test_LoadRecords_ReadsMonthlyFilesOldestFirst()
    Dim fso As Object
    Dim base As String
    Dim records As Collection

    base = TempFolder()
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder base
    WriteFile base & "\match_audit_2026-10.jsonl", "{""n"":3}" & vbCrLf & "{""n"":4}" & vbCrLf
    WriteFile base & "\match_audit_2026-09.jsonl", "{""n"":1}" & vbCrLf & "{""n"":2}" & vbCrLf
    WriteFile base & "\match_audit_2024-01.jsonl", "{""n"":99}" & vbCrLf
    WriteFile base & "\notes.txt", "{""n"":98}" & vbCrLf

    Set records = AuditReport.LoadRecords(base, 12, DateSerial(2026, 10, 9))
    TestAssert.AssertEqual 4, records.Count, "expired and foreign files are skipped"
    TestAssert.AssertEqual 1, CLng(records(1)("n")), "oldest file first"
    TestAssert.AssertEqual 4, CLng(records(4)("n")), "newest file last"
    TestAssert.AssertEqual 4, records.Count
    TestAssert.AssertEqual 0, AuditReport.LoadRecords(base & "\missing", 12, Now).Count, "missing folder"
    On Error Resume Next
    fso.DeleteFolder base, True
End Sub

Public Sub Test_WriteReport_Utf8AndNames()
    Dim base As String
    Dim path As String
    Dim stream As Object
    Dim text As String
    Dim fso As Object

    base = TempFolder() & "\reports"
    path = AuditReport.WriteReport("# " & GermanText.De("Pr[ue]fung") & vbLf, base, DateSerial(2026, 10, 9))

    TestAssert.AssertEqual base & "\match_review_2026-10-09.md", path
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    text = stream.ReadText
    stream.Close
    TestAssert.AssertTrue InStr(text, "Pr" & ChrW$(252) & "fung") > 0, "umlaut survives"

    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    fso.DeleteFolder Left$(base, Len(base) - Len("\reports")), True
End Sub

Public Sub Test_ReportNamesAndPrompt()
    TestAssert.AssertEqual "match_review_2026-10-09.md", AuditReport.ReportFileName(DateSerial(2026, 10, 9))
    TestAssert.AssertTrue Right$(AuditReport.ReportFolder(), Len("NinjaRMM-Outlook-Classic\reports")) = "NinjaRMM-Outlook-Classic\reports", AuditReport.ReportFolder()
    TestAssert.AssertTrue InStr(AuditReport.ReviewPromptText("C:\x\r.md"), "Pr" & ChrW$(252) & "fen wir gemeinsam die aktuellen Regeln und Entscheidungen") > 0
    TestAssert.AssertTrue InStr(AuditReport.ReviewPromptText("C:\x\r.md"), "neue Regeln zu erstellen. Bericht: C:\x\r.md") > 0
End Sub

Public Sub Test_Report_RepeatedDecisionsAboutTheSameMailCountOnce()
    Dim lines As String
    Dim i As Long
    Dim text As String

    For i = 1 To 3       ' three sweeps log the same decision about the same mail
        lines = lines & AuditLogic.BuildMatchEvent(DateSerial(2026, 9, i), "auto_decision", "sweep", "found", "Subject", 20106, "proposed", _
            "example.com", "example order total issue", "x", "", "", "c1", "same0001", False) & vbLf
    Next i
    lines = lines & AuditLogic.BuildMatchEvent(DateSerial(2026, 9, 4), "auto_decision", "sweep", "none", "None", 0, "none", _
        "example.com", "", "y", "", "", "c2", "other002", False) & vbLf
    text = AuditReport.BuildReport(AuditReport.ParseLines(lines), DateSerial(2026, 10, 9))

    TestAssert.AssertTrue InStr(text, GermanText.De("Eintr[ae]ge: 2 (automatische Entscheidungen 2")) > 0, text
    TestAssert.AssertTrue InStr(text, "| Subject | " & MatchLevels.LevelFor(RuleSubject) & " | 1 | 1 | 0 |") > 0, "found and proposed once: " & text
End Sub
