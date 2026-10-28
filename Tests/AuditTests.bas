Attribute VB_Name = "AuditTests"
' AuditTests v0.1.0
'
' Tests of the match audit log: pure functions (AuditLogic), the file sink in a temporary folder,
' the audit line (TicketAudit.BuildLine) and the audit entries written by the matching flow.
' No Outlook, no network. The real audit folder is never touched.

Option Explicit

Private pSink As FakeAuditSink
Private pStore As FakeTicketStore
Private pIndex As FakeTicketIndex
Private pTransport As FakeNinjaTransport

Private Sub Setup()
    TicketLog.DisableFileLogging
    MatchLevels.UseAll LevelAuto
    MatchProposals.Clear
    TicketLog.Clear
    NinjaApi.ResetState
    TicketAudit.ResetState

    Set pSink = New FakeAuditSink
    TicketAudit.UseSink pSink

    Set pStore = New FakeTicketStore
    Set pTransport = New FakeNinjaTransport
    NinjaApi.UseTransport pTransport
    pTransport.AddResponse "ticketing/ticket/20101", 200, Fixtures.TicketOpen()

    Set pIndex = New FakeTicketIndex
    pIndex.AddTicket 20101, "Example open ticket", True
End Sub

Private Sub Teardown()
    MatchLevels.ResetState
    MatchProposals.Clear
    TicketAudit.ResetState
    NinjaApi.ResetState
    Set pSink = Nothing
    Set pStore = Nothing
    Set pIndex = Nothing
    Set pTransport = Nothing
End Sub

Private Function Foreign(ByVal subject As String) As FakeMail
    Dim mail As FakeMail

    Set mail = New FakeMail
    mail.SenderName = "Max Mustermann"
    mail.SenderDomain = "vendor.example"
    mail.Subject = subject
    mail.ConversationId = "conversation-id-123"
    mail.MessageKey = "0a1b2c3d"
    Set Foreign = mail
End Function

Private Function ForeignIn(ByVal subject As String, ByVal conversationId As String) As FakeMail
    Set ForeignIn = Foreign(subject)
    ForeignIn.ConversationId = conversationId
End Function

' --- hash -----------------------------------------------------------------------

Public Sub Test_Hash32_KnownVectors()
    TestAssert.AssertEqual "811c9dc5", AuditLogic.Hash32Hex("")
    TestAssert.AssertEqual "e40c292c", AuditLogic.Hash32Hex("a")
    TestAssert.AssertEqual "bf9cf968", AuditLogic.Hash32Hex("foobar")
End Sub

Public Sub Test_Hash32_StableAndDistinct()
    TestAssert.AssertEqual AuditLogic.Hash32Hex("conversation-1"), AuditLogic.Hash32Hex("conversation-1")
    TestAssert.AssertNotEqual AuditLogic.Hash32Hex("conversation-1"), AuditLogic.Hash32Hex("conversation-2")
    TestAssert.AssertEqual 8, Len(AuditLogic.Hash32Hex("x" & ChrW$(246) & ChrW$(8364)))
    TestAssert.AssertNotEqual AuditLogic.Hash32Hex("o"), AuditLogic.Hash32Hex(ChrW$(246))
    TestAssert.AssertEqual 8, Len(AuditLogic.Hash32Hex(String$(5000, "z"))), "long text"
End Sub

' --- JSON -----------------------------------------------------------------------

Public Sub Test_JsonString()
    TestAssert.AssertEqual """plain""", AuditLogic.JsonString("plain")
    TestAssert.AssertEqual """a\""b\\c""", AuditLogic.JsonString("a""b\c")
    TestAssert.AssertEqual """l1\nl2\r\tx""", AuditLogic.JsonString("l1" & vbLf & "l2" & vbCr & vbTab & "x")
    TestAssert.AssertEqual """B\" & "u00fcro""", AuditLogic.JsonString("B" & ChrW$(252) & "ro")
    TestAssert.AssertEqual """\" & "u20ac""", AuditLogic.JsonString(ChrW$(8364))
    TestAssert.AssertEqual """\u0001""", AuditLogic.JsonString(Chr$(1))
    TestAssert.AssertEqual """""", AuditLogic.JsonString("")
End Sub

Public Sub Test_BuildMatchEvent()
    Dim line As String
    Dim ts As Date

    ts = DateSerial(2026, 10, 7) + TimeSerial(14, 5, 9)
    line = AuditLogic.BuildMatchEvent(ts, "auto_decision", "watcher", "found", "ExternalNumber", 20103, "log_only", "vendor.example", _
                                      "fehlerhafte verbuchung", "1a2b3c4d", "900000105", "", "aabbccdd", "11223344", False)

    TestAssert.AssertEqual "{""ts"":""2026-10-07T14:05:09"",""event"":""auto_decision"",""source"":""watcher"",""outcome"":""found""," & _
        """rule"":""ExternalNumber"",""ticket"":20103,""action"":""log_only"",""domain"":""vendor.example"",""subject"":""fehlerhafte verbuchung"",""subject_hash"":""1a2b3c4d""," & _
        """external"":""900000105"",""ninja_ids"":"""",""conv"":""aabbccdd"",""msg"":""11223344"",""auto"":false}", line
End Sub

Public Sub Test_BuildMatchEvent_IsSingleLineAscii()
    Dim line As String
    Dim i As Long

    line = AuditLogic.BuildMatchEvent(Now, "auto_decision", "dry_run", "none", "None", 0, "none", "x.example", _
                                      "line" & vbCrLf & "break " & ChrW$(246), "", "", "", "", "", True)
    TestAssert.AssertEqual 0, InStr(line, vbCr) + InStr(line, vbLf)
    For i = 1 To Len(line)
        If AscW(Mid$(line, i, 1)) > 126 Or AscW(Mid$(line, i, 1)) < 32 Then TestAssert.Fail "non-ASCII character at " & i
    Next i
    TestAssert.AssertTrue InStr(line, """auto"":true") > 0
End Sub

' --- file names -------------------------------------------------------------------

Public Sub Test_MonthFileName()
    TestAssert.AssertEqual "match_audit_2026-10.jsonl", AuditLogic.MonthFileName(DateSerial(2026, 10, 31))
    TestAssert.AssertEqual "match_audit_2027-01.jsonl", AuditLogic.MonthFileName(DateSerial(2027, 1, 1))
End Sub

Public Sub Test_IsExpiredFileName()
    Dim today As Date

    today = DateSerial(2026, 10, 7)
    TestAssert.AssertTrue AuditLogic.IsExpiredFileName("match_audit_2025-09.jsonl", today, 12), "13 months old"
    TestAssert.AssertFalse AuditLogic.IsExpiredFileName("match_audit_2025-10.jsonl", today, 12), "12 months old is kept"
    TestAssert.AssertFalse AuditLogic.IsExpiredFileName("match_audit_2026-10.jsonl", today, 12), "current month"
    TestAssert.AssertTrue AuditLogic.IsExpiredFileName("match_audit_2020-01.jsonl", today, 12)
    TestAssert.AssertFalse AuditLogic.IsExpiredFileName("notes.txt", today, 12), "foreign files are never deleted"
    TestAssert.AssertFalse AuditLogic.IsExpiredFileName("match_audit_2020-13.jsonl", today, 12), "invalid month"
    TestAssert.AssertFalse AuditLogic.IsExpiredFileName("match_audit_2020-01.jsonl.bak", today, 12)
    TestAssert.AssertTrue AuditLogic.IsExpiredFileName("MATCH_AUDIT_2020-01.JSONL", today, 12), "file systems ignore case"
End Sub

' --- file sink ----------------------------------------------------------------------

Private Function TempFolder() As String
    TempFolder = Environ$("TEMP") & "\ninja_audit_test_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd() * 100000))
End Function

Private Function ReadAll(ByVal path As String) As String
    Dim fso As Object
    Dim stream As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    Set stream = fso.OpenTextFile(path, 1)
    If Not stream.AtEndOfStream Then ReadAll = stream.ReadAll
    stream.Close
End Function

Private Sub RemoveFolder(ByVal path As String)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If fso.FolderExists(path) Then fso.DeleteFolder path, True
End Sub

Public Sub Test_FileSink_AppendsLinesAndCreatesFolders()
    Dim sink As FileAuditSink
    Dim base As String
    Dim text As String

    base = TempFolder() & "\deep\audit"
    Set sink = New FileAuditSink
    sink.BaseFolder = base
    sink.TestNow = DateSerial(2026, 10, 7) + TimeSerial(9, 0, 0)

    sink.AppendLine "line one"
    sink.AppendLine "line two"

    text = ReadAll(base & "\match_audit_2026-10.jsonl")
    TestAssert.AssertEqual "line one" & vbCrLf & "line two" & vbCrLf, text
    RemoveFolder Left$(base, InStr(base, "\deep") - 1)
End Sub

Public Sub Test_FileSink_NewFilePerMonth()
    Dim sink As FileAuditSink
    Dim base As String

    base = TempFolder()
    Set sink = New FileAuditSink
    sink.BaseFolder = base

    sink.TestNow = DateSerial(2026, 10, 31) + TimeSerial(23, 59, 0)
    sink.AppendLine "october"
    sink.TestNow = DateSerial(2026, 11, 1) + TimeSerial(0, 1, 0)
    sink.AppendLine "november"

    TestAssert.AssertEqual "october" & vbCrLf, ReadAll(base & "\match_audit_2026-10.jsonl")
    TestAssert.AssertEqual "november" & vbCrLf, ReadAll(base & "\match_audit_2026-11.jsonl")
    RemoveFolder base
End Sub

Public Sub Test_FileSink_PrunesOnlyExpiredAuditFiles()
    Dim sink As FileAuditSink
    Dim fso As Object
    Dim base As String

    base = TempFolder()
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder base
    fso.CreateTextFile(base & "\match_audit_2024-01.jsonl").Close
    fso.CreateTextFile(base & "\match_audit_2026-05.jsonl").Close
    fso.CreateTextFile(base & "\notes.txt").Close

    Set sink = New FileAuditSink
    sink.BaseFolder = base
    sink.TestNow = DateSerial(2026, 10, 7)
    sink.AppendLine "x"

    TestAssert.AssertFalse fso.FileExists(base & "\match_audit_2024-01.jsonl"), "old month is deleted"
    TestAssert.AssertTrue fso.FileExists(base & "\match_audit_2026-05.jsonl"), "recent month stays"
    TestAssert.AssertTrue fso.FileExists(base & "\notes.txt"), "other files stay"
    TestAssert.AssertTrue fso.FileExists(base & "\match_audit_2026-10.jsonl")
    RemoveFolder base
End Sub

' --- audit line -------------------------------------------------------------------

Public Sub Test_BuildLine_HasFeaturesButNoNamesOrBody()
    Dim mail As FakeMail
    Dim line As String

    Setup
    Set mail = Foreign("AW: [TicketNr. 900000101] Fehler beim Verbuchen")
    mail.BodyText = "Sehr geehrter Herr Mustermann, geheimer Inhalt Ticket 20101"

    line = TicketAudit.BuildLine(mail, "auto_decision", MatchFound, RuleExternalNumber, 20001, "log_only", "dry_run")

    TestAssert.AssertTrue InStr(line, """external"":""900000101""") > 0, line
    TestAssert.AssertTrue InStr(line, """subject"":""fehler beim verbuchen""") > 0, "normalized subject"
    TestAssert.AssertTrue InStr(line, """domain"":""vendor.example""") > 0
    TestAssert.AssertTrue InStr(line, """source"":""dry_run""") > 0
    TestAssert.AssertTrue InStr(line, """ninja_ids"":""20101""") > 0, "ids found in the body are features"
    TestAssert.AssertTrue InStr(line, """msg"":""0a1b2c3d""") > 0
    TestAssert.AssertEqual 0, InStr(line, "Mustermann"), "no person names"
    TestAssert.AssertEqual 0, InStr(line, "geheimer"), "no body text"
    TestAssert.AssertEqual 0, InStr(line, "conversation-id-123"), "the conversation id itself is not logged"
    TestAssert.AssertTrue InStr(line, """conv"":""" & AuditLogic.Hash32Hex("conversation-id-123") & """") > 0, "only its hash"
    Teardown
End Sub

' --- privacy: no subject for mails without a match ---------------------------------

Public Sub Test_Privacy_NoSubjectWithoutAMatch_OnlyItsHash()
    Dim mail As FakeMail
    Dim line As String

    Setup
    Set mail = Foreign("Krankmeldung Mustermann Max")

    line = TicketAudit.BuildLine(mail, "auto_decision", MatchNone, RuleNone, 0, "none", "watcher")
    TestAssert.AssertEqual 0, InStr(line, "Mustermann"), "no subject text: " & line
    TestAssert.AssertEqual 0, InStr(line, "krankmeldung"), line
    TestAssert.AssertTrue InStr(line, """subject"":""""") > 0, "subject field is empty"
    TestAssert.AssertTrue InStr(line, """subject_hash"":""" & AuditLogic.Hash32Hex("krankmeldung mustermann max") & """") > 0, "but the hash is there"
    Teardown
End Sub

Public Sub Test_Privacy_SubjectIsLoggedForFoundAmbiguousAndUserDecisions()
    Dim mail As FakeMail
    Dim line As String
    Dim h As String

    Setup
    Set mail = Foreign("AW: Example order total issue")
    h = """subject_hash"":""" & AuditLogic.Hash32Hex("example order total issue") & """"

    line = TicketAudit.BuildLine(mail, "auto_decision", MatchFound, RuleSubject, 30001, "log_only", "watcher")
    TestAssert.AssertTrue InStr(line, """subject"":""example order total issue""") > 0, line
    TestAssert.AssertTrue InStr(line, h) > 0

    line = TicketAudit.BuildLine(mail, "auto_decision", MatchAmbiguous, RuleSubject, 0, "none", "watcher")
    TestAssert.AssertTrue InStr(line, """subject"":""example order total issue""") > 0, line

    line = TicketAudit.BuildLine(mail, "user_rejected", MatchFound, RuleSubject, 30001, "rejected", "review")
    TestAssert.AssertTrue InStr(line, """event"":""user_rejected""") > 0
    TestAssert.AssertTrue InStr(line, """subject"":""example order total issue""") > 0, line
    Teardown
End Sub

Public Sub Test_Privacy_FlowWritesNoSubjectForNoMatch()
    Setup
    TicketService.MatchForeignMail Foreign("Neuer Mitarbeiter Mustermann Max"), pIndex, pStore, False

    TestAssert.AssertEqual 1, pSink.Count
    TestAssert.AssertEqual 0, InStr(pSink.LastLine(), "Mustermann"), pSink.LastLine()
    TestAssert.AssertEqual 0, InStr(TicketLog.Text(), "Mustermann"), "the text log neither: " & TicketLog.Text()
    TestAssert.AssertTrue InStr(TicketLog.Text(), "subject_hash=" & AuditLogic.Hash32Hex("neuer mitarbeiter mustermann max")) > 0, "same hash as in the audit log"
    TestAssert.AssertTrue InStr(pSink.LastLine(), "subject_hash"":""" & AuditLogic.Hash32Hex("neuer mitarbeiter mustermann max")) > 0
    Teardown
End Sub

Public Sub Test_BuildLine_EmptyIdsGiveEmptyFields()
    Dim mail As FakeMail
    Dim line As String

    Setup
    Set mail = Foreign("Hallo")
    mail.ConversationId = ""
    mail.MessageKey = ""

    line = TicketAudit.BuildLine(mail, "auto_decision", MatchNone, RuleNone, 0, "none", "watcher")
    TestAssert.AssertTrue InStr(line, """conv"":""""") > 0
    TestAssert.AssertTrue InStr(line, """msg"":""""") > 0
    TestAssert.AssertTrue InStr(line, """ticket"":0") > 0
    Teardown
End Sub

' --- matching flow -------------------------------------------------------------------

Public Sub Test_Flow_OneLinePerDecision()
    Setup
    ' own conversation per mail: the first match must not carry over to the second mail
    TicketService.MatchForeignMail ForeignIn("Ticket 20101", "c1"), pIndex, pStore, False, "dry_run"
    TicketService.MatchForeignMail ForeignIn("Hallo zusammen", "c2"), pIndex, pStore, False, "sweep"
    TicketService.MatchForeignMail ForeignIn("Ticket 20101 und Ticket 20102", "c3"), pIndex, pStore, False

    TestAssert.AssertEqual 3, pSink.Count
    TestAssert.AssertTrue InStr(pSink.Lines(1), """outcome"":""found""") > 0, pSink.Lines(1)
    TestAssert.AssertTrue InStr(pSink.Lines(1), """action"":""log_only""") > 0, pSink.Lines(1)
    TestAssert.AssertTrue InStr(pSink.Lines(1), """source"":""dry_run""") > 0, pSink.Lines(1)
    TestAssert.AssertTrue InStr(pSink.Lines(1), """ticket"":20101") > 0, pSink.Lines(1)
    TestAssert.AssertTrue InStr(pSink.Lines(2), """outcome"":""none""") > 0, pSink.Lines(2)
    TestAssert.AssertTrue InStr(pSink.Lines(2), """action"":""none""") > 0, pSink.Lines(2)
    TestAssert.AssertTrue InStr(pSink.Lines(2), """source"":""sweep""") > 0, pSink.Lines(2)
    TestAssert.AssertTrue InStr(pSink.Lines(3), """outcome"":""ambiguous""") > 0, pSink.Lines(3)
    TestAssert.AssertTrue InStr(pSink.Lines(3), """source"":""watcher""") > 0, pSink.Lines(3)
    Teardown
End Sub

Public Sub Test_Flow_ActionMovedAndStays()
    Dim mail As FakeMail

    Setup
    pIndex.AddTicket 20102, "Example closed ticket", False
    pTransport.AddResponse "ticketing/ticket/20102", 200, Fixtures.TicketClosed()

    TicketService.MatchForeignMail Foreign("Ticket 20101"), pIndex, pStore, True
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""moved""") > 0, pSink.LastLine()

    TicketService.MatchForeignMail Foreign("Ticket 20102"), pIndex, pStore, True
    TestAssert.AssertTrue InStr(pSink.LastLine(), """action"":""stays""") > 0, pSink.LastLine()
    Teardown
End Sub

Public Sub Test_Flow_TicketSenderIsNotAudited()
    Dim mail As FakeMail

    Setup
    Set mail = Foreign("Ticket 20101")
    mail.SenderName = TicketConfig.AllowedSender()

    TicketService.MatchForeignMail mail, pIndex, pStore, False
    TestAssert.AssertEqual 0, pSink.Count
    Teardown
End Sub

Public Sub Test_Flow_BrokenAuditLogNeverStopsTheMatching()
    Dim mail As FakeMail

    Setup
    pSink.FailOnAppend = True
    Set mail = Foreign("Ticket 20101")

    TestAssert.AssertEqual MatchFound, TicketService.MatchForeignMail(mail, pIndex, pStore, True)
    TestAssert.AssertTrue Not mail.MovedTo Is Nothing, "the mail is moved although the audit write failed"
    TicketService.MatchForeignMail Foreign("Ticket 20101"), pIndex, pStore, True
    TestAssert.AssertEqual 1, CountOccurrences(TicketLog.Text(), "Audit log: could not write"), "the problem is logged once"
    Teardown
End Sub

Private Function CountOccurrences(ByVal text As String, ByVal part As String) As Long
    Dim pos As Long

    pos = InStr(text, part)
    Do While pos > 0
        CountOccurrences = CountOccurrences + 1
        pos = InStr(pos + Len(part), text, part)
    Loop
End Function

Public Sub Test_Flow_NoFileIsWrittenWhileFileLoggingIsDisabled()
    Dim mail As FakeMail

    Setup
    TicketAudit.ResetState            ' no sink injected: the safety net must keep the real file out
    Set mail = Foreign("Ticket 20101")

    TestAssert.AssertEqual MatchFound, TicketService.MatchForeignMail(mail, pIndex, pStore, False)
    TestAssert.AssertEqual 0, pSink.Count
    Teardown
End Sub
