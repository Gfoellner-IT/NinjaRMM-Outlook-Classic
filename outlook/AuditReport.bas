Attribute VB_Name = "AuditReport"
' AuditReport v0.1.0
'
' Evaluates the audit log (M7): how well each matching rule did according to the decisions of the
' user, where proposals were rejected, which patterns of confirmed mails could become new rules, and
' a simple check of those patterns against later decisions (the oldest REPORT_TRAIN_PERCENT of the
' decided mails find the patterns, the rest checks them). The result is a Markdown text; nothing is
' changed anywhere. No Outlook types.
'
' Privacy: the report contains counts, domains, ticket numbers and subject stems of mails the user
' decided on (they are about tickets); never sender names, mail texts or mails without a match.

Option Explicit

Private Const RULE_NAMES As String = "NinjaId,ExternalNumber,Conversation,Subject,Learned"
Private Const TOP_DOMAINS As Long = 10

' ---------------------------------------------------------------------------
'  Reading the log
' ---------------------------------------------------------------------------

' The records (Dictionary per line) of a text with one JSON object per line; broken lines are skipped.
Public Function ParseLines(ByVal text As String) As Collection
    Dim lines() As String
    Dim i As Long
    Dim rec As Object

    Set ParseLines = New Collection
    text = Replace(text, vbCrLf, vbLf)
    lines = Split(text, vbLf)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            Set rec = ParseLine(lines(i))
            If Not rec Is Nothing Then ParseLines.Add rec
        End If
    Next i
End Function

Private Function ParseLine(ByVal line As String) As Object
    Dim rec As Variant

    If Not JsonLite.TryParse(line, rec) Then Exit Function
    If TypeName(rec) <> "Dictionary" Then Exit Function
    Set ParseLine = rec
End Function

' All records of the monthly audit files in folder that are not older than monthsBack months, oldest first.
Public Function LoadRecords(ByVal folder As String, ByVal monthsBack As Long, ByVal today As Date) As Collection
    Dim fso As Object
    Dim names() As String
    Dim count As Long
    Dim f As Object
    Dim i As Long
    Dim j As Long
    Dim tmp As String
    Dim stream As Object
    Dim text As String
    Dim rec As Variant

    Set LoadRecords = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Exit Function

    ReDim names(0 To 0)
    For Each f In fso.GetFolder(folder).Files
        If LCase$(f.Name) Like "match_audit_####-##.jsonl" Then
            If Not AuditLogic.IsExpiredFileName(f.Name, today, monthsBack) Then
                ReDim Preserve names(0 To count)
                names(count) = f.Name
                count = count + 1
            End If
        End If
    Next f
    If count = 0 Then Exit Function

    For i = 1 To count - 1                       ' file names sort chronologically
        For j = 0 To count - 1 - i
            If LCase$(names(j)) > LCase$(names(j + 1)) Then
                tmp = names(j): names(j) = names(j + 1): names(j + 1) = tmp
            End If
        Next j
    Next i

    For i = 0 To count - 1
        Set stream = fso.OpenTextFile(folder & "\" & names(i), 1, False, 0)
        text = ""
        If Not stream.AtEndOfStream Then text = stream.ReadAll
        stream.Close
        For Each rec In ParseLines(text)
            LoadRecords.Add rec
        Next rec
    Next i
End Function

' ---------------------------------------------------------------------------
'  Small pure helpers
' ---------------------------------------------------------------------------

' The stem of a normalized subject: digit runs become "#", the first words are kept.
' "item235 - error 4711 in form220" -> "item# - error # in"
Public Function SubjectStem(ByVal normalizedSubject As String, Optional ByVal maxWords As Long = 4) As String
    SubjectStem = TicketMatcher.SubjectStem(normalizedSubject, maxWords)
End Function

' What the numbers say about the level of a rule.
Public Function RecommendLevel(ByVal decisions As Long, ByVal percent As Long, ByVal currentLevel As Long) As String
    If decisions = 0 Then
        RecommendLevel = GermanText.De("noch keine R[ue]ckmeldung")
    ElseIf decisions >= TicketConfig.REPORT_LEVEL3_MIN_DECISIONS And percent >= TicketConfig.REPORT_LEVEL3_MIN_PERCENT Then
        If currentLevel < 3 Then
            RecommendLevel = GermanText.De("Stufe 3 erw[ae]gen")
        Else
            RecommendLevel = GermanText.De("l[ae]uft gut, Stufe bleibt")
        End If
    ElseIf decisions >= TicketConfig.REPORT_LOWER_MIN_DECISIONS And percent < TicketConfig.REPORT_LOWER_BELOW_PERCENT Then
        RecommendLevel = GermanText.De("Stufe senken oder Regel versch[ae]rfen")
    Else
        RecommendLevel = GermanText.De("weiter beobachten")
    End If
End Function

' "ReportFolder": the configured folder or %LOCALAPPDATA%\NinjaRMM-Outlook-Classic\reports.
Public Function ReportFolder() As String
    If Len(TicketConfig.REPORT_FOLDER) > 0 Then
        ReportFolder = TicketConfig.REPORT_FOLDER
    Else
        ReportFolder = Environ$("LOCALAPPDATA") & "\NinjaRMM-Outlook-Classic\reports"
    End If
End Function

' "match_review_2026-10-09.md"
Public Function ReportFileName(ByVal today As Date) As String
    ReportFileName = "match_review_" & Format$(today, "yyyy-mm-dd") & ".md"
End Function

' The sentence that starts the conversation about the report.
Public Function ReviewPromptText(ByVal reportPath As String) As String
    ReviewPromptText = GermanText.De("Pr[ue]fen wir gemeinsam die aktuellen Regeln und Entscheidungen, mit dem Ziel, neue Regeln zu erstellen. Bericht: ") & reportPath
End Function

' Writes the report as UTF-8 into folder (created if needed) and returns the file path.
Public Function WriteReport(ByVal reportText As String, ByVal folder As String, ByVal today As Date) As String
    Dim fso As Object
    Dim stream As Object
    Dim path As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolder fso, folder
    path = folder & "\" & ReportFileName(today)

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText reportText
    stream.SaveToFile path, 2       ' adSaveCreateOverWrite
    stream.Close
    WriteReport = path
End Function

' ---------------------------------------------------------------------------
'  The report
' ---------------------------------------------------------------------------

Public Function BuildReport(ByVal records As Collection, ByVal generatedAt As Date) As String
    Set records = WithoutRepeats(records)
    Dim c As Object                 ' counters, key -> Long
    Dim cand As Object              ' domain|stem -> True (insertion order)
    Dim candTickets As Object       ' key -> Dictionary of ticket ids
    Dim candConvs As Object         ' key -> Dictionary of conversation hashes
    Dim noneDomains As Object
    Dim missedDomains As Object
    Dim decided As Collection
    Dim rec As Object
    Dim ev As String, rule As String, action As String, outcome As String, domain As String
    Dim firstTs As String, lastTs As String, ts As String
    Dim key As String
    Dim i As Long, n As Long, trainCount As Long
    Dim setName As String
    Dim s As String

    Set c = NewDict()
    Set cand = NewDict()
    Set candTickets = NewDict()
    Set candConvs = NewDict()
    Set noneDomains = NewDict()
    Set missedDomains = NewDict()
    Set decided = New Collection

    For Each rec In records
        ev = JsonLite.GetString(rec, "event")
        rule = JsonLite.GetString(rec, "rule")
        action = JsonLite.GetString(rec, "action")
        outcome = JsonLite.GetString(rec, "outcome")
        domain = JsonLite.GetString(rec, "domain")
        ts = JsonLite.GetString(rec, "ts")

        If Len(ts) > 0 Then
            If Len(firstTs) = 0 Or ts < firstTs Then firstTs = ts
            If ts > lastTs Then lastTs = ts
        End If
        Inc c, "ev|" & ev
        Inc c, "src|" & JsonLite.GetString(rec, "source")

        Select Case ev
            Case "auto_decision"
                If outcome = "found" Then Inc c, "rf|" & rule
                If action = "proposed" Then Inc c, "rp|" & rule
                If action = "moved" Then Inc c, "rm|" & rule
                If outcome = "none" Then
                    Inc c, "none"
                    If Len(domain) > 0 Then Inc noneDomains, domain
                End If
            Case "user_confirmed"
                Inc c, "rc|" & rule
                Inc c, "rdc|" & rule & "|" & domain
                decided.Add rec
            Case "manual_assignment"
                Inc c, "ma|" & action
                If action = "missed" And Len(domain) > 0 Then Inc missedDomains, domain
                decided.Add rec
            Case "user_rejected"
                Inc c, "rr|" & rule
                Inc c, "rdr|" & rule & "|" & domain
                decided.Add rec
        End Select
    Next rec

    ' patterns: (domain, subject stem) of decided mails, with a time split for the check
    n = decided.Count
    trainCount = (n * TicketConfig.REPORT_TRAIN_PERCENT) \ 100
    If trainCount < 1 And n > 0 Then trainCount = n
    i = 0
    For Each rec In decided
        i = i + 1
        If i <= trainCount Then setName = "train" Else setName = "test"
        key = PatternKey(rec)
        If Len(key) > 0 Then
            If Not cand.Exists(key) Then
                cand.Add key, True
                candTickets.Add key, NewDict()
                candConvs.Add key, NewDict()
            End If
            If JsonLite.GetString(rec, "event") <> "user_rejected" Then
                Inc c, "pc|all|" & key
                Inc c, "pc|" & setName & "|" & key
                AddToSet candTickets(key), CStr(TicketOf(rec))
                If Len(JsonLite.GetString(rec, "conv")) > 0 Then AddToSet candConvs(key), JsonLite.GetString(rec, "conv")
            Else
                Inc c, "pr|all|" & key
                Inc c, "pr|" & setName & "|" & key
            End If
        End If
    Next rec

    s = "# " & GermanText.De("Bericht zu den Zuordnungsregeln") & vbLf & vbLf
    s = s & GermanText.De("Erzeugt am ") & Format$(generatedAt, "yyyy-mm-dd hh:nn") & GermanText.De(". Nur lesende Auswertung des Audit-Logs, nichts wurde ge[ae]ndert.") & vbLf & vbLf
    s = s & SectionOverview(c, records.Count, firstTs, lastTs)
    s = s & SectionRules(c)
    s = s & SectionManual(c, missedDomains)
    s = s & SectionHotspots(c)
    s = s & SectionCandidates(c, cand, candTickets, candConvs, n, trainCount)
    s = s & SectionNone(c, noneDomains)
    s = s & SectionNotes()
    BuildReport = s
End Function

' The same automatic decision about the same mail is logged again by every sweep and dry run; the report counts it
' once (same message key, rule, outcome and action).
Private Function WithoutRepeats(ByVal records As Collection) As Collection
    Dim result As Collection
    Dim seen As Object
    Dim rec As Object
    Dim msg As String
    Dim key As String

    Set result = New Collection
    Set seen = NewDict()
    For Each rec In records
        msg = JsonLite.GetString(rec, "msg")
        If JsonLite.GetString(rec, "event") = "auto_decision" And Len(msg) > 0 Then
            key = msg & "|" & JsonLite.GetString(rec, "rule") & "|" & JsonLite.GetString(rec, "outcome") & "|" & JsonLite.GetString(rec, "action")
            If Not seen.Exists(key) Then
                seen.Add key, True
                result.Add rec
            End If
        Else
            result.Add rec
        End If
    Next rec
    Set WithoutRepeats = result
End Function

' ---------------------------------------------------------------------------
'  Report sections
' ---------------------------------------------------------------------------

Private Function SectionOverview(ByVal c As Object, ByVal total As Long, ByVal firstTs As String, ByVal lastTs As String) As String
    Dim s As String

    s = "## " & GermanText.De("[Ue]berblick") & vbLf & vbLf
    If total = 0 Then
        SectionOverview = s & GermanText.De("Das Audit-Log enth[ae]lt noch keine Eintr[ae]ge.") & vbLf & vbLf
        Exit Function
    End If
    s = s & "- " & GermanText.De("Zeitraum: ") & Left$(firstTs, 10) & " bis " & Left$(lastTs, 10) & vbLf
    s = s & "- " & GermanText.De("Eintr[ae]ge: ") & total & " (automatische Entscheidungen " & Num(c, "ev|auto_decision") & _
        GermanText.De(", best[ae]tigt ") & Num(c, "ev|user_confirmed") & ", abgelehnt " & Num(c, "ev|user_rejected") & ", manuell " & Num(c, "ev|manual_assignment") & ")" & vbLf
    s = s & "- Quellen: Watcher " & Num(c, "src|watcher") & ", Durchlauf " & Num(c, "src|sweep") & ", Probelauf " & Num(c, "src|dry_run") & _
        ", Pr" & ChrW$(252) & "fung " & Num(c, "src|review") & vbLf & vbLf
    SectionOverview = s
End Function

Private Function SectionRules(ByVal c As Object) As String
    Dim names() As String
    Dim i As Long
    Dim rule As String
    Dim conf As Long, rej As Long, decisions As Long, pct As Long
    Dim s As String
    Dim level As Long

    names = Split(RULE_NAMES, ",")
    s = "## Regeln" & vbLf & vbLf & GermanText.De("Trefferquote = best[ae]tigt / (best[ae]tigt + abgelehnt). Stufe 3 liefert keine R[ue]ckmeldung (ohne Nachfrage verschoben).") & vbLf & vbLf
    s = s & "| Regel | Stufe heute | Treffer | Vorschl" & ChrW$(228) & "ge | Automatisch verschoben | " & GermanText.De("Best[ae]tigt") & " | Abgelehnt | Trefferquote | Empfehlung |" & vbLf
    s = s & "|---|---|---|---|---|---|---|---|---|" & vbLf

    For i = LBound(names) To UBound(names)
        rule = names(i)
        conf = Num(c, "rc|" & rule)
        rej = Num(c, "rr|" & rule)
        decisions = conf + rej
        level = MatchLevels.LevelFor(RuleFromName(rule))
        If decisions > 0 Then pct = (conf * 100) \ decisions Else pct = 0
        s = s & "| " & rule & " | " & level & " | " & Num(c, "rf|" & rule) & " | " & Num(c, "rp|" & rule) & " | " & Num(c, "rm|" & rule) & _
            " | " & conf & " | " & rej & " | " & IIf(decisions > 0, CStr(pct) & " %", "-") & " | " & RecommendLevel(decisions, pct, level) & " |" & vbLf
    Next i
    SectionRules = s & vbLf
End Function

' Mails the user moved into ticket folders by hand, and what the matcher would have done with them.
Private Function SectionManual(ByVal c As Object, ByVal missedDomains As Object) As String
    Dim s As String
    Dim total As Long
    Dim agree As Long
    Dim keys() As String
    Dim counts() As Long
    Dim n As Long, i As Long, j As Long, shown As Long
    Dim k As Variant
    Dim tk As String
    Dim tc As Long

    total = Num(c, "ev|manual_assignment")
    agree = Num(c, "ma|agree")

    s = "## Manuelle Zuordnungen" & vbLf & vbLf
    If total = 0 Then
        SectionManual = s & GermanText.De("Noch keine von Hand in Ticketordner verschobenen Mails erkannt.") & vbLf & vbLf
        Exit Function
    End If

    s = s & total & GermanText.De(" Mails wurden von Hand in einen Ticketordner verschoben. Der Matcher h[ae]tte davon ") & agree & " (" & (agree * 100) \ total & " %) " & _
        GermanText.De("dem gleichen Ticket zugeordnet.") & vbLf & vbLf
    s = s & "| Ergebnis des Matchers | Anzahl |" & vbLf & "|---|---|" & vbLf
    s = s & "| " & GermanText.De("gleiches Ticket (agree)") & " | " & agree & " |" & vbLf
    s = s & "| anderes Ticket (differ) | " & Num(c, "ma|differ") & " |" & vbLf
    s = s & "| verpasst (missed) | " & Num(c, "ma|missed") & " |" & vbLf
    s = s & "| nie zugeordnet (excluded) | " & Num(c, "ma|excluded") & " |" & vbLf
    s = s & "| nicht gepr" & ChrW$(252) & "ft (unchecked) | " & Num(c, "ma|unchecked") & " |" & vbLf & vbLf

    n = missedDomains.Count
    If n > 0 Then
        ReDim keys(0 To n - 1)
        ReDim counts(0 To n - 1)
        i = 0
        For Each k In missedDomains.Keys
            keys(i) = CStr(k)
            counts(i) = missedDomains(k)
            i = i + 1
        Next k
        For i = 0 To n - 2
            For j = 0 To n - 2 - i
                If counts(j) < counts(j + 1) Then
                    tc = counts(j): counts(j) = counts(j + 1): counts(j + 1) = tc
                    tk = keys(j): keys(j) = keys(j + 1): keys(j + 1) = tk
                End If
            Next j
        Next i
        s = s & GermanText.De("Verpasste Zuordnungen nach Absenderdomain (hier lohnen sich neue Regeln):") & vbLf & vbLf
        For i = 0 To n - 1
            If shown >= TOP_DOMAINS Then Exit For
            s = s & "- " & keys(i) & ": " & counts(i) & vbLf
            shown = shown + 1
        Next i
        s = s & vbLf
    End If
    SectionManual = s
End Function

' (rule, domain) pairs with many rejections and a poor rate: candidates for the exclusion list or a stricter rule.
Private Function SectionHotspots(ByVal c As Object) As String
    Dim k As Variant
    Dim parts() As String
    Dim conf As Long, rej As Long
    Dim s As String
    Dim hasAny As Boolean

    s = "## " & GermanText.De("Auff[ae]llige Absender (viele Ablehnungen)") & vbLf & vbLf
    For Each k In c.Keys
        If Left$(CStr(k), 4) = "rdr|" Then
            parts = Split(CStr(k), "|")
            If UBound(parts) >= 2 Then
                rej = c(k)
                conf = Num(c, "rdc|" & parts(1) & "|" & parts(2))
                If rej >= 3 And (conf * 100) \ (conf + rej) < 70 Then
                    s = s & "- " & parts(2) & " (Regel " & parts(1) & "): " & rej & " abgelehnt, " & conf & GermanText.De(" best[ae]tigt. Kandidat f[ue]r die Ausschlussliste oder eine strengere Regel.") & vbLf
                    hasAny = True
                End If
            End If
        End If
    Next k
    If Not hasAny Then s = s & GermanText.De("Keine.") & vbLf
    SectionHotspots = s & vbLf
End Function

Private Function SectionCandidates(ByVal c As Object, ByVal cand As Object, ByVal candTickets As Object, ByVal candConvs As Object, _
                                   ByVal decidedCount As Long, ByVal trainCount As Long) As String
    Dim keys() As String
    Dim counts() As Long
    Dim n As Long
    Dim k As Variant
    Dim i As Long, j As Long
    Dim tk As String
    Dim tc As Long
    Dim parts() As String
    Dim s As String
    Dim support As Long, rej As Long, testConf As Long, testRej As Long, trainConf As Long
    Dim note As String
    Dim shown As Long
    Dim snippets As String

    s = "## Kandidaten f" & ChrW$(252) & "r neue Regeln" & vbLf & vbLf
    s = s & GermanText.De("Muster aus best[ae]tigten Mails: gleiche Absenderdomain und gleicher Betreffstamm (Zahlen als # ersetzt). Ein Muster gilt ab ") & _
        TicketConfig.REPORT_MIN_SUPPORT & GermanText.De(" Best[ae]tigungen ohne Ablehnung in mindestens 2 Konversationen als Kandidat. ") & _
        GermanText.De("Pr[ue]fung: die ersten ") & TicketConfig.REPORT_TRAIN_PERCENT & GermanText.De(" % der Entscheidungen (") & trainCount & " von " & decidedCount & _
        GermanText.De(") liefern das Muster, die restlichen pr[ue]fen es.") & vbLf & vbLf

    n = cand.Count
    If n > 0 Then
        ReDim keys(0 To n - 1)
        ReDim counts(0 To n - 1)
        i = 0
        For Each k In cand.Keys
            keys(i) = CStr(k)
            counts(i) = Num(c, "pc|all|" & CStr(k))
            i = i + 1
        Next k
        For i = 0 To n - 2                           ' sort by confirmations, descending
            For j = 0 To n - 2 - i
                If counts(j) < counts(j + 1) Then
                    tc = counts(j): counts(j) = counts(j + 1): counts(j + 1) = tc
                    tk = keys(j): keys(j) = keys(j + 1): keys(j + 1) = tk
                End If
            Next j
        Next i
    End If

    s = s & "| Domain | Betreffstamm | " & GermanText.De("Best[ae]tigt") & " | Tickets | Konversationen | Training | Test | Hinweis |" & vbLf
    s = s & "|---|---|---|---|---|---|---|---|" & vbLf

    For i = 0 To n - 1
        support = Num(c, "pc|all|" & keys(i))
        rej = Num(c, "pr|all|" & keys(i))
        If support >= TicketConfig.REPORT_MIN_SUPPORT And rej = 0 And candConvs(keys(i)).Count >= 2 Then
            parts = Split(keys(i), "|")
            trainConf = Num(c, "pc|train|" & keys(i))
            testConf = Num(c, "pc|test|" & keys(i))
            testRej = Num(c, "pr|test|" & keys(i))
            If testConf + testRej = 0 Then
                note = "keine Testdaten"
            ElseIf testRej = 0 Then
                note = GermanText.De("stabil, pr[ue]fen")
            Else
                note = "Test mit Ablehnung"
            End If
            s = s & "| " & parts(0) & " | " & parts(1) & " | " & support & " | " & TicketList(candTickets(keys(i))) & " | " & candConvs(keys(i)).Count & _
                " | " & trainConf & " | " & testConf & " ok / " & testRej & " abgelehnt | " & note & " |" & vbLf
            shown = shown + 1
            snippets = snippets & "{""id"":""" & RuleId(parts(0), parts(1)) & """,""domain"":""" & parts(0) & """,""stem"":""" & parts(1) & """,""enabled"":false,""note"":""Kandidat vom Bericht""}" & vbLf
        End If
    Next i
    If shown = 0 Then s = s & "| - | - | - | - | - | - | - | " & GermanText.De("noch keine Kandidaten") & " |" & vbLf
    s = s & vbLf & GermanText.De("Eine Regel aus einem Kandidaten beschreibt die Suche (Domain und Betreffstamm: das einzige offene Ticket mit diesem Stamm im Titel), nie feste Ticketnummern. Neue Regeln starten auf Stufe 1.") & vbLf & vbLf
    If Len(snippets) > 0 Then
        s = s & GermanText.De("Vorschl[ae]ge f[ue]r die Regeldatei (match_rules.json, bisher ausgeschaltet; zum Aktivieren enabled auf true setzen):") & vbLf & vbLf & "```json" & vbLf & snippets & "```" & vbLf & vbLf
    End If
    SectionCandidates = s
End Function

Private Function SectionNone(ByVal c As Object, ByVal noneDomains As Object) As String
    Dim s As String
    Dim keys() As String
    Dim counts() As Long
    Dim n As Long, i As Long, j As Long
    Dim k As Variant
    Dim tk As String
    Dim tc As Long
    Dim shown As Long

    s = "## Mails ohne Treffer" & vbLf & vbLf
    s = s & Num(c, "none") & GermanText.De(" Entscheidungen ohne Treffer. Von diesen Mails liegt nur der Betreff-Hash im Log, sie liefern keine Muster, solange sie nicht von Hand einem Ticket zugeordnet werden. H[ae]ufigste Absenderdomains:") & vbLf & vbLf

    n = noneDomains.Count
    If n = 0 Then
        SectionNone = s & "-" & vbLf & vbLf
        Exit Function
    End If
    ReDim keys(0 To n - 1)
    ReDim counts(0 To n - 1)
    i = 0
    For Each k In noneDomains.Keys
        keys(i) = CStr(k)
        counts(i) = noneDomains(k)
        i = i + 1
    Next k
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If counts(j) < counts(j + 1) Then
                tc = counts(j): counts(j) = counts(j + 1): counts(j + 1) = tc
                tk = keys(j): keys(j) = keys(j + 1): keys(j + 1) = tk
            End If
        Next j
    Next i
    For i = 0 To n - 1
        If shown >= TOP_DOMAINS Then Exit For
        s = s & "- " & keys(i) & ": " & counts(i) & vbLf
        shown = shown + 1
    Next i
    SectionNone = s & vbLf
End Function

Private Function SectionNotes() As String
    SectionNotes = "## Hinweise" & vbLf & vbLf & _
        "- " & GermanText.De("Der Bericht enth[ae]lt Zahlen, Domains, Ticketnummern und Betreffst[ae]mme von Mails mit Ticketbezug, keine Absendernamen und keine Mailtexte.") & vbLf & _
        "- " & GermanText.De("Schwellenwerte stehen in TicketConfig (REPORT_*).") & vbLf & _
        "- " & GermanText.De("Nichts wird automatisch aktiv: Regeln und Stufen [ae]ndern wir gemeinsam und mit Tests.") & vbLf
End Function

' ---------------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------------

Private Sub AddToSet(ByVal d As Object, ByVal key As String)
    If Not d.Exists(key) Then d.Add key, True
End Sub

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Private Sub Inc(ByVal d As Object, ByVal key As String)
    If d.Exists(key) Then d(key) = d(key) + 1 Else d.Add key, 1
End Sub

Private Function Num(ByVal d As Object, ByVal key As String) As Long
    If d.Exists(key) Then Num = d(key)
End Function

Private Function PatternKey(ByVal rec As Object) As String
    Dim domain As String
    Dim stem As String

    domain = JsonLite.GetString(rec, "domain")
    stem = SubjectStem(JsonLite.GetString(rec, "subject"))
    If Len(domain) = 0 Or Len(stem) = 0 Then Exit Function
    PatternKey = domain & "|" & stem
End Function

Private Function TicketOf(ByVal rec As Object) As Long
    Dim v As Double

    If JsonLite.TryGetNumber(rec, "ticket", v) Then TicketOf = CLng(v)
End Function

Private Function TicketList(ByVal tickets As Object) As String
    Dim k As Variant
    Dim shown As Long

    For Each k In tickets.Keys
        If shown >= 5 Then
            TicketList = TicketList & ", ..."
            Exit For
        End If
        If shown > 0 Then TicketList = TicketList & ", "
        TicketList = TicketList & "#" & CStr(k)
        shown = shown + 1
    Next k
End Function

' "vendor.example" and "item# - error #" give "vendor-example-item-error"
Private Function RuleId(ByVal domain As String, ByVal stem As String) As String
    Dim i As Long
    Dim c As String
    Dim s As String
    Dim lastDash As Boolean

    For i = 1 To Len(domain & " " & stem)
        c = LCase$(Mid$(domain & " " & stem, i, 1))
        If c Like "[a-z0-9]" Then
            s = s & c
            lastDash = False
        ElseIf Not lastDash And Len(s) > 0 Then
            s = s & "-"
            lastDash = True
        End If
    Next i
    If Right$(s, 1) = "-" Then s = Left$(s, Len(s) - 1)
    RuleId = Left$(s, 48)
End Function

Private Function RuleFromName(ByVal name As String) As MatchRule
    Select Case name
        Case "NinjaId":         RuleFromName = RuleNinjaId
        Case "ExternalNumber":  RuleFromName = RuleExternalNumber
        Case "Conversation":    RuleFromName = RuleConversation
        Case "Subject":         RuleFromName = RuleSubject
        Case "Learned":         RuleFromName = RuleLearned
    End Select
End Function

Private Sub EnsureFolder(ByVal fso As Object, ByVal path As String)
    Dim parent As String

    If fso.FolderExists(path) Then Exit Sub
    parent = fso.GetParentFolderName(path)
    If Len(parent) > 0 Then EnsureFolder fso, parent
    fso.CreateFolder path
End Sub
