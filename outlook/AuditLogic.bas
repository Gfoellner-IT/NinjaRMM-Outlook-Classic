Attribute VB_Name = "AuditLogic"
' AuditLogic v0.1.0
'
' Pure functions for the match audit log: a stable short hash, JSON string escaping, one
' JSON line per match decision and the monthly file name. No Outlook types, no file access.
' The output is pure ASCII (non-ASCII characters are written as \uXXXX), so the log file can be
' appended with any text writer.
'
' Only features are logged, never mail contents or person names (see ROADMAP M7).

Option Explicit

' ---------------------------------------------------------------------------
'  Hash
' ---------------------------------------------------------------------------

' FNV-1a (32 bit) of the text as 8 lower-case hex digits. Characters above 255 contribute their
' low and high byte. Used to log conversation and message ids without the ids themselves.
' Computed with Double arithmetic because VBA has no unsigned 32-bit integer.
Public Function Hash32Hex(ByVal text As String) As String
    Dim h As Double
    Dim i As Long
    Dim c As Long

    h = 2166136261#
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1))
        If c < 0 Then c = c + 65536
        h = HashStep(h, c And 255)
        If c > 255 Then h = HashStep(h, c \ 256)
    Next i

    Hash32Hex = Hex2(Int(h / 16777216#)) & Hex2(Int(h / 65536#) Mod 256) & _
                Hex2(Int(h / 256#) Mod 256) & Hex2(h - 256# * Int(h / 256#))
End Function

' One FNV-1a step: xor the low byte with b, then multiply by 16777619 (= 2^24 + 403) mod 2^32.
Private Function HashStep(ByVal h As Double, ByVal b As Long) As Double
    Dim low As Long
    Dim t As Double

    low = CLng(h - 256# * Int(h / 256#))
    h = h - low + (low Xor b)

    low = CLng(h - 256# * Int(h / 256#))
    t = low * 16777216# + h * 403#
    HashStep = t - 4294967296# * Int(t / 4294967296#)
End Function

Private Function Hex2(ByVal value As Double) As String
    Hex2 = LCase$(Right$("0" & Hex$(CLng(value)), 2))
End Function

' ---------------------------------------------------------------------------
'  JSON
' ---------------------------------------------------------------------------

' Text as a quoted JSON string; control characters and non-ASCII characters are escaped.
Public Function JsonString(ByVal text As String) As String
    Dim i As Long
    Dim c As Long
    Dim result As String

    result = """"
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1))
        If c < 0 Then c = c + 65536
        Select Case c
            Case 34:  result = result & "\"""
            Case 92:  result = result & "\\"
            Case 10:  result = result & "\n"
            Case 13:  result = result & "\r"
            Case 9:   result = result & "\t"
            Case 32 To 126
                result = result & Chr$(c)
            Case Else
                result = result & "\u" & Right$("0000" & LCase$(Hex$(c)), 4)
        End Select
    Next i
    JsonString = result & """"
End Function

' One audit line (without line break) for a match decision or a decision of the user.
'   ts            time of the decision (local time)
'   eventType     auto_decision | user_confirmed | user_rejected | manual_assignment
'   source        watcher | sweep | dry_run
'   outcome       found | ambiguous | none
'   rule          rule name or "None"
'   action        log_only | proposed | moved | stays | none | rejected
'   subject       normalized subject, "" if it is not to be logged; subjectHash is its hash (always)
'   matcherTicket for manual_assignment: the ticket the matcher would have chosen (0 = none)
Public Function BuildMatchEvent(ByVal when As Date, ByVal eventType As String, ByVal source As String, ByVal outcome As String, ByVal rule As String, _
                                ByVal ticketId As Long, ByVal action As String, ByVal senderDomain As String, _
                                ByVal normalizedSubject As String, ByVal subjectHash As String, ByVal externalNumbers As String, ByVal ninjaIds As String, _
                                ByVal conversationHash As String, ByVal messageKey As String, ByVal autoHeader As Boolean, _
                                Optional ByVal matcherTicket As Long = 0) As String
    Dim s As String

    s = "{""ts"":" & JsonString(Format$(when, "yyyy-mm-dd\Thh:nn:ss"))
    s = s & ",""event"":" & JsonString(eventType)
    s = s & ",""source"":" & JsonString(source)
    s = s & ",""outcome"":" & JsonString(outcome)
    s = s & ",""rule"":" & JsonString(rule)
    s = s & ",""ticket"":" & CStr(ticketId)
    s = s & ",""action"":" & JsonString(action)
    s = s & ",""domain"":" & JsonString(senderDomain)
    s = s & ",""subject"":" & JsonString(normalizedSubject)
    s = s & ",""subject_hash"":" & JsonString(subjectHash)
    s = s & ",""external"":" & JsonString(externalNumbers)
    s = s & ",""ninja_ids"":" & JsonString(ninjaIds)
    s = s & ",""conv"":" & JsonString(conversationHash)
    s = s & ",""msg"":" & JsonString(messageKey)
    s = s & ",""auto"":" & IIf(autoHeader, "true", "false")
    If matcherTicket > 0 Then s = s & ",""matcher_ticket"":" & CStr(matcherTicket)
    BuildMatchEvent = s & "}"
End Function

' ---------------------------------------------------------------------------
'  Files
' ---------------------------------------------------------------------------

' "match_audit_2026-10.jsonl"
Public Function MonthFileName(ByVal when As Date) As String
    MonthFileName = "match_audit_" & Format$(when, "yyyy-mm") & ".jsonl"
End Function

' True if fileName is an audit file of a month that lies more than retentionMonths before today.
' Other files are never touched (returns False).
Public Function IsExpiredFileName(ByVal fileName As String, ByVal today As Date, ByVal retentionMonths As Long) As Boolean
    Dim year As Long
    Dim month As Long
    Dim fileMonths As Long
    Dim nowMonths As Long

    If Not (LCase$(fileName) Like "match_audit_####-##.jsonl") Then Exit Function

    year = CLng(Mid$(fileName, 13, 4))
    month = CLng(Mid$(fileName, 18, 2))
    If month < 1 Or month > 12 Then Exit Function

    fileMonths = year * 12 + month
    nowMonths = Format$(today, "yyyy") * 12 + Format$(today, "mm")
    IsExpiredFileName = (nowMonths - fileMonths > retentionMonths)
End Function
