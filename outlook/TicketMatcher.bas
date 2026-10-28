Attribute VB_Name = "TicketMatcher"
' TicketMatcher v0.2.0
'
' Pure logic: which existing NinjaOne ticket does a foreign mail belong to? No Outlook
' types, candidates come from ITicketIndex. Rules, first decisive one wins:
'
'   1 NinjaId        a NinjaOne id in subject or text ("Ticket 20108", "(#20108)"), must exist
'   2 ExternalNumber vendor ticket number in the subject ("[TicketNr. 900000105]") found in
'                    an internal ticket. Found = unambiguous. Not found = the mail STAYS in
'                    the inbox, no other rule is tried.
'   3 Conversation   another mail of the same conversation is already assigned
'   4 Subject        normalized subject equals the summary of exactly one open ticket
'
' When in doubt (no hit, several hits, excluded mail) the result is MatchNone or
' MatchAmbiguous and the mail stays in the inbox.

Option Explicit

Public Enum MatchOutcome
    MatchNone = 0
    MatchFound = 1
    MatchAmbiguous = 2
End Enum

Public Enum MatchRule
    RuleNone = 0
    RuleNinjaId = 1
    RuleExternalNumber = 2
    RuleConversation = 3
    RuleSubject = 4
    RuleLearned = 5
End Enum

' ---------------------------------------------------------------------------
'  Entry point
' ---------------------------------------------------------------------------

Public Function MatchMail(ByVal subjectText As String, ByVal bodyText As String, ByVal conversationId As String, _
                          ByVal senderDomain As String, ByVal hasAutoHeader As Boolean, _
                          ByVal index As ITicketIndex, ByRef ticketId As Long, ByRef rule As MatchRule) As MatchOutcome
    Dim ids As Collection
    Dim numbers As Collection
    Dim found As Collection
    Dim ticket As MatchTicket
    Dim normalized As String
    Dim openCount As Long
    Dim hit As MatchTicket
    Dim i As Long

    ticketId = 0
    rule = RuleNone
    If index Is Nothing Then Exit Function
    If IsExcluded(senderDomain, hasAutoHeader) Then Exit Function

    ' Rule 1: NinjaOne id
    Set ids = New Collection
    If MatchLevels.LevelFor(RuleNinjaId) <> LevelOff Then
        Set ids = ExtractNinjaIds(subjectText)
        If ids.Count = 0 Then Set ids = ExtractNinjaIds(bodyText)
    End If
    If ids.Count > 1 Then
        MatchMail = MatchAmbiguous
        rule = RuleNinjaId
        Exit Function
    End If
    If ids.Count = 1 Then
        Set ticket = index.GetTicket(CLng(ids(1)))
        If Not ticket Is Nothing Then
            ticketId = ticket.Id
            rule = RuleNinjaId
            MatchMail = MatchFound
            Exit Function
        End If
    End If

    ' Rule 2: external ticket number
    Set numbers = New Collection
    If MatchLevels.LevelFor(RuleExternalNumber) <> LevelOff Then Set numbers = ExtractExternalNumbers(subjectText)
    If numbers.Count > 1 Then
        MatchMail = MatchAmbiguous
        rule = RuleExternalNumber
        Exit Function
    End If
    If numbers.Count = 1 Then
        rule = RuleExternalNumber
        Set found = index.FindByExternalNumber(CStr(numbers(1)))
        If found Is Nothing Then Exit Function
        For i = 1 To found.Count
            If found(i).IsOpen Then
                openCount = openCount + 1
                Set hit = found(i)
            End If
        Next i
        If found.Count = 1 Then
            ticketId = found(1).Id
            MatchMail = MatchFound
        ElseIf openCount = 1 Then
            ticketId = hit.Id
            MatchMail = MatchFound
        ElseIf found.Count > 1 Then
            MatchMail = MatchAmbiguous
        End If
        Exit Function
    End If

    ' Rule 3: conversation
    If Len(conversationId) > 0 And MatchLevels.LevelFor(RuleConversation) <> LevelOff Then
        ticketId = index.FindByConversation(conversationId)
        If ticketId > 0 Then
            rule = RuleConversation
            MatchMail = MatchFound
            Exit Function
        End If
    End If

    ' Rule 4: normalized subject against open tickets
    MatchMail = MatchBySubject(subjectText, index, ticketId)
    If MatchMail <> MatchNone Then
        rule = RuleSubject
        If MatchMail = MatchAmbiguous Then ticketId = 0
        Exit Function
    End If

    ' Rule 5: learned rules (domain and subject stem from the rules file, see LearnedRules)
    MatchMail = MatchByLearnedRule(subjectText, senderDomain, index, ticketId)
    If MatchMail <> MatchNone Then
        rule = RuleLearned
        If MatchMail = MatchAmbiguous Then ticketId = 0
    End If
End Function

' Rule 4: the normalized subject equals the summary of exactly one open ticket.
Private Function MatchBySubject(ByVal subjectText As String, ByVal index As ITicketIndex, ByRef ticketId As Long) As MatchOutcome
    Dim normalized As String
    Dim found As Collection
    Dim hit As MatchTicket
    Dim openCount As Long
    Dim i As Long

    ticketId = 0
    If MatchLevels.LevelFor(RuleSubject) = LevelOff Then Exit Function
    normalized = NormalizeSubject(subjectText)
    If Not IsSpecificSubject(normalized) Then Exit Function
    Set found = index.FindOpenBySubject(normalized)
    If found Is Nothing Then Exit Function

    For i = 1 To found.Count
        If found(i).IsOpen And NormalizeSubject(found(i).Summary) = normalized Then
            openCount = openCount + 1
            Set hit = found(i)
        End If
    Next i
    If openCount = 1 Then
        ticketId = hit.Id
        MatchBySubject = MatchFound
    ElseIf openCount > 1 Then
        MatchBySubject = MatchAmbiguous
    End If
End Function

' Rule 5: a learned rule (domain + subject stem) names the stem of a ticket group; exactly one open
' ticket whose title starts with that stem is the match.
Private Function MatchByLearnedRule(ByVal subjectText As String, ByVal senderDomain As String, ByVal index As ITicketIndex, _
                                    ByRef ticketId As Long) As MatchOutcome
    Dim prefix As String
    Dim found As Collection
    Dim hit As MatchTicket
    Dim openCount As Long
    Dim i As Long

    ticketId = 0
    If MatchLevels.LevelFor(RuleLearned) = LevelOff Then Exit Function

    prefix = LearnedRules.MatchingStem(senderDomain, SubjectStem(NormalizeSubject(subjectText)))
    If Len(prefix) = 0 Then Exit Function

    Set found = index.FindOpenByStem(prefix)
    If found Is Nothing Then Exit Function
    For i = 1 To found.Count
        If found(i).IsOpen And LearnedRules.StemStartsWith(SubjectStem(NormalizeSubject(found(i).Summary)), prefix) Then
            openCount = openCount + 1
            Set hit = found(i)
        End If
    Next i
    If openCount = 1 Then
        ticketId = hit.Id
        MatchByLearnedRule = MatchFound
    ElseIf openCount > 1 Then
        MatchByLearnedRule = MatchAmbiguous
    End If
End Function

' The stem of a normalized subject: digit runs become "#", the first words are kept (at most maxWords).
' "item235 - error 4711 in form220" gives "item# - error #"
Public Function SubjectStem(ByVal normalizedSubject As String, Optional ByVal maxWords As Long = 4) As String
    Dim i As Long
    Dim c As String
    Dim prevDigit As Boolean
    Dim s As String
    Dim words() As String

    For i = 1 To Len(normalizedSubject)
        c = Mid$(normalizedSubject, i, 1)
        If c Like "[0-9]" Then
            If Not prevDigit Then s = s & "#"
            prevDigit = True
        Else
            s = s & c
            prevDigit = False
        End If
    Next i

    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    If Len(s) = 0 Then Exit Function

    words = Split(s, " ")
    If UBound(words) + 1 > maxWords Then ReDim Preserve words(0 To maxWords - 1)
    SubjectStem = Join(words, " ")
End Function

' ---------------------------------------------------------------------------
'  Exclusions
' ---------------------------------------------------------------------------

' True for mails that are never matched: mailing lists / automatic mails and excluded domains.
Public Function IsExcluded(ByVal senderDomain As String, ByVal hasAutoHeader As Boolean) As Boolean
    Dim domains As Variant
    Dim i As Long

    If hasAutoHeader Then
        IsExcluded = True
        Exit Function
    End If

    senderDomain = LCase$(Trim$(senderDomain))
    If Len(senderDomain) = 0 Then Exit Function

    domains = Split(TicketConfig.EXCLUDED_DOMAINS, ",")
    For i = LBound(domains) To UBound(domains)
        If Len(Trim$(domains(i))) > 0 Then
            If senderDomain = LCase$(Trim$(domains(i))) Then
                IsExcluded = True
                Exit Function
            End If
        End If
    Next i
End Function

' True if the raw message header block contains a header of an automatic mail or mailing list.
Public Function HasAutoHeaderLine(ByVal headers As String) As Boolean
    Dim re As Object

    If Len(headers) = 0 Then Exit Function
    Set re = NewRegex("^(?:X-Auto-Response-Suppress|List-Id|List-Unsubscribe|Precedence\s*:\s*(?:bulk|list|junk)|Auto-Submitted\s*:\s*auto-)", False)
    re.Multiline = True
    ' Precedence/Auto-Submitted already carry their value in the pattern, the others only the name.
    HasAutoHeaderLine = re.Test(headers)
End Function

' One log line for a match decision, e.g.
' "Match found rule=ExternalNumber ticket=#20103 domain=vendor.example subject='...'".
' The subject is cut to TicketConfig.MATCH_LOG_SUBJECT_CHARS characters, line breaks removed. Privacy:
' for a mail without a match (MatchNone) only the hash of the normalized subject is logged
' ("subject_hash=1a2b3c4d", the same value as in the audit log) unless MATCH_LOG_SUBJECT_ON_NONE is set.
Public Function DescribeMatch(ByVal outcome As MatchOutcome, ByVal rule As MatchRule, ByVal ticketId As Long, _
                              ByVal senderDomain As String, ByVal subjectText As String) As String
    Dim text As String

    Select Case outcome
        Case MatchFound
            text = "Match found"
        Case MatchAmbiguous
            text = "Match ambiguous"
        Case Else
            text = "Match none"
    End Select

    If rule <> RuleNone Then text = text & " rule=" & RuleName(rule)
    If ticketId > 0 Then text = text & " ticket=#" & CStr(ticketId)
    If Len(senderDomain) > 0 Then text = text & " domain=" & senderDomain

    If TicketConfig.MATCH_LOG_SUBJECT_CHARS > 0 Then
        If outcome <> MatchNone Or TicketConfig.MATCH_LOG_SUBJECT_ON_NONE Then
            subjectText = Replace(Replace(Replace(subjectText, vbCrLf, " "), vbCr, " "), vbLf, " ")
            text = text & " subject='" & Left$(subjectText, TicketConfig.MATCH_LOG_SUBJECT_CHARS) & "'"
        ElseIf Len(subjectText) > 0 Then
            text = text & " subject_hash=" & AuditLogic.Hash32Hex(NormalizeSubject(subjectText))
        End If
    End If
    DescribeMatch = text
End Function

Public Function RuleName(ByVal rule As MatchRule) As String
    Select Case rule
        Case RuleNinjaId:         RuleName = "NinjaId"
        Case RuleExternalNumber:  RuleName = "ExternalNumber"
        Case RuleConversation:    RuleName = "Conversation"
        Case RuleSubject:         RuleName = "Subject"
        Case RuleLearned:         RuleName = "Learned"
        Case Else:                RuleName = "None"
    End Select
End Function

' ---------------------------------------------------------------------------
'  Extraction
' ---------------------------------------------------------------------------

' Distinct NinjaOne ticket ids (Collection of Long) from "Ticket 20108", "Ticket #20108",
' "(#20108)" and ".../ticket/20108". Digit runs of other lengths (vendor numbers) are ignored.
Public Function ExtractNinjaIds(ByVal text As String) As Collection
    Dim distinct As Collection
    Dim result As Collection
    Dim re As Object
    Dim matches As Object
    Dim m As Object
    Dim i As Long

    Set distinct = New Collection
    Set result = New Collection
    Set ExtractNinjaIds = result
    If Len(text) = 0 Then Exit Function

    Set re = NewRegex(TicketConfig.REGEX_NINJA_ID, True)
    Set matches = re.Execute(text)
    For Each m In matches
        AddDistinct distinct, CStr(CLng(m.SubMatches(0)))
    Next m

    For i = 1 To distinct.Count
        result.Add CLng(distinct(i))
    Next i
End Function

' Distinct vendor ticket numbers (Collection of String) from the subject, e.g.
' "[TicketNr. 900000105]" or "[CSM-TicketNr. 900000103]".
Public Function ExtractExternalNumbers(ByVal subjectText As String) As Collection
    Dim result As Collection
    Dim re As Object
    Dim matches As Object
    Dim m As Object

    Set result = New Collection
    Set ExtractExternalNumbers = result
    If Len(subjectText) = 0 Then Exit Function

    Set re = NewRegex(TicketConfig.REGEX_EXTERNAL_TICKET, True)
    Set matches = re.Execute(subjectText)
    For Each m In matches
        AddDistinct result, CStr(m.SubMatches(0))
    Next m
End Function

' ---------------------------------------------------------------------------
'  Subject normalization
' ---------------------------------------------------------------------------

' Lower case subject without reply/forward prefixes, vendor ticket tags, the vendor name
' prefix and the "TICKET (...) / " prefix of form tickets; white space collapsed.
Public Function NormalizeSubject(ByVal subjectText As String) As String
    Dim re As Object
    Dim s As String

    s = subjectText

    Set re = NewRegex("\[(?:CSM-?)?TicketNr\.?\s*\d+\]", True)
    s = re.Replace(s, " ")

    Set re = NewRegex("^\s*(?:(?:AW|RE|WG|FW|FWD|Antwort|Weiterleitung)\s*:\s*)+", False)
    s = re.Replace(s, "")

    Set re = NewRegex("^\s*(?:CSM|CMS)\s*\|\s*", False)
    s = re.Replace(s, "")

    ' Separators left behind by a removed tag ("CSM | [TicketNr. 1] - text").
    Set re = NewRegex("^[\s\-:|]+", False)
    s = re.Replace(s, "")

    Set re = NewRegex("^\s*" & TicketConfig.RegexTicketReplace(), False)
    s = re.Replace(s, "")

    ' Reply prefixes can also sit behind a removed ticket tag ("AW: [TicketNr. 1] AW: text").
    Set re = NewRegex("^\s*(?:(?:AW|RE|WG|FW|FWD)\s*:\s*)+", False)
    s = re.Replace(s, "")

    Set re = NewRegex(TicketConfig.OrganizationPrefixPattern(), False)
    s = re.Replace(s, "")

    Set re = NewRegex("^[\s\-:|]+", False)
    s = re.Replace(s, "")

    Set re = NewRegex("\s+", True)
    s = re.Replace(s, " ")

    NormalizeSubject = LCase$(Trim$(s))
End Function

' A subject may identify a ticket on its own only if it is long enough and has enough words;
' "Netzwerkausfall" and similar one-liners never match by subject alone.
Public Function IsSpecificSubject(ByVal normalizedSubject As String) As Boolean
    Dim words As Variant

    If Len(normalizedSubject) < TicketConfig.MATCH_MIN_SUBJECT_LEN Then Exit Function
    words = Split(normalizedSubject, " ")
    IsSpecificSubject = (UBound(words) - LBound(words) + 1 >= TicketConfig.MATCH_MIN_SUBJECT_WORDS)
End Function

' ---------------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------------

Private Function NewRegex(ByVal pattern As String, ByVal globalMatch As Boolean) As Object
    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = True
    re.Global = globalMatch
    Set NewRegex = re
End Function

Private Sub AddDistinct(ByVal target As Collection, ByVal value As String)
    On Error Resume Next
    target.Add value, "k" & value
    On Error GoTo 0
End Sub
