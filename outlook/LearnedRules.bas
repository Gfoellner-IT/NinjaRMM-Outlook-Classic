Attribute VB_Name = "LearnedRules"
' LearnedRules v0.1.0
'
' The learned matching rules (M7 step 4): a rules file next to the audit log,
' %LOCALAPPDATA%\NinjaRMM-Outlook-Classic\audit\match_rules.json. Rules are written by hand (or together
' with Claude after the report suggested a pattern); nothing ever adds a rule on its own.
'
'   {"version":1,"rules":[
'     {"id":"vendor-item235","domain":"vendor.example","stem":"item# - error","enabled":true,"note":"why"}
'   ]}
'
' A rule describes a SEARCH, not a ticket number: a mail from the domain ("*" = any) whose subject stem
' starts with the rule's stem words is assigned to the single open ticket whose title starts with the same
' stem words. Stems are written like the report shows them (lower case, digit runs as #). The rules run
' as the last matching rule at the level MATCH_LEVEL_LEARNED (default 1: log only). No Outlook types.

Option Explicit

Private pRules As Collection        ' of Dictionary (domain, stem)
Private pLoaded As Boolean
Private pUseText As Boolean
Private pText As String

Public Function RulesPath() As String
    RulesPath = Environ$("LOCALAPPDATA") & "\NinjaRMM-Outlook-Classic\audit\" & TicketConfig.LEARNED_RULES_FILE
End Function

' Number of valid, enabled rules.
Public Function Count() As Long
    EnsureLoaded
    Count = pRules.Count
End Function

' The stem of the first rule that fits a mail (its domain and the stem of its subject), "" if none.
Public Function MatchingStem(ByVal senderDomain As String, ByVal subjectStem As String) As String
    Dim rule As Object
    Dim domain As String

    EnsureLoaded
    domain = LCase$(Trim$(senderDomain))
    subjectStem = LCase$(Trim$(subjectStem))
    If Len(subjectStem) = 0 Then Exit Function

    For Each rule In pRules
        If rule("domain") = "*" Or rule("domain") = domain Then
            If StemStartsWith(subjectStem, CStr(rule("stem"))) Then
                MatchingStem = CStr(rule("stem"))
                Exit Function
            End If
        End If
    Next rule
End Function

' True if the words of stem begin with all words of prefix ("item# - error #" starts with "item# - error").
Public Function StemStartsWith(ByVal stem As String, ByVal prefix As String) As Boolean
    Dim s() As String
    Dim p() As String
    Dim i As Long

    stem = LCase$(Trim$(stem))
    prefix = LCase$(Trim$(prefix))
    If Len(stem) = 0 Or Len(prefix) = 0 Then Exit Function

    s = Split(stem, " ")
    p = Split(prefix, " ")
    If UBound(p) > UBound(s) Then Exit Function
    For i = 0 To UBound(p)
        If s(i) <> p(i) Then Exit Function
    Next i
    StemStartsWith = True
End Function

' Reads the rules file again at the next use (after the file was edited).
Public Sub Reload()
    pLoaded = False
End Sub

' ---------------------------------------------------------------------------
'  Test seams
' ---------------------------------------------------------------------------

' Uses this JSON text instead of the rules file.
Public Sub UseJson(ByVal text As String)
    pUseText = True
    pText = text
    pLoaded = False
End Sub

Public Sub ResetState()
    pUseText = False
    pText = ""
    pLoaded = False
End Sub

' ---------------------------------------------------------------------------
'  Implementation
' ---------------------------------------------------------------------------

Private Sub EnsureLoaded()
    Dim text As String

    If pLoaded Then Exit Sub
    pLoaded = True
    Set pRules = New Collection

    If pUseText Then
        text = pText
    ElseIf TicketLog.IsFileLoggingDisabled() Then
        Exit Sub                            ' unit tests never read the real rules file
    Else
        text = ReadUtf8(RulesPath())
    End If
    If Len(Trim$(text)) = 0 Then Exit Sub

    ParseRules text
End Sub

Private Sub ParseRules(ByVal text As String)
    Dim root As Variant
    Dim list As Object
    Dim item As Variant
    Dim rule As Object
    Dim domain As String
    Dim stem As String
    Dim skipped As Long

    If Not JsonLite.TryParse(text, root) Then
        TicketLog.Warn "Rules file: not valid JSON, no learned rules are used"
        Exit Sub
    End If
    If TypeName(root) <> "Dictionary" Then
        TicketLog.Warn "Rules file: unexpected content, no learned rules are used"
        Exit Sub
    End If
    Set list = JsonLite.GetNode(root, "rules")
    If list Is Nothing Then Exit Sub
    If TypeName(list) <> "Collection" Then Exit Sub

    For Each item In list
        If IsObject(item) Then
            If TypeName(item) = "Dictionary" Then
                domain = LCase$(Trim$(JsonLite.GetString(item, "domain")))
                stem = LCase$(Trim$(JsonLite.GetString(item, "stem")))
                If Not JsonLite.GetBoolean(item, "enabled", True) Then
                    ' switched off: not an error
                ElseIf Len(domain) = 0 Or Len(stem) < TicketConfig.LEARNED_STEM_MIN_CHARS Then
                    skipped = skipped + 1
                Else
                    Set rule = CreateObject("Scripting.Dictionary")
                    rule.Add "domain", domain
                    rule.Add "stem", stem
                    pRules.Add rule
                End If
            End If
        End If
    Next item

    If skipped > 0 Then TicketLog.Warn "Rules file: " & skipped & " rule(s) skipped (domain missing or stem shorter than " & TicketConfig.LEARNED_STEM_MIN_CHARS & " characters)"
End Sub

Private Function ReadUtf8(ByVal path As String) As String
    Dim fso As Object
    Dim stream As Object

    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadUtf8 = stream.ReadText
    stream.Close
    Exit Function

Failed:
    TicketLog.Warn "Rules file could not be read (" & Err.Description & ")"
End Function
