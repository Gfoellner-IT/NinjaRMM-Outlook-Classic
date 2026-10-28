Attribute VB_Name = "MatchLevels"
' MatchLevels v0.1.0
'
' Safety level of each matching rule (M8). The effective level of a rule is the lower of its own
' configured level and the global cap TicketConfig.MATCH_LEVEL_MAX:
'
' The learned rules from the rules file (LearnedRules, M7) count as one rule here (MATCH_LEVEL_LEARNED).
'   LevelOff      the rule is switched off: no lookup, no log, the mail is not looked at by this rule
'   LevelLogOnly  the decision is logged, nothing is moved (default)
'   LevelConfirm  the mail is queued as a proposal and moved only after the user confirms
'   LevelAuto     the mail is moved without asking
'
' Whatever the level: closed or unreadable tickets, ambiguous matches and excluded mails are never
' moved. No Outlook types. UseLevels/UseAll/ResetState are test seams.

Option Explicit

Public Enum MatchLevel
    LevelOff = 0
    LevelLogOnly = 1
    LevelConfirm = 2
    LevelAuto = 3
End Enum

Private pOverride As Boolean
Private pNinjaId As Long
Private pExternalNumber As Long
Private pConversation As Long
Private pSubject As Long
Private pLearned As Long
Private pCap As Long

' Effective level of a rule (LevelOff for RuleNone).
Public Function LevelFor(ByVal rule As MatchRule) As MatchLevel
    Dim configured As Long
    Dim cap As Long

    If pOverride Then cap = pCap Else cap = TicketConfig.MATCH_LEVEL_MAX

    Select Case rule
        Case RuleNinjaId
            If pOverride Then configured = pNinjaId Else configured = TicketConfig.MATCH_LEVEL_NINJA_ID
        Case RuleExternalNumber
            If pOverride Then configured = pExternalNumber Else configured = TicketConfig.MATCH_LEVEL_EXTERNAL_NUMBER
        Case RuleConversation
            If pOverride Then configured = pConversation Else configured = TicketConfig.MATCH_LEVEL_CONVERSATION
        Case RuleSubject
            If pOverride Then configured = pSubject Else configured = TicketConfig.MATCH_LEVEL_SUBJECT
        Case RuleLearned
            If pOverride Then configured = pLearned Else configured = TicketConfig.MATCH_LEVEL_LEARNED
        Case Else
            Exit Function
    End Select

    LevelFor = Clamp(Lower(configured, cap))
End Function

' True if at least one rule has exactly this effective level.
Public Function AnyRuleAtLevel(ByVal level As MatchLevel) As Boolean
    AnyRuleAtLevel = (LevelFor(RuleNinjaId) = level) Or (LevelFor(RuleExternalNumber) = level) Or _
                     (LevelFor(RuleConversation) = level) Or (LevelFor(RuleSubject) = level) Or _
                     (LevelFor(RuleLearned) = level)
End Function

' True if at least one rule is active (otherwise nothing needs to be looked at).
Public Function AnyRuleActive() As Boolean
    AnyRuleActive = (LevelFor(RuleNinjaId) <> LevelOff) Or (LevelFor(RuleExternalNumber) <> LevelOff) Or _
                    (LevelFor(RuleConversation) <> LevelOff) Or (LevelFor(RuleSubject) <> LevelOff) Or _
                    (LevelFor(RuleLearned) <> LevelOff)
End Function

Public Function LevelName(ByVal level As MatchLevel) As String
    Select Case level
        Case LevelLogOnly:  LevelName = "log_only"
        Case LevelConfirm:  LevelName = "confirm"
        Case LevelAuto:     LevelName = "auto"
        Case Else:          LevelName = "off"
    End Select
End Function

' ---------------------------------------------------------------------------
'  Test seams
' ---------------------------------------------------------------------------

Public Sub UseLevels(ByVal ninjaId As Long, ByVal externalNumber As Long, ByVal conversation As Long, _
                     ByVal subject As Long, Optional ByVal cap As Long = 3, Optional ByVal learned As Long = 0)
    pOverride = True
    pNinjaId = ninjaId
    pExternalNumber = externalNumber
    pConversation = conversation
    pSubject = subject
    pLearned = learned
    pCap = cap
End Sub

Public Sub UseAll(ByVal level As MatchLevel)
    UseLevels level, level, level, level, LevelAuto, level
End Sub

Public Sub ResetState()
    pOverride = False
End Sub

' ---------------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------------

Private Function Lower(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then Lower = a Else Lower = b
End Function

Private Function Clamp(ByVal value As Long) As Long
    If value < LevelOff Then
        Clamp = LevelOff
    ElseIf value > LevelAuto Then
        Clamp = LevelAuto
    Else
        Clamp = value
    End If
End Function
