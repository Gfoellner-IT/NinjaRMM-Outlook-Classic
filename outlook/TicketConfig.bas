Attribute VB_Name = "TicketConfig"
' TicketConfig v0.2.0
'
' Configuration of the ticket automation. Files are ASCII-only, values with
' non-ASCII characters are built with ChrW (246 = o umlaut).

Option Explicit

' Outlook folder names (user visible - renaming them requires renaming the folders too).
Public Const FOLDER_TICKETS As String = "Tickets"
Public Const FOLDER_ARCHIVE As String = "Archiv"

' Ticket mail subject, e.g. "[example] (#3621) TICKET (B..ro - IT) / Printer broken".


' Mail body line that marks a closed ticket ("Status: ... Geschlossen"). Only used by the offline fallback.
Public Const REGEX_STATUS_CLOSED As String = "Status:\s?.*?\s?Geschlossen"

' If the NinjaOne API cannot be reached, a ticket folder is still archived when one of its
' mails says "Status: ... Geschlossen". The archive date is then the date of that mail.
' Set to False to do nothing while the API is unavailable (safest, retried at the next run).
Public Const ENABLE_OFFLINE_FALLBACK As Boolean = True

' Wait time (ms) before the mail watcher processes a newly arrived mail, so that other rules
' can move it first.
Public Const EVENT_DELAY_MS As Long = 1000

' --- Matching of foreign mails to existing tickets (TicketMatcher) -------------------------

' NinjaOne ticket id in free text: "Ticket 20108", "Ticket #20108", "Ticket-Nr. 20108", "(#20108)",
' ".../ticket/20108". Only 4 to 6 digits; longer digit runs (vendor numbers) do not match.
Public Const REGEX_NINJA_ID As String = "(?:ticket\s*-?\s*(?:nr\.?\s*)?#?\s*|/ticket/|\(#)(\d{4,6})(?!\d)"

' Vendor (external) ticket number in a subject: "[TicketNr. 900000105]", "[CSM-TicketNr. 900000103]".
Public Const REGEX_EXTERNAL_TICKET As String = "\[(?:CSM-?)?TicketNr\.?\s*(\d{6,12})\]"

' Sender domains (comma separated) whose mails are never matched.


' A subject alone may only match a ticket if it has at least this many characters and words.
Public Const MATCH_MIN_SUBJECT_LEN As Long = 20
Public Const MATCH_MIN_SUBJECT_WORDS As Long = 3

' Board that lists all tickets (open and closed): "Alle Tickets".


' Ticket list cache of NinjaTicketIndex: page size of board queries and minutes until the
' list of open tickets / the answer for an external number is fetched again.
Public Const INDEX_PAGE_SIZE As Long = 200
Public Const INDEX_CACHE_MINUTES As Long = 10

' Safety level of each matching rule (see MatchLevels):
'   0 = off (rule not used at all)   1 = log only, never move (default)
'   2 = move after the user confirms   3 = move automatically
' The effective level of a rule is the lower of its own level and MATCH_LEVEL_MAX.
Public Const MATCH_LEVEL_NINJA_ID As Long = 1
Public Const MATCH_LEVEL_EXTERNAL_NUMBER As Long = 1
Public Const MATCH_LEVEL_CONVERSATION As Long = 1
Public Const MATCH_LEVEL_SUBJECT As Long = 1
Public Const MATCH_LEVEL_LEARNED As Long = 1       ' learned rules from the rules file (match_rules.json)
Public Const MATCH_LEVEL_MAX As Long = 3

' Dialog of the confirmation level: 1 = one message box per proposal, 2 = review form ReviewForm with a list
' (falls back to message boxes if the form is not in the VBA project).
Public Const REVIEW_DIALOG As Long = 2

' Learned rules (LearnedRules): name of the rules file next to the audit log and the minimum length of a stem.
Public Const LEARNED_RULES_FILE As String = "match_rules.json"
Public Const LEARNED_STEM_MIN_CHARS As Long = 8

' Minutes between two "There are new proposals, review them now?" questions (level 2).
Public Const MATCH_HINT_MINUTES As Long = 15

' The sweep of RunEmailRule only looks at foreign mails received within this many days
' (the dry run RunMatchDryRun looks at the newest mails regardless of age).
Public Const MATCH_SWEEP_DAYS As Long = 14

' Characters of the mail body that are read for the match (a ticket id sits near the top).
Public Const MATCH_BODY_MAX_CHARS As Long = 4000

' --- Audit log of the match decisions (TicketAudit, M7) ---------------------------------------

' False switches the audit log off completely.
Public Const ENABLE_AUDIT_LOG As Boolean = True

' Months a monthly audit file is kept (older files are deleted by the first write after Outlook starts).
Public Const AUDIT_RETENTION_MONTHS As Long = 12

' --- Report on the audit log (AuditReport, M7) ---------------------------------------------

' Folder for the reports ("" = %LOCALAPPDATA%\NinjaRMM-Outlook-Classic\reports) and how many months back are read.
' Manual assignments (mails moved into ticket folders by hand) are looked for among the mails received this many days back.
Public Const MANUAL_SCAN_DAYS As Long = 30

Public Const REPORT_FOLDER As String = ""
Public Const REPORT_MONTHS_BACK As Long = 12

' A pattern is only reported as a candidate for a new rule with at least this many confirmations
' (and without rejections, in at least 2 conversations).
Public Const REPORT_MIN_SUPPORT As Long = 5

' Level recommendations: raise to 3 after this many decisions at this precision (percent);
' lower or tighten below the other precision with at least the other number of decisions.
Public Const REPORT_LEVEL3_MIN_DECISIONS As Long = 30
Public Const REPORT_LEVEL3_MIN_PERCENT As Long = 97
Public Const REPORT_LOWER_MIN_DECISIONS As Long = 10
Public Const REPORT_LOWER_BELOW_PERCENT As Long = 80

' Share of the decided mails (oldest first) used to find patterns; the rest checks them.
Public Const REPORT_TRAIN_PERCENT As Long = 70

' Write the normalized subject into the audit log (False = no subject at all) and its maximum length.
Public Const AUDIT_LOG_SUBJECT As Boolean = True
Public Const AUDIT_SUBJECT_MAX_CHARS As Long = 120

' False (default): mails without a match are logged with the hash of their subject only, never the
' subject itself (subjects can contain person names). Set True for debugging.
Public Const AUDIT_LOG_SUBJECT_ON_NONE As Boolean = False

' Characters of the subject written to the log for a match decision (0 = no subject in the log).
Public Const MATCH_LOG_SUBJECT_CHARS As Long = 60

' Same for the text log (TicketLog): False = mails without a match only get the hash of their subject.
Public Const MATCH_LOG_SUBJECT_ON_NONE As Boolean = False

' Only mails from this sender (display name) are sorted into ticket folders.
Public Function AllowedSender() As String
    AllowedSender = LocalSettings.Value("AllowedSender", "IT Support | B" & ChrW$(252) & "ro")
End Function

' Optional prefix in the subject text: "TICKET (B..ro - IT) / " becomes "IT / ".
Public Function RegexTicketReplace() As String
    RegexTicketReplace = LocalSettings.Value("TicketPrefixPattern", "TICKET \(B" & ChrW$(252) & "ro - ([A-Z]+)\)( / )")
End Function

Public Function REGEX_SUBJECT_PATTERN() As String
    REGEX_SUBJECT_PATTERN = LocalSettings.Value("TicketSubjectPattern", "^\[example\]\s*\(#(\d+)\)(.*)")
End Function

Public Function NINJA_ALL_TICKETS_BOARD() As Long
    NINJA_ALL_TICKETS_BOARD = LocalSettings.AllTicketsBoard()
End Function

Public Function OrganizationPrefixPattern() As String
    OrganizationPrefixPattern = LocalSettings.Value("OrganizationPrefixPattern", "^\s*B.{1,2}ro\s*:?\s*-?\s*")
End Function
Public Function EXCLUDED_DOMAINS() As String
    EXCLUDED_DOMAINS = LocalSettings.Value("ExcludedDomains", "github.com,newsletter.example")
End Function
