VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IIncomingMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' IIncomingMail v0.1.0
'
' The parts of a received mail the ticket automation needs. The Outlook implementation
' is OutlookIncomingMail, unit tests use FakeMail.
Option Explicit

Public Property Get SenderName() As String
End Property

Public Property Get Subject() As String
End Property

' Domain of the sender's SMTP address, lower case ("" if unknown).
Public Property Get SenderDomain() As String
End Property

' Outlook conversation id (same for all mails of a thread, "" if unknown).
Public Property Get ConversationId() As String
End Property

' True for automatic mails and mailing lists (Auto-Submitted, List-Id, Precedence, ...).
Public Property Get HasAutoHeader() As Boolean
End Property

' Plain text body, cut to TicketConfig.MATCH_BODY_MAX_CHARS ("" if unavailable).
Public Property Get BodyText() As String
End Property

Public Property Get ReceivedTime() As Date
End Property

' Stable key of the mail (8 hex digits, hash of the internet message id), the same after moving
' the mail. Used by the audit log to find a mail again; "" if unknown.
Public Property Get MessageKey() As String
End Property
