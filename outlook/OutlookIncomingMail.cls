VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutlookIncomingMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OutlookIncomingMail v0.1.0
'
' Thin adapter: IIncomingMail on top of an Outlook.MailItem. Contains no logic.
Implements IIncomingMail
Option Explicit

Private pMail As Outlook.MailItem

Public Sub Init(ByVal mail As Outlook.MailItem)
    Set pMail = mail
End Sub

Public Property Get Raw() As Outlook.MailItem
    Set Raw = pMail
End Property

Private Property Get IIncomingMail_SenderName() As String
    IIncomingMail_SenderName = pMail.SenderName
End Property

Private Property Get IIncomingMail_Subject() As String
    IIncomingMail_Subject = pMail.Subject
End Property

Private Property Get IIncomingMail_SenderDomain() As String
    Dim address As String
    Dim user As Outlook.ExchangeUser
    Dim p As Long

    On Error GoTo Failed
    If pMail.SenderEmailType = "EX" Then
        Set user = pMail.Sender.GetExchangeUser
        If Not user Is Nothing Then address = user.PrimarySmtpAddress
    Else
        address = pMail.SenderEmailAddress
    End If

    p = InStrRev(address, "@")
    If p > 0 Then IIncomingMail_SenderDomain = LCase$(Trim$(Mid$(address, p + 1)))
    Exit Property

Failed:
    IIncomingMail_SenderDomain = ""
End Property

Private Property Get IIncomingMail_ConversationId() As String
    On Error Resume Next
    IIncomingMail_ConversationId = pMail.ConversationID
    If Err.Number <> 0 Then IIncomingMail_ConversationId = ""
End Property

Private Property Get IIncomingMail_HasAutoHeader() As Boolean
    Dim headers As String

    On Error GoTo Failed
    headers = CStr(pMail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F"))
    IIncomingMail_HasAutoHeader = TicketMatcher.HasAutoHeaderLine(headers)
    Exit Property

Failed:
    IIncomingMail_HasAutoHeader = False
End Property

Private Property Get IIncomingMail_BodyText() As String
    On Error Resume Next
    IIncomingMail_BodyText = Left$(pMail.Body, TicketConfig.MATCH_BODY_MAX_CHARS)
    If Err.Number <> 0 Then IIncomingMail_BodyText = ""
End Property

Private Property Get IIncomingMail_ReceivedTime() As Date
    On Error Resume Next
    IIncomingMail_ReceivedTime = pMail.ReceivedTime
End Property

' Hash of the internet message id (stays the same when the mail is moved between folders,
' unlike the EntryID). Without a message id: hash of received time and subject.
Private Property Get IIncomingMail_MessageKey() As String
    Dim id As String

    On Error Resume Next
    id = CStr(pMail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F"))
    If Err.Number <> 0 Then id = ""
    Err.Clear
    If Len(id) = 0 Then id = Format$(pMail.ReceivedTime, "yyyymmddhhnnss") & "|" & pMail.Subject
    IIncomingMail_MessageKey = AuditLogic.Hash32Hex(id)
End Property
