VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FakeMail v0.1.0
'
' Test double for IIncomingMail. MovedTo is set by FakeTicketStore.MoveMailToFolder.
Implements IIncomingMail
Option Explicit

Public SenderName As String
Public Subject As String
Public SenderDomain As String
Public ConversationId As String
Public HasAutoHeader As Boolean
Public BodyText As String
Public ReceivedTime As Date
Public MessageKey As String
Public MovedTo As FakeTicketFolder

Private Property Get IIncomingMail_SenderName() As String
    IIncomingMail_SenderName = SenderName
End Property

Private Property Get IIncomingMail_Subject() As String
    IIncomingMail_Subject = Subject
End Property

Private Property Get IIncomingMail_SenderDomain() As String
    IIncomingMail_SenderDomain = SenderDomain
End Property

Private Property Get IIncomingMail_ConversationId() As String
    IIncomingMail_ConversationId = ConversationId
End Property

Private Property Get IIncomingMail_HasAutoHeader() As Boolean
    IIncomingMail_HasAutoHeader = HasAutoHeader
End Property

Private Property Get IIncomingMail_BodyText() As String
    IIncomingMail_BodyText = BodyText
End Property

Private Property Get IIncomingMail_ReceivedTime() As Date
    IIncomingMail_ReceivedTime = ReceivedTime
End Property

Private Property Get IIncomingMail_MessageKey() As String
    IIncomingMail_MessageKey = MessageKey
End Property
