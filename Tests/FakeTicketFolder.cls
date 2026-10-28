VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeTicketFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FakeTicketFolder v0.1.0
'
' Test double for ITicketFolder (an in-memory folder).
Implements ITicketFolder
Option Explicit

Public FolderName As String
Public MailCount As Long

' Set these to simulate a mail in the folder whose text says "Status: ... Geschlossen".
Public HasClosedStatusMail As Boolean
Public ClosedStatusMailDate As Date

Public MailChecks As Long   ' how often the mails were scanned
Public ArchivedFor As Date  ' archive date the store used (0 = today), set by FakeTicketStore.MoveToArchive
Public IsArchived As Boolean
Public RenameCount As Long
Public FailOnNameRead As Boolean   ' simulates an Outlook error when the folder is accessed

' Mails "in the folder" for RecentMails (FakeMail objects with a ReceivedTime).
Public Mails As Collection
Public RecentMailsRequests As Long

Private Property Get ITicketFolder_Name() As String
    If FailOnNameRead Then Err.Raise 5, "FakeTicketFolder", "simulated folder error"
    ITicketFolder_Name = FolderName
End Property

Private Property Let ITicketFolder_Name(ByVal value As String)
    FolderName = value
    RenameCount = RenameCount + 1
End Property

Private Function ITicketFolder_FindClosedStatusMail(ByRef receivedTime As Date) As Boolean
    MailChecks = MailChecks + 1
    If HasClosedStatusMail Then
        receivedTime = ClosedStatusMailDate
        ITicketFolder_FindClosedStatusMail = True
    End If
End Function

Public Sub AddMail(ByVal mail As FakeMail)
    If Mails Is Nothing Then Set Mails = New Collection
    Mails.Add mail
End Sub

Private Function ITicketFolder_RecentMails(ByVal since As Date) As Collection
    Dim result As Collection
    Dim mail As FakeMail

    RecentMailsRequests = RecentMailsRequests + 1
    Set result = New Collection
    If Not Mails Is Nothing Then
        For Each mail In Mails
            If mail.ReceivedTime >= since Then result.Add mail
        Next mail
    End If
    Set ITicketFolder_RecentMails = result
End Function
