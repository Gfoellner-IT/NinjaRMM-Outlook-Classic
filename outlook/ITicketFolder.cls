VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITicketFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ITicketFolder v0.1.0
'
' One ticket folder ("#3621 (Text)"). The Outlook implementation is OutlookTicketFolder,
' unit tests use FakeTicketFolder.
Option Explicit

Public Property Get Name() As String
End Property

Public Property Let Name(ByVal value As String)
End Property

' Looks for mails in the folder whose text says "Status: ... Geschlossen" (offline fallback).
' Returns True and the newest receive time of such a mail, False if there is none.
Public Function FindClosedStatusMail(ByRef receivedTime As Date) As Boolean
End Function

' The mails of the folder received since the given time (Collection of IIncomingMail), used to find
' mails the user moved into the ticket folder by hand. Empty if there are none.
Public Function RecentMails(ByVal since As Date) As Collection
End Function
