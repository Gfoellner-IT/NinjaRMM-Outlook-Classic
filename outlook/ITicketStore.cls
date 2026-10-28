VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITicketStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ITicketStore v0.1.0
'
' The "Tickets" folder with its ticket folders and the archive below it
' (Archiv -> year -> month). The Outlook implementation is OutlookTicketStore,
' unit tests use FakeTicketStore.
Option Explicit

' Snapshot (Collection of ITicketFolder) of all active ticket folders, without the archive folder.
Public Function ListActiveFolders() As Collection
End Function

' Active ticket folder for a ticket number like "#3621", Nothing if there is none.
Public Function FindActiveFolder(ByVal ticketNumber As String) As ITicketFolder
End Function

' Folder of that ticket anywhere in the archive, Nothing if there is none.
Public Function FindArchivedFolder(ByVal ticketNumber As String) As ITicketFolder
End Function

Public Function CreateActiveFolder(ByVal folderName As String) As ITicketFolder
End Function

' Moves a folder found in the archive back to the active ticket folders.
Public Sub RestoreFromArchive(ByVal folder As ITicketFolder)
End Sub

' Moves an active folder to Archiv -> year -> month of closedDate (0 = today).
Public Sub MoveToArchive(ByVal folder As ITicketFolder, ByVal closedDate As Date)
End Sub

Public Sub MoveMailToFolder(ByVal mail As IIncomingMail, ByVal folder As ITicketFolder)
End Sub
