VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutlookTicketStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OutlookTicketStore v0.1.0
'
' Thin adapter: ITicketStore on top of the Outlook folder "Tickets" (below the inbox)
' with its ticket folders and the archive Archiv -> year -> month.
' The decisions are made in TicketService; this class only talks to Outlook.
Implements ITicketStore
Option Explicit

Private pTickets As Outlook.Folder

Public Sub Init(ByVal ticketsFolder As Outlook.Folder)
    Set pTickets = ticketsFolder
End Sub

' ---------------------------------------------------------------------------
'  ITicketStore
' ---------------------------------------------------------------------------

Private Function ITicketStore_ListActiveFolders() As Collection
    Dim result As Collection
    Dim i As Long
    Dim child As Outlook.Folder

    Set result = New Collection
    For i = pTickets.Folders.Count To 1 Step -1
        Set child = pTickets.Folders(i)
        If Not IsArchiveFolder(child) Then result.Add Wrap(child)
    Next i
    Set ITicketStore_ListActiveFolders = result
End Function

Private Function ITicketStore_FindActiveFolder(ByVal ticketNumber As String) As ITicketFolder
    Dim child As Outlook.Folder

    For Each child In pTickets.Folders
        If Not IsArchiveFolder(child) Then
            If TicketParsing.FolderMatchesTicket(child.Name, ticketNumber) Then
                Set ITicketStore_FindActiveFolder = Wrap(child)
                Exit Function
            End If
        End If
    Next child
End Function

Private Function ITicketStore_FindArchivedFolder(ByVal ticketNumber As String) As ITicketFolder
    Dim archive As Outlook.Folder
    Dim found As Outlook.Folder

    Set archive = FindChild(pTickets, TicketConfig.FOLDER_ARCHIVE)
    If archive Is Nothing Then Exit Function

    Set found = FindTicketFolderRecursively(archive, ticketNumber)
    If Not found Is Nothing Then Set ITicketStore_FindArchivedFolder = Wrap(found)
End Function

Private Function ITicketStore_CreateActiveFolder(ByVal folderName As String) As ITicketFolder
    Set ITicketStore_CreateActiveFolder = Wrap(GetOrCreateFolder(pTickets, folderName))
End Function

Private Sub ITicketStore_RestoreFromArchive(ByVal folder As ITicketFolder)
    RawFolder(folder).MoveTo pTickets
End Sub

Private Sub ITicketStore_MoveToArchive(ByVal folder As ITicketFolder, ByVal closedDate As Date)
    Dim archive As Outlook.Folder
    Dim yearFolder As Outlook.Folder
    Dim monthFolder As Outlook.Folder

    Set archive = GetOrCreateFolder(pTickets, TicketConfig.FOLDER_ARCHIVE)
    Set yearFolder = GetOrCreateFolder(archive, TicketParsing.ArchiveYearName(closedDate))
    Set monthFolder = GetOrCreateFolder(yearFolder, TicketParsing.ArchiveMonthName(closedDate))
    RawFolder(folder).MoveTo monthFolder
End Sub

Private Sub ITicketStore_MoveMailToFolder(ByVal mail As IIncomingMail, ByVal folder As ITicketFolder)
    Dim adapter As OutlookIncomingMail

    Set adapter = mail
    adapter.Raw.Move RawFolder(folder)
End Sub

' ---------------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------------

Private Function Wrap(ByVal folder As Outlook.Folder) As ITicketFolder
    Dim adapter As OutlookTicketFolder

    Set adapter = New OutlookTicketFolder
    adapter.Init folder
    Set Wrap = adapter
End Function

Private Function RawFolder(ByVal folder As ITicketFolder) As Outlook.Folder
    Dim adapter As OutlookTicketFolder

    Set adapter = folder
    Set RawFolder = adapter.Raw
End Function

Private Function IsArchiveFolder(ByVal folder As Outlook.Folder) As Boolean
    IsArchiveFolder = (LCase$(folder.Name) = LCase$(TicketConfig.FOLDER_ARCHIVE))
End Function

Private Function FindChild(ByVal parent As Outlook.Folder, ByVal folderName As String) As Outlook.Folder
    Dim child As Outlook.Folder

    For Each child In parent.Folders
        If LCase$(child.Name) = LCase$(folderName) Then
            Set FindChild = child
            Exit Function
        End If
    Next child
End Function

Private Function GetOrCreateFolder(ByVal parent As Outlook.Folder, ByVal folderName As String) As Outlook.Folder
    Dim found As Outlook.Folder

    Set found = FindChild(parent, folderName)
    If found Is Nothing Then Set found = parent.Folders.Add(folderName)
    Set GetOrCreateFolder = found
End Function

Private Function FindTicketFolderRecursively(ByVal parent As Outlook.Folder, ByVal ticketNumber As String) As Outlook.Folder
    Dim child As Outlook.Folder
    Dim found As Outlook.Folder

    For Each child In parent.Folders
        If TicketParsing.FolderMatchesTicket(child.Name, ticketNumber) Then
            Set FindTicketFolderRecursively = child
            Exit Function
        End If
        Set found = FindTicketFolderRecursively(child, ticketNumber)
        If Not found Is Nothing Then
            Set FindTicketFolderRecursively = found
            Exit Function
        End If
    Next child
End Function
