VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeTicketStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FakeTicketStore v0.1.0
'
' Test double for ITicketStore: active and archived folders are kept in memory.
Implements ITicketStore
Option Explicit

Private pActive As Collection
Private pArchived As Collection

Public CreatedCount As Long
Public RestoredCount As Long

Public FailMoves As Boolean   ' simulates a mail that cannot be moved any more

Private Sub Class_Initialize()
    Set pActive = New Collection
    Set pArchived = New Collection
End Sub

' ---------------------------------------------------------------------------
'  Test setup and inspection
' ---------------------------------------------------------------------------

Public Function AddActive(ByVal folderName As String) As FakeTicketFolder
    Dim folder As FakeTicketFolder

    Set folder = New FakeTicketFolder
    folder.FolderName = folderName
    pActive.Add folder
    Set AddActive = folder
End Function

Public Function AddArchived(ByVal folderName As String) As FakeTicketFolder
    Dim folder As FakeTicketFolder

    Set folder = New FakeTicketFolder
    folder.FolderName = folderName
    folder.IsArchived = True
    pArchived.Add folder
    Set AddArchived = folder
End Function

Public Property Get ActiveCount() As Long
    ActiveCount = pActive.Count
End Property

Public Property Get ArchivedCount() As Long
    ArchivedCount = pArchived.Count
End Property

Public Function ActiveAt(ByVal index As Long) As FakeTicketFolder
    Set ActiveAt = pActive(index)
End Function

Public Function ArchivedAt(ByVal index As Long) As FakeTicketFolder
    Set ArchivedAt = pArchived(index)
End Function

' ---------------------------------------------------------------------------
'  ITicketStore
' ---------------------------------------------------------------------------

Private Function ITicketStore_ListActiveFolders() As Collection
    Dim result As Collection
    Dim i As Long

    Set result = New Collection
    For i = 1 To pActive.Count
        result.Add pActive(i)
    Next i
    Set ITicketStore_ListActiveFolders = result
End Function

Private Function ITicketStore_FindActiveFolder(ByVal ticketNumber As String) As ITicketFolder
    Dim i As Long

    For i = 1 To pActive.Count
        If TicketParsing.FolderMatchesTicket(pActive(i).FolderName, ticketNumber) Then
            Set ITicketStore_FindActiveFolder = pActive(i)
            Exit Function
        End If
    Next i
End Function

Private Function ITicketStore_FindArchivedFolder(ByVal ticketNumber As String) As ITicketFolder
    Dim i As Long

    For i = 1 To pArchived.Count
        If TicketParsing.FolderMatchesTicket(pArchived(i).FolderName, ticketNumber) Then
            Set ITicketStore_FindArchivedFolder = pArchived(i)
            Exit Function
        End If
    Next i
End Function

Private Function ITicketStore_CreateActiveFolder(ByVal folderName As String) As ITicketFolder
    Set ITicketStore_CreateActiveFolder = AddActive(folderName)
    CreatedCount = CreatedCount + 1
End Function

Private Sub ITicketStore_RestoreFromArchive(ByVal folder As ITicketFolder)
    Dim fake As FakeTicketFolder
    Dim i As Long

    Set fake = folder
    For i = 1 To pArchived.Count
        If pArchived(i) Is fake Then
            pArchived.Remove i
            Exit For
        End If
    Next i
    fake.IsArchived = False
    pActive.Add fake
    RestoredCount = RestoredCount + 1
End Sub

Private Sub ITicketStore_MoveToArchive(ByVal folder As ITicketFolder, ByVal closedDate As Date)
    Dim fake As FakeTicketFolder
    Dim i As Long

    Set fake = folder
    For i = 1 To pActive.Count
        If pActive(i) Is fake Then
            pActive.Remove i
            Exit For
        End If
    Next i
    fake.IsArchived = True
    fake.ArchivedFor = closedDate
    pArchived.Add fake
End Sub

Private Sub ITicketStore_MoveMailToFolder(ByVal mail As IIncomingMail, ByVal folder As ITicketFolder)
    Dim fakeMail As FakeMail
    Dim fakeFolder As FakeTicketFolder

    If FailMoves Then Err.Raise 91, "FakeTicketStore", "the mail was moved or deleted (simulated)"

    Set fakeMail = mail
    Set fakeFolder = folder
    Set fakeMail.MovedTo = fakeFolder
    fakeFolder.MailCount = fakeFolder.MailCount + 1
End Sub
