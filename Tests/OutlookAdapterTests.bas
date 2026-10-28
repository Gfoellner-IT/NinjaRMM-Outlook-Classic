Attribute VB_Name = "OutlookAdapterTests"
' OutlookAdapterTests v0.1.0
'
' Integration tests of the Outlook adapters (OutlookTicketStore, OutlookTicketFolder,
' OutlookIncomingMail) and one end-to-end run of TicketService. They use Outlook types and
' run only in the test data file "NinjaTest" (see OutlookTestGuard), Posteingang\Tickets.
' Everything they create is deleted again. Start with IntegrationRunner.RunIntegrationTests.
' The NinjaOne API is replaced by FakeNinjaTransport, no network access.
'
' Test mails are NOT created with Items.Add + Save: that saves a draft in the default mailbox
' (the real one). They are written to a temporary .msg file, opened with OpenSharedItem and
' moved into the test folder, so the real mailbox is never touched (Test_NoLeakIntoRealMailbox).

Option Explicit

Private Const PR_SENDER_NAME As String = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"

Private pTickets As Outlook.Folder
Private pInbox As Outlook.Folder
Private pStoreAdapter As OutlookTicketStore
Private pStore As ITicketStore

' Result of the last CreateMail: could the sender name be set on the test mail?
Private pSenderSettable As Boolean
Private pMailCounter As Long
Private pLeakDetected As Boolean   ' once a draft shows up in the real mailbox, no further test mails are created

' ---------------------------------------------------------------------------
'  Setup and helpers
' ---------------------------------------------------------------------------

Private Sub Setup()
    Set pTickets = OutlookTestGuard.FindTestTicketsFolder()
    If pTickets Is Nothing Then
        Err.Raise vbObjectError + 11100, "OutlookAdapterTests", _
            "Test store '" & OutlookTestGuard.TEST_STORE_NAME & "' with Posteingang\Tickets not usable. " & OutlookTestGuard.LastProblem
    End If
    Set pInbox = pTickets.Parent

    TicketLog.DisableFileLogging
    TicketLog.Clear
    NinjaApi.ResetState
    Clean

    Set pStoreAdapter = New OutlookTicketStore
    pStoreAdapter.Init pTickets
    Set pStore = pStoreAdapter
End Sub

Private Sub Teardown()
    On Error Resume Next
    NinjaApi.ResetState
    Clean
    Set pStore = Nothing
    Set pStoreAdapter = Nothing
    Set pInbox = Nothing
    Set pTickets = Nothing
End Sub

' Removes everything below Tickets, all items in the test inbox and the test store's deleted items.
' Only ever called with folders of the test store (they come from OutlookTestGuard).
Private Sub Clean()
    Dim i As Long
    Dim deleted As Outlook.Folder

    For i = pTickets.Folders.Count To 1 Step -1
        pTickets.Folders(i).Delete
    Next i
    For i = pInbox.Items.Count To 1 Step -1
        pInbox.Items(i).Delete
    Next i

    Set deleted = OutlookTestGuard.DeletedItemsOf(pTickets.Store)
    If Not deleted Is Nothing Then
        For i = deleted.Folders.Count To 1 Step -1
            deleted.Folders(i).Delete
        Next i
        For i = deleted.Items.Count To 1 Step -1
            deleted.Items(i).Delete
        Next i
    End If
End Sub

Private Function CreateMail(ByVal folder As Outlook.Folder, ByVal subject As String, ByVal body As String, _
        Optional ByVal sender As String = "") As Outlook.MailItem
    Dim draft As Outlook.MailItem
    Dim opened As Outlook.MailItem
    Dim path As String
    Dim realDrafts As Outlook.Folder
    Dim draftsBefore As Long

    If pLeakDetected Then
        Err.Raise vbObjectError + 11102, "OutlookAdapterTests", "Stopped: an earlier test mail appeared in the real Drafts folder."
    End If

    ' Only ever create mails in the test store (a safety net on top of the guard).
    If Not OutlookTestGuard.IsTestStore(folder.Store) Then
        Err.Raise vbObjectError + 11101, "OutlookAdapterTests", "Refusing to create a mail outside the test store."
    End If

    Set realDrafts = Application.Session.GetDefaultFolder(olFolderDrafts)
    draftsBefore = realDrafts.Items.Count

    pMailCounter = pMailCounter + 1
    path = Environ$("TEMP") & "\ninja_test_mail_" & Format$(Now, "hhnnss") & "_" & pMailCounter & ".msg"

    ' In-memory item -> file (never saved to a mailbox).
    Set draft = Application.CreateItem(olMailItem)
    draft.Subject = subject
    draft.Body = body
    draft.SaveAs path, olMSGUnicode
    Set draft = Nothing

    ' Open the file as an item of a temporary store, then move it into the test folder.
    Set opened = Application.Session.OpenSharedItem(path)
    pSenderSettable = True
    If Len(sender) > 0 Then pSenderSettable = SetMailProperty(opened, PR_SENDER_NAME, sender)
    opened.Save
    Set CreateMail = opened.Move(folder)

    Set opened = Nothing

    If realDrafts.Items.Count > draftsBefore Then
        pLeakDetected = True
        Err.Raise vbObjectError + 11103, "OutlookAdapterTests", "A test mail appeared in the real Drafts folder (delete it by hand: subject '" & subject & "')."
    End If

    On Error Resume Next
    Kill path
End Function

Private Function SetMailProperty(ByVal mail As Outlook.MailItem, ByVal tag As String, ByVal value As Variant) As Boolean
    On Error GoTo Failed
    mail.PropertyAccessor.SetProperty tag, value
    SetMailProperty = True
    Exit Function
Failed:
    SetMailProperty = False
End Function

Private Function Adapter(ByVal mail As Outlook.MailItem) As OutlookIncomingMail
    Set Adapter = New OutlookIncomingMail
    Adapter.Init mail
End Function

Private Function RawOf(ByVal folder As ITicketFolder) As Outlook.Folder
    Dim wrapper As OutlookTicketFolder

    Set wrapper = folder
    Set RawOf = wrapper.Raw
End Function

' Class, body length and closed flag of every item (for failure messages).
Private Function DescribeItems(ByVal folder As Outlook.Folder) As String
    Dim i As Long
    Dim item As Object

    On Error Resume Next
    DescribeItems = "items=" & folder.Items.Count
    For i = 1 To folder.Items.Count
        Set item = folder.Items(i)
        DescribeItems = DescribeItems & " [class=" & item.Class & " bodyLen=" & Len(item.Body) & _
            " closed=" & TicketParsing.IsClosedMail(item.Body) & " subject='" & item.Subject & "']"
    Next i
End Function

Private Function ChildOrNothing(ByVal parent As Outlook.Folder, ByVal folderName As String) As Outlook.Folder
    On Error Resume Next
    Set ChildOrNothing = parent.Folders(folderName)
End Function

' ---------------------------------------------------------------------------
'  Guard
' ---------------------------------------------------------------------------

Public Sub Test_Guard_AcceptsTestStoreAndRejectsRealMailbox()
    Setup
    TestAssert.AssertTrue OutlookTestGuard.IsTestStore(pTickets.Store), "test store"
    TestAssert.AssertFalse OutlookTestGuard.IsTestStore(Application.Session.GetDefaultFolder(olFolderInbox).Store), "real mailbox is rejected"
    TestAssert.AssertFalse OutlookTestGuard.IsTestStore(Nothing), "Nothing is rejected"
    Teardown
End Sub

' ---------------------------------------------------------------------------
'  OutlookTicketStore
' ---------------------------------------------------------------------------

Public Sub Test_Store_CreateFindAndList()
    Dim shortFolder As ITicketFolder
    Dim longFolder As ITicketFolder

    Setup
    Set shortFolder = pStore.CreateActiveFolder("#36 (short)")
    Set longFolder = pStore.CreateActiveFolder("#3621 (long)")

    TestAssert.AssertEqual "#36 (short)", pStore.FindActiveFolder("#36").Name, "#36 is not confused with #3621"
    TestAssert.AssertEqual "#3621 (long)", pStore.FindActiveFolder("#3621").Name
    TestAssert.AssertTrue pStore.FindActiveFolder("#99") Is Nothing, "unknown ticket"
    TestAssert.AssertEqual 2, pStore.ListActiveFolders().Count

    pStore.MoveToArchive shortFolder, DateSerial(2026, 10, 6)
    TestAssert.AssertEqual 1, pStore.ListActiveFolders().Count, "the archive folder is not listed as a ticket folder"
    TestAssert.AssertTrue pStore.FindActiveFolder("#36") Is Nothing, "archived folder is no longer active"
    Teardown
End Sub

Public Sub Test_Store_CreateActiveFolderReusesSameName()
    Setup
    pStore.CreateActiveFolder "#5 (x)"
    pStore.CreateActiveFolder "#5 (x)"

    TestAssert.AssertEqual 1, pTickets.Folders.Count
    Teardown
End Sub

Public Sub Test_Store_ArchivePathAndRestore()
    Dim folder As ITicketFolder
    Dim day As Date
    Dim inArchive As Outlook.Folder

    Setup
    day = DateSerial(2026, 10, 6)
    Set folder = pStore.CreateActiveFolder("#10 (x)")
    pStore.MoveToArchive folder, day

    Set inArchive = ChildOrNothing(ChildOrNothing(ChildOrNothing(pTickets, TicketConfig.FOLDER_ARCHIVE), "2026"), TicketParsing.ArchiveMonthName(day))
    TestAssert.AssertFalse inArchive Is Nothing, "Archiv\2026\<month> exists"
    TestAssert.AssertFalse ChildOrNothing(inArchive, "#10 (x)") Is Nothing, "folder is in Archiv\2026\<month>"
    TestAssert.AssertFalse pStore.FindArchivedFolder("#10") Is Nothing, "found in the archive"
    TestAssert.AssertTrue pStore.FindActiveFolder("#10") Is Nothing, "not active"

    pStore.RestoreFromArchive pStore.FindArchivedFolder("#10")
    TestAssert.AssertFalse pStore.FindActiveFolder("#10") Is Nothing, "active again"
    TestAssert.AssertTrue pStore.FindArchivedFolder("#10") Is Nothing, "no longer in the archive"
    Teardown
End Sub

Public Sub Test_Store_ArchiveWithoutDateUsesToday()
    Dim folder As ITicketFolder

    Setup
    Set folder = pStore.CreateActiveFolder("#11 (x)")
    pStore.MoveToArchive folder, 0

    TestAssert.AssertFalse ChildOrNothing(ChildOrNothing(ChildOrNothing(pTickets, TicketConfig.FOLDER_ARCHIVE), CStr(Year(Date))), TicketParsing.ArchiveMonthName(Date)) Is Nothing, "Archiv\<this year>\<this month>"
    Teardown
End Sub

Public Sub Test_Store_FindArchivedWithoutArchiveFolder()
    Setup
    TestAssert.AssertTrue pStore.FindArchivedFolder("#1") Is Nothing, "no archive folder yet"
    TestAssert.AssertEqual 0, pTickets.Folders.Count, "searching does not create the archive folder"
    Teardown
End Sub

' ---------------------------------------------------------------------------
'  OutlookTicketFolder
' ---------------------------------------------------------------------------

Public Sub Test_Folder_Rename()
    Dim folder As ITicketFolder

    Setup
    Set folder = pStore.CreateActiveFolder("#5 (old)")
    folder.Name = "#5 (new)"

    TestAssert.AssertEqual "#5 (new)", folder.Name
    TestAssert.AssertTrue ChildOrNothing(pTickets, "#5 (old)") Is Nothing, "old name is gone"
    TestAssert.AssertFalse ChildOrNothing(pTickets, "#5 (new)") Is Nothing, "new name exists"
    Teardown
End Sub

Public Sub Test_Folder_FindClosedStatusMail()
    Dim folder As ITicketFolder
    Dim found As Date

    Setup
    Set folder = pStore.CreateActiveFolder("#7 (x)")
    CreateMail RawOf(folder), "[example] (#7) Opened", "Status: Neu -> Offen"
    CreateMail RawOf(folder), "[example] (#7) Closed", "Hello" & vbCrLf & "Status: Offen -> Geschlossen"

    ' The receive time of a test mail cannot be chosen (the date logic is covered by the unit tests);
    ' here it only matters that the closing mail is found and a date comes back.
    TestAssert.AssertTrue folder.FindClosedStatusMail(found), "a mail says closed. " & DescribeItems(RawOf(folder))
    TestAssert.AssertTrue found > 0, "a receive time is returned"
    Teardown
End Sub

Public Sub Test_Folder_NoClosedStatusMail()
    Dim folder As ITicketFolder
    Dim found As Date

    Setup
    Set folder = pStore.CreateActiveFolder("#8 (x)")
    CreateMail RawOf(folder), "[example] (#8) Open", "Status: Neu -> Offen"

    TestAssert.AssertFalse folder.FindClosedStatusMail(found)
    Teardown
End Sub

' ---------------------------------------------------------------------------
'  OutlookIncomingMail
' ---------------------------------------------------------------------------

Public Sub Test_Mail_AdapterReadsSenderAndSubject()
    Dim mail As Outlook.MailItem
    Dim incoming As IIncomingMail

    Setup
    Set mail = CreateMail(pInbox, "[example] (#9) Subject", "body", TicketConfig.AllowedSender())
    Set incoming = Adapter(mail)

    TestAssert.AssertEqual "[example] (#9) Subject", incoming.Subject
    TestAssert.AssertTrue pSenderSettable, "the sender name of a test mail can be set (PR_SENDER_NAME)"
    TestAssert.AssertEqual TicketConfig.AllowedSender(), incoming.SenderName
    Teardown
End Sub

' The fields the matcher reads (M6). A mail created in the test PST has no transport headers,
' so it is not an automatic mail; the sender domain depends on how the test mail was created.
Public Sub Test_Mail_AdapterReadsMatchFields()
    Dim mail As Outlook.MailItem
    Dim incoming As IIncomingMail

    Setup
    Set mail = CreateMail(pInbox, "Match fields", "Ticket 4711 is meant", "Some Sender")
    Set incoming = Adapter(mail)

    TestAssert.AssertEqual "Ticket 4711 is meant", Trim$(Replace(incoming.BodyText, vbCrLf, ""))
    TestAssert.AssertFalse incoming.HasAutoHeader, "no transport headers"
    TestAssert.AssertTrue Len(incoming.ConversationId) > 0, "a conversation id exists"
    TestAssert.AssertEqual LCase$(incoming.SenderDomain), incoming.SenderDomain, "domain is lower case"
    TestAssert.AssertTrue incoming.ReceivedTime >= 0
    Teardown
End Sub

' The audit log finds a mail again by its message key, so the key must not change when the
' mail is moved into a ticket folder (the EntryID does change).
Public Sub Test_Mail_MessageKeySurvivesMove()
    Dim folder As ITicketFolder
    Dim mail As Outlook.MailItem
    Dim before As String
    Dim moved As Outlook.MailItem

    Setup
    Set mail = CreateMail(pInbox, "Key stays the same", "body")
    before = KeyOf(mail)
    TestAssert.AssertEqual 8, Len(before), "key is 8 hex digits"
    Set folder = pStore.CreateActiveFolder("#15 (key test)")

    pStore.MoveMailToFolder Adapter(mail), folder
    Set moved = RawOf(folder).Items(1)
    TestAssert.AssertEqual before, KeyOf(moved)
    Teardown
End Sub

' RecentMails finds the mails of a ticket folder received since a time (manual assignments, M7).
' A mail created in the test PST has no real receive time (Outlook reports a date far in the future for
' unsent items), so only the cut-off behaviour is checked.
Public Sub Test_Folder_RecentMails()
    Dim folder As ITicketFolder
    Dim mail As Outlook.MailItem
    Dim recent As Collection
    Dim incoming As IIncomingMail

    Setup
    Set folder = pStore.CreateActiveFolder("#16 (recent mails)")
    Set mail = CreateMail(RawOf(folder), "Recent mail", "body")

    Set recent = folder.RecentMails(Now - 1)
    TestAssert.AssertEqual 1, recent.Count
    Set incoming = recent(1)
    TestAssert.AssertEqual "Recent mail", incoming.Subject
    TestAssert.AssertEqual 0, folder.RecentMails(DateSerial(4600, 1, 1)).Count, "nothing received after year 4600"
    Teardown
End Sub

Public Sub Test_Mail_MoveToTicketFolder()
    Dim folder As ITicketFolder
    Dim mail As Outlook.MailItem

    Setup
    Set mail = CreateMail(pInbox, "[example] (#12) Move me", "body")
    TestAssert.AssertEqual 1, pInbox.Items.Count, "the test mail is in the test inbox"
    Set folder = pStore.CreateActiveFolder("#12 (x)")

    pStore.MoveMailToFolder Adapter(mail), folder
    TestAssert.AssertEqual 1, RawOf(folder).Items.Count, "mail is in the ticket folder"
    TestAssert.AssertEqual 0, pInbox.Items.Count, "mail left the inbox"
    Teardown
End Sub

Public Sub Test_NoLeakIntoRealMailbox()
    Dim realDrafts As Outlook.Folder
    Dim before As Long
    Dim after As Long

    Setup
    Set realDrafts = Application.Session.GetDefaultFolder(olFolderDrafts)
    before = realDrafts.Items.Count

    CreateMail pInbox, "[example] (#13) Leak check", "body", TicketConfig.AllowedSender()
    CreateMail pInbox, "[example] (#14) Leak check", "body"
    after = realDrafts.Items.Count

    TestAssert.AssertEqual before, after, "creating test mails must not add drafts to the real mailbox"
    TestAssert.AssertEqual 2, pInbox.Items.Count, "both mails are in the test inbox"
    Teardown
End Sub

' ---------------------------------------------------------------------------
'  Event-driven sorting (TicketMailWatcher, event Items.ItemAdd)
' ---------------------------------------------------------------------------

' Lets Outlook deliver events for up to the given seconds, or until condition becomes True.
Private Sub PumpEvents(ByVal seconds As Double, ByVal waitForActiveFolder As Boolean)
    Dim started As Double

    started = Timer
    Do While Timer - started < seconds
        DoEvents
        If waitForActiveFolder Then
            If pStore.ListActiveFolders().Count > 0 Then Exit Do
        End If
    Loop
End Sub

Public Sub Test_Watcher_SortsArrivingMailWithoutManualRun()
    Dim fake As FakeNinjaTransport
    Dim watcher As TicketMailWatcher

    Setup
    Set fake = New FakeNinjaTransport
    fake.AddResponse "ticketing/ticket/4712", 200, Fixtures.TicketOpen()
    NinjaApi.UseTransport fake

    Set watcher = New TicketMailWatcher
    watcher.StartWatching pInbox, pStore, 0
    TestAssert.AssertTrue watcher.IsWatching

    CreateMail pInbox, "[example] (#4712) Event mail", "body", TicketConfig.AllowedSender()
    PumpEvents 10, True
    watcher.StopWatching

    TestAssert.AssertEqual 1, pStore.ListActiveFolders().Count, "the arriving mail was sorted by the event, no macro was started"
    TestAssert.AssertEqual "#4712 (Example open ticket)", pStore.FindActiveFolder("#4712").Name
    TestAssert.AssertEqual 0, pInbox.Items.Count, "the inbox is empty again"
    Teardown
End Sub

Public Sub Test_Watcher_LeavesMailsOfOtherSendersAlone()
    Dim watcher As TicketMailWatcher

    Setup
    Set watcher = New TicketMailWatcher
    watcher.StartWatching pInbox, pStore, 0

    CreateMail pInbox, "[example] (#4713) Not from the ticket system", "body", "someone else"
    PumpEvents 3, False
    watcher.StopWatching

    TestAssert.AssertEqual 0, pStore.ListActiveFolders().Count, "no folder created"
    TestAssert.AssertEqual 1, pInbox.Items.Count, "the mail stays in the inbox"
    Teardown
End Sub

Public Sub Test_Watcher_StopsWhenStopped()
    Dim watcher As TicketMailWatcher

    Setup
    Set watcher = New TicketMailWatcher
    watcher.StartWatching pInbox, pStore, 0
    watcher.StopWatching
    TestAssert.AssertFalse watcher.IsWatching

    CreateMail pInbox, "[example] (#4714) After stop", "body", TicketConfig.AllowedSender()
    PumpEvents 3, False

    TestAssert.AssertEqual 0, pStore.ListActiveFolders().Count, "a stopped watcher does nothing"
    TestAssert.AssertEqual 1, pInbox.Items.Count
    Teardown
End Sub

' ---------------------------------------------------------------------------
'  Progress window (non-modal ProgressForm)
' ---------------------------------------------------------------------------

Public Sub Test_Progress_FormOpensAndClosesWithTheSink()
    Dim sink As IProgressSink

    ProgressHelper.ProgressEnd
    Set sink = New ProgressSinkForm
    TestAssert.AssertFalse ProgressHelper.IsActive(), "no form before the run"

    sink.Begin 3
    TestAssert.AssertTrue ProgressHelper.IsActive(), "the form is open (non-modal)"
    sink.Advance "one"
    sink.Advance "two"
    sink.Advance "three"
    sink.Finish

    TestAssert.AssertFalse ProgressHelper.IsActive(), "the form is closed after Finish"
    ProgressHelper.ProgressEnd
End Sub

Public Sub Test_Progress_ArchiveRunClosesTheForm()
    Dim fake As FakeNinjaTransport

    Setup
    Set fake = New FakeNinjaTransport
    fake.AddResponse "ticketing/ticket/20102", 200, Fixtures.TicketClosed()
    fake.AddResponse "ticketing/ticket/20102/log-entry?type=SAVE", 200, Fixtures.LogClosedManual()
    NinjaApi.UseTransport fake
    pStore.CreateActiveFolder "#20102 (Example)"

    TestAssert.AssertEqual 1, TicketService.ArchiveClosedTickets(pStore, True, New ProgressSinkForm), "folder archived"
    TestAssert.AssertFalse ProgressHelper.IsActive(), "the form is closed after the run"
    ProgressHelper.ProgressEnd
    Teardown
End Sub

' ---------------------------------------------------------------------------
'  End to end: TicketService with the real adapters and a fake NinjaOne
' ---------------------------------------------------------------------------

Public Sub Test_EndToEnd_MailArchiveRestore()
    Dim fake As FakeNinjaTransport
    Dim mail As Outlook.MailItem
    Dim active As OutlookTicketFolder

    Setup
    Set fake = New FakeNinjaTransport
    NinjaApi.UseTransport fake

    ' 1. A ticket mail arrives: folder is created and named after the current ticket subject.
    fake.AddResponse "ticketing/ticket/4711", 200, Fixtures.TicketOpen()
    Set mail = CreateMail(pInbox, "[example] (#4711) First mail", "hello", TicketConfig.AllowedSender())
    TestAssert.AssertTrue pSenderSettable, "test mails need a settable sender name"
    TestAssert.AssertTrue TicketService.ProcessMail(Adapter(mail), pStore), "mail was processed"

    TestAssert.AssertEqual 1, pStore.ListActiveFolders().Count
    TestAssert.AssertEqual "#4711 (Example open ticket)", pStore.FindActiveFolder("#4711").Name, "folder named after the API subject"
    TestAssert.AssertEqual 0, pInbox.Items.Count, "inbox is empty again"

    ' 2. The ticket is closed in NinjaOne: the archive run moves the folder.
    fake.AddResponse "ticketing/ticket/4711", 200, Fixtures.TicketClosed()
    fake.AddResponse "ticketing/ticket/4711/log-entry?type=SAVE", 200, Fixtures.LogClosedManual()
    TestAssert.AssertEqual 1, TicketService.ArchiveClosedTickets(pStore, True), "one folder archived"
    TestAssert.AssertEqual 0, pStore.ListActiveFolders().Count
    TestAssert.AssertFalse pStore.FindArchivedFolder("#4711") Is Nothing, "folder is in the archive"

    ' 3. A new mail for the archived ticket: the folder comes back.
    fake.AddResponse "ticketing/ticket/4711", 200, Fixtures.TicketOpen()
    Set mail = CreateMail(pInbox, "[example] (#4711) Second mail", "again", TicketConfig.AllowedSender())
    TestAssert.AssertTrue TicketService.ProcessMail(Adapter(mail), pStore), "second mail was processed"

    TestAssert.AssertEqual 1, pStore.ListActiveFolders().Count, "folder restored"
    TestAssert.AssertTrue pStore.FindArchivedFolder("#4711") Is Nothing, "no longer archived"
    Set active = pStore.FindActiveFolder("#4711")
    TestAssert.AssertEqual 2, RawOf(active).Items.Count, "both mails are in the folder"
    Teardown
End Sub

Private Function KeyOf(ByVal mail As Outlook.MailItem) As String
    Dim incoming As IIncomingMail

    Set incoming = Adapter(mail)
    KeyOf = incoming.MessageKey
End Function
