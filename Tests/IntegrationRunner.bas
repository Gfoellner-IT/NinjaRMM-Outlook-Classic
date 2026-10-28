Attribute VB_Name = "IntegrationRunner"
' GENERATED FILE - do not edit. Regenerate with: tools\New-TestRunnerModule.ps1
'
' Integration tests of the Outlook adapters. They use Outlook types and run ONLY in the
' test PST (see OutlookTestGuard); they cannot run headless. Start RunIntegrationTests
' from the VBA editor (F5). Result: Immediate window and %TEMP%\ninja_integration_results.txt.

Option Explicit

Public Const INTEGRATION_RESULT_FILE_NAME As String = "ninja_integration_results.txt"

Public Sub RunIntegrationTests()
    TestAssert.Reset
    TicketLog.DisableFileLogging
    On Error Resume Next

    TestAssert.BeginTest "OutlookAdapterTests.Test_Guard_AcceptsTestStoreAndRejectsRealMailbox"
    OutlookAdapterTests.Test_Guard_AcceptsTestStoreAndRejectsRealMailbox
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Store_CreateFindAndList"
    OutlookAdapterTests.Test_Store_CreateFindAndList
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Store_CreateActiveFolderReusesSameName"
    OutlookAdapterTests.Test_Store_CreateActiveFolderReusesSameName
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Store_ArchivePathAndRestore"
    OutlookAdapterTests.Test_Store_ArchivePathAndRestore
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Store_ArchiveWithoutDateUsesToday"
    OutlookAdapterTests.Test_Store_ArchiveWithoutDateUsesToday
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Store_FindArchivedWithoutArchiveFolder"
    OutlookAdapterTests.Test_Store_FindArchivedWithoutArchiveFolder
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Folder_Rename"
    OutlookAdapterTests.Test_Folder_Rename
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Folder_FindClosedStatusMail"
    OutlookAdapterTests.Test_Folder_FindClosedStatusMail
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Folder_NoClosedStatusMail"
    OutlookAdapterTests.Test_Folder_NoClosedStatusMail
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Mail_AdapterReadsSenderAndSubject"
    OutlookAdapterTests.Test_Mail_AdapterReadsSenderAndSubject
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Mail_AdapterReadsMatchFields"
    OutlookAdapterTests.Test_Mail_AdapterReadsMatchFields
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Mail_MessageKeySurvivesMove"
    OutlookAdapterTests.Test_Mail_MessageKeySurvivesMove
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Folder_RecentMails"
    OutlookAdapterTests.Test_Folder_RecentMails
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Mail_MoveToTicketFolder"
    OutlookAdapterTests.Test_Mail_MoveToTicketFolder
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_NoLeakIntoRealMailbox"
    OutlookAdapterTests.Test_NoLeakIntoRealMailbox
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Watcher_SortsArrivingMailWithoutManualRun"
    OutlookAdapterTests.Test_Watcher_SortsArrivingMailWithoutManualRun
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Watcher_LeavesMailsOfOtherSendersAlone"
    OutlookAdapterTests.Test_Watcher_LeavesMailsOfOtherSendersAlone
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Watcher_StopsWhenStopped"
    OutlookAdapterTests.Test_Watcher_StopsWhenStopped
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Progress_FormOpensAndClosesWithTheSink"
    OutlookAdapterTests.Test_Progress_FormOpensAndClosesWithTheSink
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_Progress_ArchiveRunClosesTheForm"
    OutlookAdapterTests.Test_Progress_ArchiveRunClosesTheForm
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "OutlookAdapterTests.Test_EndToEnd_MailArchiveRestore"
    OutlookAdapterTests.Test_EndToEnd_MailArchiveRestore
    TestAssert.EndTest Err.Number, Err.Description

    On Error GoTo 0
    Publish "Integration tests (test PST NinjaTest)", IIf(TestAssert.AllPassed, "RESULT: OK", "RESULT: FAILED")
End Sub

Private Sub Publish(ByVal title As String, ByVal resultLine As String)
    Dim text As String
    Dim fso As Object
    Dim stream As Object

    text = title & " - " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
           TestAssert.Report() & _
           "Tests passed: " & TestAssert.PassedTests & _
           ", failed: " & TestAssert.FailedTests & _
           ", assertions: " & TestAssert.AssertionCount & vbCrLf & _
           resultLine & vbCrLf

    Debug.Print text
    On Error GoTo Done
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(Environ$("TEMP") & "\" & INTEGRATION_RESULT_FILE_NAME, True, False)
    stream.Write text
    stream.Close
Done:
End Sub
