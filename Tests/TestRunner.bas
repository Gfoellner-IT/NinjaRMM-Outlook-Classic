Attribute VB_Name = "TestRunner"
' GENERATED FILE - do not edit. Regenerate with: tools\New-TestRunnerModule.ps1
'
' Entry points for all unit tests. Start RunAllTests from the VBA editor (F5)
' or headless with tools\Invoke-VbaTests.ps1.
' Results go to the Immediate window and to %TEMP%\ninja_test_results.txt.
'
' Outlook has no Application.Run, so every test is called explicitly. The calls are
' generated from the "Public Sub Test_*" procedures of the test modules.
'
' Unit tests must not touch the mailbox or the network. Tests that use the
' Windows Credential Manager only use unique, temporary entries.

Option Explicit

Public Const RESULT_FILE_NAME As String = "ninja_test_results.txt"

Public Sub RunAllTests()
    TestAssert.Reset

    RunSelfTests
    RunAuditReportTests
    RunAuditTests
    RunJsonLiteTests
    RunLearnedRulesTests
    RunLocalSettingsTests
    RunManualScanTests
    RunMatchLevelsTests
    RunMatchPromptTextTests
    RunMatchReviewTests
    RunNinjaApiTests
    RunNinjaAuthTests
    RunNinjaCredentialStoreTests
    RunNinjaTicketIndexTests
    RunNinjaTicketLogicTests
    RunNinjaTokenLogicTests
    RunReviewFormTests
    RunTicketDecisionTests
    RunTicketMatcherTests
    RunTicketMatchServiceTests
    RunTicketParsingTests
    RunTicketServiceTests
    RunTimeUtilTests

    PublishResults "Unit tests", IIf(TestAssert.AllPassed, "RESULT: OK", "RESULT: FAILED")
End Sub

' Opt-in: real requests against NinjaOne (needs network and the Credential Manager entry).
Public Sub RunLiveTests()
    TestAssert.Reset
    On Error Resume Next

    TestAssert.BeginTest "LiveTests.Test_Live_TokenFromRealEndpoint"
    LiveTests.Test_Live_TokenFromRealEndpoint
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LiveTests.Test_Live_TicketStates"
    LiveTests.Test_Live_TicketStates
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LiveTests.Test_Live_TicketIndex"
    LiveTests.Test_Live_TicketIndex
    TestAssert.EndTest Err.Number, Err.Description

    On Error GoTo 0
    PublishResults "Live tests", IIf(TestAssert.AllPassed, "RESULT: OK", "RESULT: FAILED")
End Sub

' Proves that the runner really reports failures: two deliberately failing
' tests and one passing test must give exactly "2 failed, 1 passed".
Public Sub RunFailureProbe()
    Dim ok As Boolean

    TestAssert.Reset
    On Error Resume Next

    TestAssert.BeginTest "Probe.AssertionFails"
    SelfTests.Probe_AssertionFails
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "Probe.RaisesError"
    SelfTests.Probe_RaisesError
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "Probe.Passes"
    SelfTests.Probe_Passes
    TestAssert.EndTest Err.Number, Err.Description

    On Error GoTo 0
    ok = (TestAssert.FailedTests = 2 And TestAssert.PassedTests = 1)
    PublishResults "Failure probe (expected: 2 failed, 1 passed)", IIf(ok, "RESULT: PROBE OK", "RESULT: PROBE FAILED")
End Sub

' ---------------------------------------------------------------------------
'  Suites
' ---------------------------------------------------------------------------

Private Sub RunSelfTests()
    On Error Resume Next

    TestAssert.BeginTest "SelfTests.Test_AreEqual_Strings"
    SelfTests.Test_AreEqual_Strings
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "SelfTests.Test_AreEqual_Numbers"
    SelfTests.Test_AreEqual_Numbers
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "SelfTests.Test_AreEqual_NullEmptyObject"
    SelfTests.Test_AreEqual_NullEmptyObject
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "SelfTests.Test_Describe"
    SelfTests.Test_Describe
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "SelfTests.Test_Fixtures_Available"
    SelfTests.Test_Fixtures_Available
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunAuditReportTests()
    On Error Resume Next

    TestAssert.BeginTest "AuditReportTests.Test_GermanText_Markers"
    AuditReportTests.Test_GermanText_Markers
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_SubjectStem"
    AuditReportTests.Test_SubjectStem
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_RecommendLevel"
    AuditReportTests.Test_RecommendLevel
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_ParseLines_SkipsBrokenLines"
    AuditReportTests.Test_ParseLines_SkipsBrokenLines
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_Report_RulesTableAndRecommendation"
    AuditReportTests.Test_Report_RulesTableAndRecommendation
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_Report_Overview"
    AuditReportTests.Test_Report_Overview
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_Report_Hotspots"
    AuditReportTests.Test_Report_Hotspots
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_Report_Candidates"
    AuditReportTests.Test_Report_Candidates
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_Report_CandidateCheckAgainstLaterDecisions"
    AuditReportTests.Test_Report_CandidateCheckAgainstLaterDecisions
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_Report_MailsWithoutMatch"
    AuditReportTests.Test_Report_MailsWithoutMatch
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_Report_ContainsNoMailTextsOrFullSubjects"
    AuditReportTests.Test_Report_ContainsNoMailTextsOrFullSubjects
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_Report_EmptyLog"
    AuditReportTests.Test_Report_EmptyLog
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_LoadRecords_ReadsMonthlyFilesOldestFirst"
    AuditReportTests.Test_LoadRecords_ReadsMonthlyFilesOldestFirst
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_WriteReport_Utf8AndNames"
    AuditReportTests.Test_WriteReport_Utf8AndNames
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_ReportNamesAndPrompt"
    AuditReportTests.Test_ReportNamesAndPrompt
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditReportTests.Test_Report_RepeatedDecisionsAboutTheSameMailCountOnce"
    AuditReportTests.Test_Report_RepeatedDecisionsAboutTheSameMailCountOnce
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunAuditTests()
    On Error Resume Next

    TestAssert.BeginTest "AuditTests.Test_Hash32_KnownVectors"
    AuditTests.Test_Hash32_KnownVectors
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_Hash32_StableAndDistinct"
    AuditTests.Test_Hash32_StableAndDistinct
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_JsonString"
    AuditTests.Test_JsonString
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_BuildMatchEvent"
    AuditTests.Test_BuildMatchEvent
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_BuildMatchEvent_IsSingleLineAscii"
    AuditTests.Test_BuildMatchEvent_IsSingleLineAscii
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_MonthFileName"
    AuditTests.Test_MonthFileName
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_IsExpiredFileName"
    AuditTests.Test_IsExpiredFileName
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_FileSink_AppendsLinesAndCreatesFolders"
    AuditTests.Test_FileSink_AppendsLinesAndCreatesFolders
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_FileSink_NewFilePerMonth"
    AuditTests.Test_FileSink_NewFilePerMonth
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_FileSink_PrunesOnlyExpiredAuditFiles"
    AuditTests.Test_FileSink_PrunesOnlyExpiredAuditFiles
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_BuildLine_HasFeaturesButNoNamesOrBody"
    AuditTests.Test_BuildLine_HasFeaturesButNoNamesOrBody
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_Privacy_NoSubjectWithoutAMatch_OnlyItsHash"
    AuditTests.Test_Privacy_NoSubjectWithoutAMatch_OnlyItsHash
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_Privacy_SubjectIsLoggedForFoundAmbiguousAndUserDecisions"
    AuditTests.Test_Privacy_SubjectIsLoggedForFoundAmbiguousAndUserDecisions
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_Privacy_FlowWritesNoSubjectForNoMatch"
    AuditTests.Test_Privacy_FlowWritesNoSubjectForNoMatch
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_BuildLine_EmptyIdsGiveEmptyFields"
    AuditTests.Test_BuildLine_EmptyIdsGiveEmptyFields
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_Flow_OneLinePerDecision"
    AuditTests.Test_Flow_OneLinePerDecision
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_Flow_ActionMovedAndStays"
    AuditTests.Test_Flow_ActionMovedAndStays
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_Flow_TicketSenderIsNotAudited"
    AuditTests.Test_Flow_TicketSenderIsNotAudited
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_Flow_BrokenAuditLogNeverStopsTheMatching"
    AuditTests.Test_Flow_BrokenAuditLogNeverStopsTheMatching
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "AuditTests.Test_Flow_NoFileIsWrittenWhileFileLoggingIsDisabled"
    AuditTests.Test_Flow_NoFileIsWrittenWhileFileLoggingIsDisabled
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunJsonLiteTests()
    On Error Resume Next

    TestAssert.BeginTest "JsonLiteTests.Test_Json_Scalars"
    JsonLiteTests.Test_Json_Scalars
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "JsonLiteTests.Test_Json_NumberTypes"
    JsonLiteTests.Test_Json_NumberTypes
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "JsonLiteTests.Test_Json_StringEscapes"
    JsonLiteTests.Test_Json_StringEscapes
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "JsonLiteTests.Test_Json_NestedStructures"
    JsonLiteTests.Test_Json_NestedStructures
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "JsonLiteTests.Test_Json_EmptyContainersAndWhitespace"
    JsonLiteTests.Test_Json_EmptyContainersAndWhitespace
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "JsonLiteTests.Test_Json_ArrayOfObjects"
    JsonLiteTests.Test_Json_ArrayOfObjects
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "JsonLiteTests.Test_Json_KeysAreCaseSensitiveAndLastDuplicateWins"
    JsonLiteTests.Test_Json_KeysAreCaseSensitiveAndLastDuplicateWins
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "JsonLiteTests.Test_Json_InvalidInput"
    JsonLiteTests.Test_Json_InvalidInput
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "JsonLiteTests.Test_Json_TryParse"
    JsonLiteTests.Test_Json_TryParse
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "JsonLiteTests.Test_Json_ParserCanBeReused"
    JsonLiteTests.Test_Json_ParserCanBeReused
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "JsonLiteTests.Test_Json_PathAccess"
    JsonLiteTests.Test_Json_PathAccess
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "JsonLiteTests.Test_Json_LargeDocument"
    JsonLiteTests.Test_Json_LargeDocument
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunLearnedRulesTests()
    On Error Resume Next

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_StemStartsWith"
    LearnedRulesTests.Test_Learned_StemStartsWith
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_LoadsOnlyValidEnabledRules"
    LearnedRulesTests.Test_Learned_LoadsOnlyValidEnabledRules
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_MatchingStem"
    LearnedRulesTests.Test_Learned_MatchingStem
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_BrokenFileMeansNoRules"
    LearnedRulesTests.Test_Learned_BrokenFileMeansNoRules
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_NoRulesFileInTests"
    LearnedRulesTests.Test_Learned_NoRulesFileInTests
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_MatchesTheSingleOpenTicketOfTheGroup"
    LearnedRulesTests.Test_Learned_MatchesTheSingleOpenTicketOfTheGroup
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_TwoOpenTicketsAreAmbiguous"
    LearnedRulesTests.Test_Learned_TwoOpenTicketsAreAmbiguous
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_NoMatchWithoutTicketDomainOrOpenState"
    LearnedRulesTests.Test_Learned_NoMatchWithoutTicketDomainOrOpenState
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_OtherRulesComeFirst"
    LearnedRulesTests.Test_Learned_OtherRulesComeFirst
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_SwitchedOffLevelSkipsTheRule"
    LearnedRulesTests.Test_Learned_SwitchedOffLevelSkipsTheRule
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_DefaultLevelIsLogOnly"
    LearnedRulesTests.Test_Learned_DefaultLevelIsLogOnly
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_ProposalAtLevel2"
    LearnedRulesTests.Test_Learned_ProposalAtLevel2
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_IndexFindsOpenTicketsByStem"
    LearnedRulesTests.Test_Learned_IndexFindsOpenTicketsByStem
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LearnedRulesTests.Test_Learned_ReportProposesRulesForTheRulesFile"
    LearnedRulesTests.Test_Learned_ReportProposesRulesForTheRulesFile
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunLocalSettingsTests()
    On Error Resume Next

    TestAssert.BeginTest "LocalSettingsTests.Test_Settings_Defaults"
    LocalSettingsTests.Test_Settings_Defaults
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LocalSettingsTests.Test_Settings_ReadsUnicodeAndNormalizesBase"
    LocalSettingsTests.Test_Settings_ReadsUnicodeAndNormalizesBase
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LocalSettingsTests.Test_Settings_RejectsUnsafeApiBase"
    LocalSettingsTests.Test_Settings_RejectsUnsafeApiBase
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LocalSettingsTests.Test_Settings_RejectsInvalidBoard"
    LocalSettingsTests.Test_Settings_RejectsInvalidBoard
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "LocalSettingsTests.Test_Settings_UnitDefaultsIgnoreFile"
    LocalSettingsTests.Test_Settings_UnitDefaultsIgnoreFile
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunManualScanTests()
    On Error Resume Next

    TestAssert.BeginTest "ManualScanTests.Test_Scan_AgreeWhenTheMatcherWouldHaveChosenTheSameTicket"
    ManualScanTests.Test_Scan_AgreeWhenTheMatcherWouldHaveChosenTheSameTicket
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_DifferWhenTheMatcherWouldHaveChosenAnotherTicket"
    ManualScanTests.Test_Scan_DifferWhenTheMatcherWouldHaveChosenAnotherTicket
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_MissedWhenTheMatcherFindsNothing"
    ManualScanTests.Test_Scan_MissedWhenTheMatcherFindsNothing
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_ExcludedMailsAreMarked"
    ManualScanTests.Test_Scan_ExcludedMailsAreMarked
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_AllRulesOffMeansUnchecked"
    ManualScanTests.Test_Scan_AllRulesOffMeansUnchecked
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_EveryMailIsLoggedOnce"
    ManualScanTests.Test_Scan_EveryMailIsLoggedOnce
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_KnownMailsAreSkipped"
    ManualScanTests.Test_Scan_KnownMailsAreSkipped
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_TicketSenderMailsAreNotManual"
    ManualScanTests.Test_Scan_TicketSenderMailsAreNotManual
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_MailsWithoutMessageKeyAreSkipped"
    ManualScanTests.Test_Scan_MailsWithoutMessageKeyAreSkipped
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_OnlyRecentMails"
    ManualScanTests.Test_Scan_OnlyRecentMails
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_FoldersThatAreNoTicketFoldersAreSkipped"
    ManualScanTests.Test_Scan_FoldersThatAreNoTicketFoldersAreSkipped
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_ABrokenFolderDoesNotStopTheScan"
    ManualScanTests.Test_Scan_ABrokenFolderDoesNotStopTheScan
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_RemembersTheConversationForTheMatcher"
    ManualScanTests.Test_Scan_RemembersTheConversationForTheMatcher
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Scan_SeenKeysFromRecords"
    ManualScanTests.Test_Scan_SeenKeysFromRecords
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Report_ShowsManualAssignments"
    ManualScanTests.Test_Report_ShowsManualAssignments
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ManualScanTests.Test_Report_NoManualAssignmentsYet"
    ManualScanTests.Test_Report_NoManualAssignmentsYet
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunMatchLevelsTests()
    On Error Resume Next

    TestAssert.BeginTest "MatchLevelsTests.Test_Levels_DefaultsAreLogOnly"
    MatchLevelsTests.Test_Levels_DefaultsAreLogOnly
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Levels_PerRuleAndCap"
    MatchLevelsTests.Test_Levels_PerRuleAndCap
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Levels_OutOfRangeValuesAreClamped"
    MatchLevelsTests.Test_Levels_OutOfRangeValuesAreClamped
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Levels_Names"
    MatchLevelsTests.Test_Levels_Names
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Off_NinjaIdRuleIsSkipped"
    MatchLevelsTests.Test_Off_NinjaIdRuleIsSkipped
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Off_ExternalNumberRuleDoesNotStopTheOthers"
    MatchLevelsTests.Test_Off_ExternalNumberRuleDoesNotStopTheOthers
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Off_ConversationAndSubjectRules"
    MatchLevelsTests.Test_Off_ConversationAndSubjectRules
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Off_AllRulesOffDoesNothingAtAll"
    MatchLevelsTests.Test_Off_AllRulesOffDoesNothingAtAll
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Level1_OnlyLogs"
    MatchLevelsTests.Test_Level1_OnlyLogs
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Level2_QueuesAProposal"
    MatchLevelsTests.Test_Level2_QueuesAProposal
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Level3_MovesWithoutAsking"
    MatchLevelsTests.Test_Level3_MovesWithoutAsking
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_DryRunNeverActs_WhateverTheLevel"
    MatchLevelsTests.Test_DryRunNeverActs_WhateverTheLevel
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Level_IsPerRule"
    MatchLevelsTests.Test_Level_IsPerRule
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Cap_LowersTheLevel"
    MatchLevelsTests.Test_Cap_LowersTheLevel
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_ClosedTicketIsNeverProposedOrMoved"
    MatchLevelsTests.Test_ClosedTicketIsNeverProposedOrMoved
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_AmbiguousIsNeverProposedOrMoved"
    MatchLevelsTests.Test_AmbiguousIsNeverProposedOrMoved
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_TicketSenderIsNotProposed"
    MatchLevelsTests.Test_TicketSenderIsNotProposed
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Proposals_SameMailIsNotAddedTwice"
    MatchLevelsTests.Test_Proposals_SameMailIsNotAddedTwice
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Proposals_RemoveAndClear"
    MatchLevelsTests.Test_Proposals_RemoveAndClear
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchLevelsTests.Test_Levels_AnyRuleAtLevel"
    MatchLevelsTests.Test_Levels_AnyRuleAtLevel
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunMatchPromptTextTests()
    On Error Resume Next

    TestAssert.BeginTest "MatchPromptTextTests.Test_PromptText_Hint"
    MatchPromptTextTests.Test_PromptText_Hint
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchPromptTextTests.Test_PromptText_Proposal"
    MatchPromptTextTests.Test_PromptText_Proposal
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchPromptTextTests.Test_PromptText_ProposalWithoutSummaryAndLongSubject"
    MatchPromptTextTests.Test_PromptText_ProposalWithoutSummaryAndLongSubject
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchPromptTextTests.Test_PromptText_RuleLabels"
    MatchPromptTextTests.Test_PromptText_RuleLabels
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchPromptTextTests.Test_PromptText_AnswersToDecisions"
    MatchPromptTextTests.Test_PromptText_AnswersToDecisions
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchPromptTextTests.Test_PromptText_OtherTexts"
    MatchPromptTextTests.Test_PromptText_OtherTexts
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunMatchReviewTests()
    On Error Resume Next

    TestAssert.BeginTest "MatchReviewTests.Test_Review_AcceptMovesTheMailAndDropsTheProposal"
    MatchReviewTests.Test_Review_AcceptMovesTheMailAndDropsTheProposal
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Review_AcceptUsesTheExistingFolder"
    MatchReviewTests.Test_Review_AcceptUsesTheExistingFolder
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Review_AcceptWhenTheTicketWasClosedMeanwhile"
    MatchReviewTests.Test_Review_AcceptWhenTheTicketWasClosedMeanwhile
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Review_AcceptWhenTheMailCannotBeMovedAnyMore"
    MatchReviewTests.Test_Review_AcceptWhenTheMailCannotBeMovedAnyMore
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Review_RejectRemembersTheMail"
    MatchReviewTests.Test_Review_RejectRemembersTheMail
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Review_RejectedMailIsNotProposedAgain"
    MatchReviewTests.Test_Review_RejectedMailIsNotProposedAgain
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Rejections_SurviveARestart"
    MatchReviewTests.Test_Rejections_SurviveARestart
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Rejections_EmptyKeyAndBrokenStore"
    MatchReviewTests.Test_Rejections_EmptyKeyAndBrokenStore
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Rejections_FileStoreRoundTrip"
    MatchReviewTests.Test_Rejections_FileStoreRoundTrip
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Review_SkipLeavesTheProposal"
    MatchReviewTests.Test_Review_SkipLeavesTheProposal
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Review_CancelDecidesNothing"
    MatchReviewTests.Test_Review_CancelDecidesNothing
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Review_MixedDecisionsInOrder"
    MatchReviewTests.Test_Review_MixedDecisionsInOrder
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Review_MissingOrInvalidDecisionsMeanSkip"
    MatchReviewTests.Test_Review_MissingOrInvalidDecisionsMeanSkip
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Review_NothingPendingDoesNotAsk"
    MatchReviewTests.Test_Review_NothingPendingDoesNotAsk
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Offer_AsksOnceForNewProposals"
    MatchReviewTests.Test_Offer_AsksOnceForNewProposals
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Offer_YesStartsTheReview"
    MatchReviewTests.Test_Offer_YesStartsTheReview
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Offer_NothingPendingNeverAsks"
    MatchReviewTests.Test_Offer_NothingPendingNeverAsks
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Offer_NewProposalsWaitForTheCooldown"
    MatchReviewTests.Test_Offer_NewProposalsWaitForTheCooldown
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_Offer_NoSecondDialogWhileTheFirstIsOpen"
    MatchReviewTests.Test_Offer_NoSecondDialogWhileTheFirstIsOpen
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "MatchReviewTests.Test_EndToEnd_Level2ThenAccept"
    MatchReviewTests.Test_EndToEnd_Level2ThenAccept
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunNinjaApiTests()
    On Error Resume Next

    TestAssert.BeginTest "NinjaApiTests.Test_Api_StateOpen"
    NinjaApiTests.Test_Api_StateOpen
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaApiTests.Test_Api_StateClosed"
    NinjaApiTests.Test_Api_StateClosed
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaApiTests.Test_Api_NotFoundCountsAsClosedButIsDistinguishable"
    NinjaApiTests.Test_Api_NotFoundCountsAsClosedButIsDistinguishable
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaApiTests.Test_Api_UnavailableIsNeverClosed"
    NinjaApiTests.Test_Api_UnavailableIsNeverClosed
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaApiTests.Test_Api_UnknownPathIsUnavailable"
    NinjaApiTests.Test_Api_UnknownPathIsUnavailable
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaApiTests.Test_Api_InvalidTicketIdMakesNoRequest"
    NinjaApiTests.Test_Api_InvalidTicketIdMakesNoRequest
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaApiTests.Test_Api_Subject"
    NinjaApiTests.Test_Api_Subject
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaApiTests.Test_Api_ClosedDate"
    NinjaApiTests.Test_Api_ClosedDate
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaApiTests.Test_Api_ClosedDateUnknown"
    NinjaApiTests.Test_Api_ClosedDateUnknown
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaApiTests.Test_Api_OneRequestPerQuery"
    NinjaApiTests.Test_Api_OneRequestPerQuery
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunNinjaAuthTests()
    On Error Resume Next

    TestAssert.BeginTest "NinjaAuthTests.Test_GetAccessToken_Success"
    NinjaAuthTests.Test_GetAccessToken_Success
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaAuthTests.Test_GetAccessToken_UsesCache"
    NinjaAuthTests.Test_GetAccessToken_UsesCache
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaAuthTests.Test_GetAccessToken_RefetchAfterInvalidate"
    NinjaAuthTests.Test_GetAccessToken_RefetchAfterInvalidate
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaAuthTests.Test_GetAccessToken_ShortLivedTokenIsNotCached"
    NinjaAuthTests.Test_GetAccessToken_ShortLivedTokenIsNotCached
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaAuthTests.Test_GetAccessToken_NoCredential"
    NinjaAuthTests.Test_GetAccessToken_NoCredential
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaAuthTests.Test_GetAccessToken_AuthRejected"
    NinjaAuthTests.Test_GetAccessToken_AuthRejected
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaAuthTests.Test_GetAccessToken_NoResponseIsUnavailable"
    NinjaAuthTests.Test_GetAccessToken_NoResponseIsUnavailable
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaAuthTests.Test_GetAccessToken_ServerErrorIsUnavailable"
    NinjaAuthTests.Test_GetAccessToken_ServerErrorIsUnavailable
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaAuthTests.Test_GetAccessToken_InvalidResponse"
    NinjaAuthTests.Test_GetAccessToken_InvalidResponse
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunNinjaCredentialStoreTests()
    On Error Resume Next

    TestAssert.BeginTest "NinjaCredentialStoreTests.Test_Credential_RoundTrip"
    NinjaCredentialStoreTests.Test_Credential_RoundTrip
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaCredentialStoreTests.Test_Credential_Overwrite"
    NinjaCredentialStoreTests.Test_Credential_Overwrite
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaCredentialStoreTests.Test_Credential_Missing"
    NinjaCredentialStoreTests.Test_Credential_Missing
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaCredentialStoreTests.Test_Credential_EmptySecretRejected"
    NinjaCredentialStoreTests.Test_Credential_EmptySecretRejected
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunNinjaTicketIndexTests()
    On Error Resume Next

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_BoardBody"
    NinjaTicketIndexTests.Test_BoardBody
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_BoardBody_EscapesValue"
    NinjaTicketIndexTests.Test_BoardBody_EscapesValue
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_ParseBoardRun_Page"
    NinjaTicketIndexTests.Test_ParseBoardRun_Page
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_ParseBoardRun_ClosedAndResolved"
    NinjaTicketIndexTests.Test_ParseBoardRun_ClosedAndResolved
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_ParseBoardRun_EmptyAndInvalid"
    NinjaTicketIndexTests.Test_ParseBoardRun_EmptyAndInvalid
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_GetTicket"
    NinjaTicketIndexTests.Test_Index_GetTicket
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_ExternalNumber_FiltersAndCaches"
    NinjaTicketIndexTests.Test_Index_ExternalNumber_FiltersAndCaches
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_ExternalNumber_CacheExpires"
    NinjaTicketIndexTests.Test_Index_ExternalNumber_CacheExpires
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_ExternalNumber_NoTicketIsCachedToo"
    NinjaTicketIndexTests.Test_Index_ExternalNumber_NoTicketIsCachedToo
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_ExternalNumber_FailureIsNotCached"
    NinjaTicketIndexTests.Test_Index_ExternalNumber_FailureIsNotCached
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_ExternalNumber_HttpErrorAndGarbage"
    NinjaTicketIndexTests.Test_Index_ExternalNumber_HttpErrorAndGarbage
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_OpenBySubject_PagesAndGroups"
    NinjaTicketIndexTests.Test_Index_OpenBySubject_PagesAndGroups
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_OpenBySubject_SinglePageStopsEarly"
    NinjaTicketIndexTests.Test_Index_OpenBySubject_SinglePageStopsEarly
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_OpenBySubject_Reloads"
    NinjaTicketIndexTests.Test_Index_OpenBySubject_Reloads
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_OpenBySubject_FailureKeepsOldCache"
    NinjaTicketIndexTests.Test_Index_OpenBySubject_FailureKeepsOldCache
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_OpenBySubject_FailureOnFirstLoadGivesNothingAndRetries"
    NinjaTicketIndexTests.Test_Index_OpenBySubject_FailureOnFirstLoadGivesNothingAndRetries
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_OpenBySubject_SecondPageFailureDiscardsPartialList"
    NinjaTicketIndexTests.Test_Index_OpenBySubject_SecondPageFailureDiscardsPartialList
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_Conversation"
    NinjaTicketIndexTests.Test_Index_Conversation
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_WithMatcher_ExternalNumberOnlyOneOpenTicket"
    NinjaTicketIndexTests.Test_Index_WithMatcher_ExternalNumberOnlyOneOpenTicket
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_WithMatcher_ApiDownMeansNoMatch"
    NinjaTicketIndexTests.Test_Index_WithMatcher_ApiDownMeansNoMatch
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketIndexTests.Test_Index_WithMatcher_SubjectRule"
    NinjaTicketIndexTests.Test_Index_WithMatcher_SubjectRule
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunNinjaTicketLogicTests()
    On Error Resume Next

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_TicketState_Open"
    NinjaTicketLogicTests.Test_TicketState_Open
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_TicketState_Closed"
    NinjaTicketLogicTests.Test_TicketState_Closed
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_TicketState_DeletedCountsAsClosed"
    NinjaTicketLogicTests.Test_TicketState_DeletedCountsAsClosed
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_TicketState_NotFound"
    NinjaTicketLogicTests.Test_TicketState_NotFound
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_TicketState_Other404IsUnavailable"
    NinjaTicketLogicTests.Test_TicketState_Other404IsUnavailable
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_TicketState_ErrorsAreUnavailable"
    NinjaTicketLogicTests.Test_TicketState_ErrorsAreUnavailable
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_TicketState_UnreadableAnswerIsUnavailable"
    NinjaTicketLogicTests.Test_TicketState_UnreadableAnswerIsUnavailable
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_TicketSubject"
    NinjaTicketLogicTests.Test_TicketSubject
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_UnixToDate"
    NinjaTicketLogicTests.Test_UnixToDate
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_ClosedDate_ClosedManually"
    NinjaTicketLogicTests.Test_ClosedDate_ClosedManually
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_ClosedDate_ClosedByAutomation"
    NinjaTicketLogicTests.Test_ClosedDate_ClosedByAutomation
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_ClosedDate_EnglishStatusName"
    NinjaTicketLogicTests.Test_ClosedDate_EnglishStatusName
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_ClosedDate_ReopenedAndClosedAgainUsesNewest"
    NinjaTicketLogicTests.Test_ClosedDate_ReopenedAndClosedAgainUsesNewest
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_ClosedDate_DeletedTicketUsesDeleteEntry"
    NinjaTicketLogicTests.Test_ClosedDate_DeletedTicketUsesDeleteEntry
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_ClosedDate_NoClosingEntryGivesZero"
    NinjaTicketLogicTests.Test_ClosedDate_NoClosingEntryGivesZero
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_ClosedDate_ToleratesStrangeEntries"
    NinjaTicketLogicTests.Test_ClosedDate_ToleratesStrangeEntries
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTicketLogicTests.Test_IsClosedStatusName"
    NinjaTicketLogicTests.Test_IsClosedStatusName
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunNinjaTokenLogicTests()
    On Error Resume Next

    TestAssert.BeginTest "NinjaTokenLogicTests.Test_UrlEncode_UnreservedUnchanged"
    NinjaTokenLogicTests.Test_UrlEncode_UnreservedUnchanged
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTokenLogicTests.Test_UrlEncode_ReservedCharacters"
    NinjaTokenLogicTests.Test_UrlEncode_ReservedCharacters
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTokenLogicTests.Test_UrlEncode_NonAsciiAsUtf8"
    NinjaTokenLogicTests.Test_UrlEncode_NonAsciiAsUtf8
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTokenLogicTests.Test_BuildTokenRequestBody"
    NinjaTokenLogicTests.Test_BuildTokenRequestBody
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTokenLogicTests.Test_ParseTokenResponse_Valid"
    NinjaTokenLogicTests.Test_ParseTokenResponse_Valid
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTokenLogicTests.Test_ParseTokenResponse_Invalid"
    NinjaTokenLogicTests.Test_ParseTokenResponse_Invalid
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTokenLogicTests.Test_ParseTokenResponse_ScopeOptional"
    NinjaTokenLogicTests.Test_ParseTokenResponse_ScopeOptional
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTokenLogicTests.Test_ParseErrorResponse"
    NinjaTokenLogicTests.Test_ParseErrorResponse
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTokenLogicTests.Test_ComputeExpiry"
    NinjaTokenLogicTests.Test_ComputeExpiry
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTokenLogicTests.Test_IsTokenValid"
    NinjaTokenLogicTests.Test_IsTokenValid
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTokenLogicTests.Test_Json_NumberIgnoresLocale"
    NinjaTokenLogicTests.Test_Json_NumberIgnoresLocale
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "NinjaTokenLogicTests.Test_Json_StringEscapes"
    NinjaTokenLogicTests.Test_Json_StringEscapes
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunReviewFormTests()
    On Error Resume Next

    TestAssert.BeginTest "ReviewFormTests.Test_Form_ListsTheProposals"
    ReviewFormTests.Test_Form_ListsTheProposals
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ReviewFormTests.Test_Form_AcceptAll"
    ReviewFormTests.Test_Form_AcceptAll
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ReviewFormTests.Test_Form_PartialDecisionsKeepTheOrder"
    ReviewFormTests.Test_Form_PartialDecisionsKeepTheOrder
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ReviewFormTests.Test_Form_NothingTickedChangesNothing"
    ReviewFormTests.Test_Form_NothingTickedChangesNothing
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ReviewFormTests.Test_Form_NoProposals"
    ReviewFormTests.Test_Form_NoProposals
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ReviewFormTests.Test_Form_ResultMatchesWhatTheReviewExpects"
    ReviewFormTests.Test_Form_ResultMatchesWhatTheReviewExpects
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "ReviewFormTests.Test_PromptText_ListLine"
    ReviewFormTests.Test_PromptText_ListLine
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunTicketDecisionTests()
    On Error Resume Next

    TestAssert.BeginTest "TicketDecisionTests.Test_Decision_ClosedIsArchivedByApi"
    TicketDecisionTests.Test_Decision_ClosedIsArchivedByApi
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketDecisionTests.Test_Decision_NotFoundIsArchivedByApi"
    TicketDecisionTests.Test_Decision_NotFoundIsArchivedByApi
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketDecisionTests.Test_Decision_OpenIsNeverArchived"
    TicketDecisionTests.Test_Decision_OpenIsNeverArchived
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketDecisionTests.Test_Decision_UnavailableUsesFallbackOnlyWhenEnabledAndMailSaysClosed"
    TicketDecisionTests.Test_Decision_UnavailableUsesFallbackOnlyWhenEnabledAndMailSaysClosed
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketDecisionTests.Test_Decision_NeedsMailCheckOnlyWhenApiUnavailableAndFallbackOn"
    TicketDecisionTests.Test_Decision_NeedsMailCheckOnlyWhenApiUnavailableAndFallbackOn
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketDecisionTests.Test_Decision_FallbackIsOnByDefault"
    TicketDecisionTests.Test_Decision_FallbackIsOnByDefault
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunTicketMatcherTests()
    On Error Resume Next

    TestAssert.BeginTest "TicketMatcherTests.Test_ExtractNinjaIds"
    TicketMatcherTests.Test_ExtractNinjaIds
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_ExtractNinjaIds_IgnoresVendorAndOthers"
    TicketMatcherTests.Test_ExtractNinjaIds_IgnoresVendorAndOthers
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_ExtractExternalNumbers"
    TicketMatcherTests.Test_ExtractExternalNumbers
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_NormalizeSubject"
    TicketMatcherTests.Test_NormalizeSubject
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_IsSpecificSubject"
    TicketMatcherTests.Test_IsSpecificSubject
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_IsExcluded"
    TicketMatcherTests.Test_IsExcluded
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_NinjaId_InSubject"
    TicketMatcherTests.Test_Match_NinjaId_InSubject
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_NinjaId_InBody_WhenSubjectHasNone"
    TicketMatcherTests.Test_Match_NinjaId_InBody_WhenSubjectHasNone
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_NinjaId_Unknown_FallsThrough"
    TicketMatcherTests.Test_Match_NinjaId_Unknown_FallsThrough
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_NinjaId_Several_IsAmbiguous"
    TicketMatcherTests.Test_Match_NinjaId_Several_IsAmbiguous
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_ExternalNumber_Found"
    TicketMatcherTests.Test_Match_ExternalNumber_Found
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_ExternalNumber_NotInAnyTicket_StaysInInbox"
    TicketMatcherTests.Test_Match_ExternalNumber_NotInAnyTicket_StaysInInbox
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_ExternalNumber_TwoTickets_OneOpen_Wins"
    TicketMatcherTests.Test_Match_ExternalNumber_TwoTickets_OneOpen_Wins
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_ExternalNumber_TwoOpenTickets_IsAmbiguous"
    TicketMatcherTests.Test_Match_ExternalNumber_TwoOpenTickets_IsAmbiguous
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_ExternalNumber_TwoClosedTickets_IsAmbiguous"
    TicketMatcherTests.Test_Match_ExternalNumber_TwoClosedTickets_IsAmbiguous
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_ExternalNumber_SingleClosedTicket_IsFound"
    TicketMatcherTests.Test_Match_ExternalNumber_SingleClosedTicket_IsFound
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_ExternalNumber_TwoNumbers_IsAmbiguous"
    TicketMatcherTests.Test_Match_ExternalNumber_TwoNumbers_IsAmbiguous
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_Conversation"
    TicketMatcherTests.Test_Match_Conversation
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_Subject_IgnoresPrefixes"
    TicketMatcherTests.Test_Match_Subject_IgnoresPrefixes
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_Subject_GenericSubjectNeverMatches"
    TicketMatcherTests.Test_Match_Subject_GenericSubjectNeverMatches
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_Subject_TwoOpenTickets_IsAmbiguous"
    TicketMatcherTests.Test_Match_Subject_TwoOpenTickets_IsAmbiguous
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_Subject_ClosedTicketDoesNotMatch"
    TicketMatcherTests.Test_Match_Subject_ClosedTicketDoesNotMatch
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_Subject_NoTicket"
    TicketMatcherTests.Test_Match_Subject_NoTicket
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_ExcludedMail_NeverQueriesIndex"
    TicketMatcherTests.Test_Match_ExcludedMail_NeverQueriesIndex
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_NinjaId_BeatsExternalNumber"
    TicketMatcherTests.Test_Match_NinjaId_BeatsExternalNumber
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_Match_NoIndex"
    TicketMatcherTests.Test_Match_NoIndex
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_HasAutoHeaderLine"
    TicketMatcherTests.Test_HasAutoHeaderLine
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatcherTests.Test_DescribeMatch"
    TicketMatcherTests.Test_DescribeMatch
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunTicketMatchServiceTests()
    On Error Resume Next

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_TicketSenderIsIgnored"
    TicketMatchServiceTests.Test_MatchMail_TicketSenderIsIgnored
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_LogOnlyNeverMoves"
    TicketMatchServiceTests.Test_MatchMail_LogOnlyNeverMoves
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_LogLineNamesRuleDomainAndSubject"
    TicketMatchServiceTests.Test_MatchMail_LogLineNamesRuleDomainAndSubject
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_MovesIntoExistingFolder"
    TicketMatchServiceTests.Test_MatchMail_MovesIntoExistingFolder
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_CreatesFolderNamedAfterTicketSubject"
    TicketMatchServiceTests.Test_MatchMail_CreatesFolderNamedAfterTicketSubject
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_RestoresArchivedFolderOfOpenTicket"
    TicketMatchServiceTests.Test_MatchMail_RestoresArchivedFolderOfOpenTicket
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_ClosedTicketStaysInInbox"
    TicketMatchServiceTests.Test_MatchMail_ClosedTicketStaysInInbox
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_UnreachableApiStaysInInbox"
    TicketMatchServiceTests.Test_MatchMail_UnreachableApiStaysInInbox
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_NoMatchOrAmbiguousNeverMoves"
    TicketMatchServiceTests.Test_MatchMail_NoMatchOrAmbiguousNeverMoves
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_UnknownVendorNumberStaysInInbox"
    TicketMatchServiceTests.Test_MatchMail_UnknownVendorNumberStaysInInbox
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_ExcludedMailIsNotMoved"
    TicketMatchServiceTests.Test_MatchMail_ExcludedMailIsNotMoved
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_ThreadFollowsAfterVendorNumberMatch"
    TicketMatchServiceTests.Test_MatchMail_ThreadFollowsAfterVendorNumberMatch
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketMatchServiceTests.Test_MatchMail_SubjectGuessIsNotRemembered"
    TicketMatchServiceTests.Test_MatchMail_SubjectGuessIsNotRemembered
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunTicketParsingTests()
    On Error Resume Next

    TestAssert.BeginTest "TicketParsingTests.Test_ParseSubject_Valid"
    TicketParsingTests.Test_ParseSubject_Valid
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketParsingTests.Test_ParseSubject_Invalid"
    TicketParsingTests.Test_ParseSubject_Invalid
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketParsingTests.Test_BuildFolderName"
    TicketParsingTests.Test_BuildFolderName
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketParsingTests.Test_BuildFolderName_ShortensTicketPrefix"
    TicketParsingTests.Test_BuildFolderName_ShortensTicketPrefix
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketParsingTests.Test_TicketIdFromFolderName"
    TicketParsingTests.Test_TicketIdFromFolderName
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketParsingTests.Test_FolderMatchesTicket"
    TicketParsingTests.Test_FolderMatchesTicket
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketParsingTests.Test_IsClosedMail"
    TicketParsingTests.Test_IsClosedMail
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketParsingTests.Test_ArchivePathNames"
    TicketParsingTests.Test_ArchivePathNames
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketParsingTests.Test_ArchivePathNames_ZeroMeansToday"
    TicketParsingTests.Test_ArchivePathNames_ZeroMeansToday
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketParsingTests.Test_AllowedSender_ContainsRealUmlaut"
    TicketParsingTests.Test_AllowedSender_ContainsRealUmlaut
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketParsingTests.Test_RegexTicketReplace_ContainsRealUmlaut"
    TicketParsingTests.Test_RegexTicketReplace_ContainsRealUmlaut
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunTicketServiceTests()
    On Error Resume Next

    TestAssert.BeginTest "TicketServiceTests.Test_Mail_NewTicketCreatesFolderAndMovesMail"
    TicketServiceTests.Test_Mail_NewTicketCreatesFolderAndMovesMail
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Mail_FolderIsRenamedToCurrentTicketSubject"
    TicketServiceTests.Test_Mail_FolderIsRenamedToCurrentTicketSubject
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Mail_ExistingFolderIsReused"
    TicketServiceTests.Test_Mail_ExistingFolderIsReused
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Mail_ShorterTicketNumberDoesNotMatchLongerFolder"
    TicketServiceTests.Test_Mail_ShorterTicketNumberDoesNotMatchLongerFolder
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Mail_ArchivedFolderIsRestored"
    TicketServiceTests.Test_Mail_ArchivedFolderIsRestored
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Mail_OtherSenderIsIgnored"
    TicketServiceTests.Test_Mail_OtherSenderIsIgnored
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Mail_NonTicketSubjectIsIgnored"
    TicketServiceTests.Test_Mail_NonTicketSubjectIsIgnored
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Mail_AbsurdTicketNumberIsIgnored"
    TicketServiceTests.Test_Mail_AbsurdTicketNumberIsIgnored
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Mail_FolderNameStaysIfSubjectIsAlreadyCurrent"
    TicketServiceTests.Test_Mail_FolderNameStaysIfSubjectIsAlreadyCurrent
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_ClosedTicketIsArchivedWithClosingDate"
    TicketServiceTests.Test_Archive_ClosedTicketIsArchivedWithClosingDate
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_OpenTicketStays"
    TicketServiceTests.Test_Archive_OpenTicketStays
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_OpenTicketStaysEvenIfOldMailSaysClosed"
    TicketServiceTests.Test_Archive_OpenTicketStaysEvenIfOldMailSaysClosed
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_DeletedTicketIsArchivedWithTodaysDate"
    TicketServiceTests.Test_Archive_DeletedTicketIsArchivedWithTodaysDate
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_ApiUnavailable_FallbackArchivesWithMailDate"
    TicketServiceTests.Test_Archive_ApiUnavailable_FallbackArchivesWithMailDate
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_ApiUnavailable_FallbackDisabledDoesNothing"
    TicketServiceTests.Test_Archive_ApiUnavailable_FallbackDisabledDoesNothing
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_ApiUnavailable_MailNotClosedStays"
    TicketServiceTests.Test_Archive_ApiUnavailable_MailNotClosedStays
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_ServerErrorCountsAsUnavailable"
    TicketServiceTests.Test_Archive_ServerErrorCountsAsUnavailable
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_NonTicketFolderIsSkipped"
    TicketServiceTests.Test_Archive_NonTicketFolderIsSkipped
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_FolderIsRenamedToCurrentSubject"
    TicketServiceTests.Test_Archive_FolderIsRenamedToCurrentSubject
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_OneRequestPerOpenTicket"
    TicketServiceTests.Test_Archive_OneRequestPerOpenTicket
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_StopsAskingTheApiWhenItIsDown"
    TicketServiceTests.Test_Archive_StopsAskingTheApiWhenItIsDown
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_ApiRecoveringResetsTheFailureCount"
    TicketServiceTests.Test_Archive_ApiRecoveringResetsTheFailureCount
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_ErrorInOneFolderDoesNotStopTheRun"
    TicketServiceTests.Test_Archive_ErrorInOneFolderDoesNotStopTheRun
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_ResultIsLogged"
    TicketServiceTests.Test_Archive_ResultIsLogged
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_ReportsProgress"
    TicketServiceTests.Test_Archive_ReportsProgress
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_ProgressFinishesEvenWhenAFolderFails"
    TicketServiceTests.Test_Archive_ProgressFinishesEvenWhenAFolderFails
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TicketServiceTests.Test_Archive_ProgressIsReportedForAnEmptyStore"
    TicketServiceTests.Test_Archive_ProgressIsReportedForAnEmptyStore
    TestAssert.EndTest Err.Number, Err.Description

End Sub

Private Sub RunTimeUtilTests()
    On Error Resume Next

    TestAssert.BeginTest "TimeUtilTests.Test_UtcToLocal_ZeroStaysZero"
    TimeUtilTests.Test_UtcToLocal_ZeroStaysZero
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TimeUtilTests.Test_UtcToLocal_OffsetIsPlausible"
    TimeUtilTests.Test_UtcToLocal_OffsetIsPlausible
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TimeUtilTests.Test_UtcToLocal_KeepsMinutesAndSeconds"
    TimeUtilTests.Test_UtcToLocal_KeepsMinutesAndSeconds
    TestAssert.EndTest Err.Number, Err.Description

    TestAssert.BeginTest "TimeUtilTests.Test_UtcToLocal_IsMonotonic"
    TimeUtilTests.Test_UtcToLocal_IsMonotonic
    TestAssert.EndTest Err.Number, Err.Description

End Sub

' ---------------------------------------------------------------------------
'  Output
' ---------------------------------------------------------------------------

Private Sub PublishResults(ByVal title As String, ByVal resultLine As String)
    Dim text As String

    text = title & " - " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
           TestAssert.Report() & _
           "Tests passed: " & TestAssert.PassedTests & _
           ", failed: " & TestAssert.FailedTests & _
           ", assertions: " & TestAssert.AssertionCount & vbCrLf & _
           resultLine & vbCrLf

    Debug.Print text
    WriteResultFile text
End Sub

Private Sub WriteResultFile(ByVal text As String)
    Dim fso As Object
    Dim stream As Object
    Dim path As String

    On Error GoTo Done
    path = Environ$("TEMP") & "\" & RESULT_FILE_NAME
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(path, True, False)
    stream.Write text
    stream.Close
Done:
End Sub
