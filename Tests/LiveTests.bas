Attribute VB_Name = "LiveTests"
' LiveTests v0.1.0
'
' Opt-in tests against the real NinjaOne token endpoint and the real Credential
' Manager entry (NinjaAuth.NINJA_CREDENTIAL_TARGET). Read-only: only requests a token.
' Not part of RunAllTests. Start with TestRunner.RunLiveTests or
' tools\Invoke-VbaTests.ps1 -Live. The token is never printed.

Option Explicit

Public Sub Test_Live_TokenFromRealEndpoint()
    Dim token As String
    Dim errNumber As Long
    Dim errDescription As String

    NinjaAuth.ResetState
    On Error Resume Next
    token = NinjaAuth.GetAccessToken()
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0

    TestAssert.AssertEqual 0, errNumber, errDescription
    TestAssert.AssertTrue Len(token) > 20, "token received"
    TestAssert.AssertEqual 200, NinjaAuth.LastHttpStatus

    ' Second call must come from the cache (same value, no new request needed).
    TestAssert.AssertEqual token, NinjaAuth.GetAccessToken(), "cached token is reused"
    NinjaAuth.ResetState
End Sub

' Live ticket ids and vendor references are installation settings, never repository fixtures.
Private Function LiveTicketId(ByVal key As String) As Long
    Dim text As String
    text = LocalSettings.Value(key, "")
    If Len(text) = 0 Or Len(text) > 9 Or Not IsNumeric(text) Then
        Err.Raise vbObjectError + 11300, "LiveTests", "Configure positive LiveOpenTicketId, LiveClosedTicketId and LiveVendorTicketId in local settings"
    End If
    LiveTicketId = CLng(text)
    If LiveTicketId <= 0 Then Err.Raise vbObjectError + 11300, "LiveTests", "Live test ticket ids must be positive"
End Function

Public Sub Test_Live_TicketStates()
    Dim openId As Long
    Dim closedId As Long
    openId = LiveTicketId("LiveOpenTicketId")
    closedId = LiveTicketId("LiveClosedTicketId")
    NinjaApi.ResetState
    NinjaAuth.ResetState
    TestAssert.AssertEqual NinjaTicketOpen, NinjaApi.GetTicketState(openId), NinjaApi.LastError
    TestAssert.AssertEqual NinjaTicketClosed, NinjaApi.GetTicketState(closedId), NinjaApi.LastError
    TestAssert.AssertTrue Len(NinjaApi.GetTicketSubject(closedId)) > 0
    TestAssert.AssertTrue NinjaApi.GetTicketClosedDate(closedId) > 0, "selected closed ticket has a closing event"
    TestAssert.AssertEqual 0#, CDbl(NinjaApi.GetTicketClosedDate(openId)), "selected open ticket has no prior closing event"
    NinjaApi.ResetState
    NinjaAuth.ResetState
End Sub

Public Sub Test_Live_TicketIndex()
    Dim impl As NinjaTicketIndex
    Dim index As ITicketIndex
    Dim found As Collection
    Dim number As String
    Dim expectedId As Long
    expectedId = LiveTicketId("LiveVendorTicketId")
    number = LocalSettings.Value("LiveVendorNumber", "")
    If Len(number) = 0 Then Err.Raise vbObjectError + 11301, "LiveTests", "Configure LiveVendorNumber in local settings"
    NinjaAuth.ResetState
    Set impl = New NinjaTicketIndex
    Set index = impl
    Set found = index.FindByExternalNumber(number)
    TestAssert.AssertEqual 1, found.Count, impl.LastError
    If found.Count = 1 Then TestAssert.AssertEqual expectedId, found(1).Id
    TestAssert.AssertTrue Not index.GetTicket(expectedId) Is Nothing
    NinjaAuth.ResetState
End Sub