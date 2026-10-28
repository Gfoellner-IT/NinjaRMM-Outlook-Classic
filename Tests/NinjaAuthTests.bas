Attribute VB_Name = "NinjaAuthTests"
' NinjaAuthTests v0.1.0
'
' Tests of NinjaAuth with a fake token endpoint (FakeTokenFetcher) and a unique,
' temporary credential entry. No network access.

Option Explicit

Private Const TEST_CLIENT_ID As String = "test-client"
Private Const TEST_SECRET As String = "test-secret&1"

Private pTarget As String
Private pFake As FakeTokenFetcher

' ---------------------------------------------------------------------------
'  Fixture helpers
' ---------------------------------------------------------------------------

Private Sub Setup()
    NinjaAuth.ResetState
    pTarget = NinjaCredentialStoreTests.UniqueTarget()
    NinjaCredentialStore.WriteCredential pTarget, TEST_CLIENT_ID, TEST_SECRET

    Set pFake = New FakeTokenFetcher
    pFake.ResponseStatus = 200
    pFake.ResponseBody = Fixtures.TokenResponse()

    NinjaAuth.UseCredentialTarget pTarget
    NinjaAuth.UseFetcher pFake
End Sub

Private Sub Teardown()
    NinjaCredentialStore.DeleteCredential pTarget
    NinjaAuth.ResetState
    Set pFake = Nothing
End Sub

Private Function TryGetToken(ByRef errNumber As Long, ByRef errDescription As String) As String
    On Error Resume Next
    TryGetToken = NinjaAuth.GetAccessToken()
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
End Function

' ---------------------------------------------------------------------------
'  Tests
' ---------------------------------------------------------------------------

Public Sub Test_GetAccessToken_Success()
    Dim token As String
    Dim errNumber As Long
    Dim errDescription As String

    Setup
    token = TryGetToken(errNumber, errDescription)

    TestAssert.AssertEqual 0, errNumber, errDescription
    TestAssert.AssertEqual "FAKE-ACCESS-TOKEN-FOR-TESTS", token
    TestAssert.AssertEqual 1, pFake.CallCount
    TestAssert.AssertEqual NinjaAuth.NINJA_TOKEN_URL, pFake.LastUrl
    TestAssert.AssertEqual _
        "grant_type=client_credentials&client_id=test-client&client_secret=test-secret%261&scope=monitoring", _
        pFake.LastBody
    Teardown
End Sub

Public Sub Test_GetAccessToken_UsesCache()
    Dim errNumber As Long
    Dim errDescription As String

    Setup
    TryGetToken errNumber, errDescription
    TryGetToken errNumber, errDescription
    TryGetToken errNumber, errDescription

    TestAssert.AssertEqual 1, pFake.CallCount, "token is requested only once"
    Teardown
End Sub

Public Sub Test_GetAccessToken_RefetchAfterInvalidate()
    Dim errNumber As Long
    Dim errDescription As String

    Setup
    TryGetToken errNumber, errDescription
    NinjaAuth.InvalidateToken
    TryGetToken errNumber, errDescription

    TestAssert.AssertEqual 2, pFake.CallCount
    Teardown
End Sub

Public Sub Test_GetAccessToken_ShortLivedTokenIsNotCached()
    Dim errNumber As Long
    Dim errDescription As String

    Setup
    pFake.ResponseBody = "{""access_token"":""short"",""expires_in"":30,""token_type"":""Bearer""}"
    TryGetToken errNumber, errDescription
    TryGetToken errNumber, errDescription

    TestAssert.AssertEqual 2, pFake.CallCount, "a token that expires within the skew is renewed on every call"
    Teardown
End Sub

Public Sub Test_GetAccessToken_NoCredential()
    Dim errNumber As Long
    Dim errDescription As String

    Setup
    NinjaCredentialStore.DeleteCredential pTarget
    TryGetToken errNumber, errDescription

    TestAssert.AssertError NinjaAuth.ERR_NO_CREDENTIAL, errNumber, errDescription
    TestAssert.AssertEqual 0, pFake.CallCount, "no request without credential"
    Teardown
End Sub

Public Sub Test_GetAccessToken_AuthRejected()
    Dim errNumber As Long
    Dim errDescription As String

    Setup
    pFake.ResponseStatus = 401
    pFake.ResponseBody = Fixtures.TokenError()
    TryGetToken errNumber, errDescription

    TestAssert.AssertError NinjaAuth.ERR_TOKEN_REQUEST, errNumber, errDescription
    TestAssert.AssertTrue InStr(errDescription, "invalid_client") > 0, "error code is reported"
    TestAssert.AssertTrue InStr(errDescription, "401") > 0, "status is reported"
    TestAssert.AssertEqual 401, NinjaAuth.LastHttpStatus
    TestAssert.AssertFalse NinjaAuth.IsServiceUnavailable(errNumber), "a rejected login is not an outage"
    TestAssert.AssertEqual 0, InStr(errDescription, TEST_SECRET), "secret is never part of the message"
    Teardown
End Sub

Public Sub Test_GetAccessToken_NoResponseIsUnavailable()
    Dim errNumber As Long
    Dim errDescription As String

    Setup
    pFake.ResponseStatus = 0
    pFake.ResponseBody = ""
    TryGetToken errNumber, errDescription

    TestAssert.AssertError NinjaAuth.ERR_TOKEN_REQUEST, errNumber, errDescription
    TestAssert.AssertEqual 0, NinjaAuth.LastHttpStatus
    TestAssert.AssertTrue NinjaAuth.IsServiceUnavailable(errNumber), "no response = unavailable"
    Teardown
End Sub

Public Sub Test_GetAccessToken_ServerErrorIsUnavailable()
    Dim errNumber As Long
    Dim errDescription As String

    Setup
    pFake.ResponseStatus = 503
    pFake.ResponseBody = ""
    TryGetToken errNumber, errDescription

    TestAssert.AssertError NinjaAuth.ERR_TOKEN_REQUEST, errNumber, errDescription
    TestAssert.AssertTrue NinjaAuth.IsServiceUnavailable(errNumber), "5xx = unavailable"
    Teardown
End Sub

Public Sub Test_GetAccessToken_InvalidResponse()
    Dim errNumber As Long
    Dim errDescription As String

    Setup
    pFake.ResponseBody = "{}"
    TryGetToken errNumber, errDescription

    TestAssert.AssertError NinjaAuth.ERR_TOKEN_RESPONSE, errNumber, errDescription
    TestAssert.AssertFalse NinjaAuth.IsServiceUnavailable(errNumber)
    Teardown
End Sub
