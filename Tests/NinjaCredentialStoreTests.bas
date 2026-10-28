Attribute VB_Name = "NinjaCredentialStoreTests"
' NinjaCredentialStoreTests v0.1.0
'
' Round trip through the real Windows Credential Manager API. Uses a unique,
' session-only test entry that is deleted again; no existing entry is touched.

Option Explicit

Public Function UniqueTarget() As String
    Randomize
    UniqueTarget = "NinjaRMM-Outlook-Classic/UnitTest/" & Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int(Rnd * 100000))
End Function

Public Sub Test_Credential_RoundTrip()
    Dim target As String
    Dim user As String
    Dim secret As String
    Dim expectedSecret As String

    target = UniqueTarget()
    expectedSecret = "p@ss w0rd!""&=" & ChrW$(252) & ChrW$(8364)

    TestAssert.AssertTrue NinjaCredentialStore.WriteCredential(target, "unit-test-client", expectedSecret), "write"
    TestAssert.AssertTrue NinjaCredentialStore.ReadCredential(target, user, secret), "read"
    TestAssert.AssertEqual "unit-test-client", user, "user name"
    TestAssert.AssertEqual expectedSecret, secret, "secret incl. special characters"

    TestAssert.AssertTrue NinjaCredentialStore.DeleteCredential(target), "delete"
    TestAssert.AssertFalse NinjaCredentialStore.ReadCredential(target, user, secret), "gone after delete"
    TestAssert.AssertEqual "", secret, "secret cleared when entry is missing"
End Sub

Public Sub Test_Credential_Overwrite()
    Dim target As String
    Dim user As String
    Dim secret As String

    target = UniqueTarget()
    NinjaCredentialStore.WriteCredential target, "first", "secret-one"
    NinjaCredentialStore.WriteCredential target, "second", "secret-two"
    TestAssert.AssertTrue NinjaCredentialStore.ReadCredential(target, user, secret)
    TestAssert.AssertEqual "second", user
    TestAssert.AssertEqual "secret-two", secret
    NinjaCredentialStore.DeleteCredential target
End Sub

Public Sub Test_Credential_Missing()
    Dim user As String
    Dim secret As String

    TestAssert.AssertFalse NinjaCredentialStore.ReadCredential(UniqueTarget(), user, secret)
    TestAssert.AssertEqual "", user
    TestAssert.AssertEqual "", secret
End Sub

Public Sub Test_Credential_EmptySecretRejected()
    TestAssert.AssertFalse NinjaCredentialStore.WriteCredential(UniqueTarget(), "user", "")
End Sub
