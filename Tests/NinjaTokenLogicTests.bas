Attribute VB_Name = "NinjaTokenLogicTests"
' NinjaTokenLogicTests v0.1.0
'
' Unit tests for the pure helpers in NinjaTokenLogic.

Option Explicit

Public Sub Test_UrlEncode_UnreservedUnchanged()
    TestAssert.AssertEqual "AZaz09-._~", NinjaTokenLogic.UrlEncode("AZaz09-._~")
    TestAssert.AssertEqual "", NinjaTokenLogic.UrlEncode("")
End Sub

Public Sub Test_UrlEncode_ReservedCharacters()
    TestAssert.AssertEqual "a%20b%26c%3Dd", NinjaTokenLogic.UrlEncode("a b&c=d")
    TestAssert.AssertEqual "100%25", NinjaTokenLogic.UrlEncode("100%")
    TestAssert.AssertEqual "%22q%22", NinjaTokenLogic.UrlEncode("""q""")
    TestAssert.AssertEqual "a%2Bb%2Fc", NinjaTokenLogic.UrlEncode("a+b/c")
End Sub

Public Sub Test_UrlEncode_NonAsciiAsUtf8()
    TestAssert.AssertEqual "%C3%BC", NinjaTokenLogic.UrlEncode(ChrW$(252)), "u umlaut"
    TestAssert.AssertEqual "%E2%82%AC", NinjaTokenLogic.UrlEncode(ChrW$(8364)), "euro sign"
End Sub

Public Sub Test_BuildTokenRequestBody()
    TestAssert.AssertEqual _
        "grant_type=client_credentials&client_id=abc&client_secret=s%26cret&scope=monitoring", _
        NinjaTokenLogic.BuildTokenRequestBody("abc", "s&cret", "monitoring")
    TestAssert.AssertEqual _
        "grant_type=client_credentials&client_id=abc&client_secret=x&scope=monitoring%20management", _
        NinjaTokenLogic.BuildTokenRequestBody("abc", "x", "monitoring management")
End Sub

Public Sub Test_ParseTokenResponse_Valid()
    Dim token As String
    Dim expiresIn As Long
    Dim scope As String

    TestAssert.AssertTrue NinjaTokenLogic.ParseTokenResponse(Fixtures.TokenResponse(), token, expiresIn, scope)
    TestAssert.AssertEqual "FAKE-ACCESS-TOKEN-FOR-TESTS", token
    TestAssert.AssertEqual 3600, expiresIn
    TestAssert.AssertEqual "monitoring", scope
End Sub

Public Sub Test_ParseTokenResponse_Invalid()
    Dim token As String
    Dim expiresIn As Long
    Dim scope As String

    TestAssert.AssertFalse NinjaTokenLogic.ParseTokenResponse("", token, expiresIn, scope), "empty"
    TestAssert.AssertFalse NinjaTokenLogic.ParseTokenResponse("{}", token, expiresIn, scope), "empty object"
    TestAssert.AssertFalse NinjaTokenLogic.ParseTokenResponse("not json", token, expiresIn, scope), "garbage"
    TestAssert.AssertFalse NinjaTokenLogic.ParseTokenResponse(Fixtures.TokenError(), token, expiresIn, scope), "error response"
    TestAssert.AssertFalse NinjaTokenLogic.ParseTokenResponse("{""access_token"":"""",""expires_in"":3600}", token, expiresIn, scope), "empty token"
    TestAssert.AssertFalse NinjaTokenLogic.ParseTokenResponse("{""access_token"":""x""}", token, expiresIn, scope), "no expires_in"
    TestAssert.AssertFalse NinjaTokenLogic.ParseTokenResponse("{""access_token"":""x"",""expires_in"":0}", token, expiresIn, scope), "expires_in 0"
End Sub

Public Sub Test_ParseTokenResponse_ScopeOptional()
    Dim token As String
    Dim expiresIn As Long
    Dim scope As String

    TestAssert.AssertTrue NinjaTokenLogic.ParseTokenResponse("{""access_token"":""x"",""expires_in"":60}", token, expiresIn, scope)
    TestAssert.AssertEqual "", scope
End Sub

Public Sub Test_ParseErrorResponse()
    TestAssert.AssertEqual "invalid_client: Client authentication failed", NinjaTokenLogic.ParseErrorResponse(Fixtures.TokenError())
    TestAssert.AssertEqual "invalid_scope", NinjaTokenLogic.ParseErrorResponse("{""error"":""invalid_scope""}")
    TestAssert.AssertEqual "", NinjaTokenLogic.ParseErrorResponse(Fixtures.TokenResponse())
    TestAssert.AssertEqual "", NinjaTokenLogic.ParseErrorResponse("")
End Sub

Public Sub Test_ComputeExpiry()
    Dim received As Date
    received = DateSerial(2026, 10, 7) + TimeSerial(10, 0, 0)

    TestAssert.AssertDateNear received + TimeSerial(0, 59, 0), NinjaTokenLogic.ComputeExpiry(received, 3600), 0, "3600 s minus 60 s skew"
    TestAssert.AssertDateNear received, NinjaTokenLogic.ComputeExpiry(received, 30), 0, "shorter than skew expires immediately"
    TestAssert.AssertDateNear received, NinjaTokenLogic.ComputeExpiry(received, 60), 0, "exactly the skew"
End Sub

Public Sub Test_IsTokenValid()
    Dim t As Date
    t = DateSerial(2026, 10, 7) + TimeSerial(10, 0, 0)

    TestAssert.AssertTrue NinjaTokenLogic.IsTokenValid("x", t + 1, t), "not yet expired"
    TestAssert.AssertFalse NinjaTokenLogic.IsTokenValid("x", t, t), "expires now"
    TestAssert.AssertFalse NinjaTokenLogic.IsTokenValid("x", t - 1, t), "expired"
    TestAssert.AssertFalse NinjaTokenLogic.IsTokenValid("", t + 1, t), "no token"
End Sub

Public Sub Test_Json_NumberIgnoresLocale()
    Dim value As Double

    TestAssert.AssertTrue NinjaTokenLogic.TryGetJsonNumber("{""createTime"":1791295288.695002,""x"":1}", "createTime", value)
    TestAssert.AssertTrue Abs(value - 1791295288.695002) < 0.000001, "fraction kept with a dot as separator"
    TestAssert.AssertTrue NinjaTokenLogic.TryGetJsonNumber("{ ""n"" : -5 }", "n", value)
    TestAssert.AssertEqual -5, value
    TestAssert.AssertFalse NinjaTokenLogic.TryGetJsonNumber("{""n"":""5""}", "n", value), "a string is not a number"
End Sub

Public Sub Test_Json_StringEscapes()
    Dim value As String

    TestAssert.AssertTrue NinjaTokenLogic.TryGetJsonString("{""a"":""x\""y\\z\/w""}", "a", value)
    TestAssert.AssertEqual "x""y\z/w", value
    TestAssert.AssertFalse NinjaTokenLogic.TryGetJsonString("{""a"":1}", "a", value), "number is not a string"
    TestAssert.AssertFalse NinjaTokenLogic.TryGetJsonString("{""a"":""1""}", "b", value), "missing key"
    TestAssert.AssertFalse NinjaTokenLogic.TryGetJsonString("{""a"":""1""}", "a.*", value), "key must be plain"
End Sub
