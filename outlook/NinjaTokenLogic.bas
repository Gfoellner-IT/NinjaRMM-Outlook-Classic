Attribute VB_Name = "NinjaTokenLogic"
' NinjaTokenLogic v0.1.0
'
' Pure helper functions for the OAuth2 client-credentials flow:
' building the token request, reading the token response, expiry handling.
' No Outlook types, no network, no credential store - fully unit testable.
'
' The JSON helpers only understand flat top-level string and number values
' (all the token endpoint returns). They are replaced by a real JSON parser in M2.

Option Explicit

' Safety margin: a token is treated as expired this many seconds before it really is.
Public Const TOKEN_EXPIRY_SKEW_SECONDS As Long = 60

' ---------------------------------------------------------------------------
'  Request
' ---------------------------------------------------------------------------

' Builds the application/x-www-form-urlencoded body for the token request.
Public Function BuildTokenRequestBody(ByVal clientId As String, ByVal clientSecret As String, ByVal scope As String) As String
    BuildTokenRequestBody = "grant_type=client_credentials" & _
        "&client_id=" & UrlEncode(clientId) & _
        "&client_secret=" & UrlEncode(clientSecret) & _
        "&scope=" & UrlEncode(scope)
End Function

' Percent-encodes text as UTF-8 (RFC 3986 unreserved characters stay as they are).
Public Function UrlEncode(ByVal text As String) As String
    Dim i As Long
    Dim code As Long
    Dim result As String

    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1))
        If code < 0 Then code = code + 65536

        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                ' 0-9 A-Z a-z - . _ ~
                result = result & ChrW$(code)
            Case Is < 128
                result = result & "%" & Right$("0" & Hex$(code), 2)
            Case Is < 2048
                result = result & "%" & Hex$(192 + (code \ 64)) & _
                                  "%" & Hex$(128 + (code And 63))
            Case Else
                result = result & "%" & Hex$(224 + (code \ 4096)) & _
                                  "%" & Hex$(128 + ((code \ 64) And 63)) & _
                                  "%" & Hex$(128 + (code And 63))
        End Select
    Next i

    UrlEncode = result
End Function

' ---------------------------------------------------------------------------
'  Response
' ---------------------------------------------------------------------------

' Reads access_token, expires_in and (optional) scope from a token response.
' Returns False if the response is not a valid token response.
Public Function ParseTokenResponse(ByVal json As String, ByRef accessToken As String, _
        ByRef expiresInSeconds As Long, ByRef grantedScope As String) As Boolean
    Dim expires As Double

    accessToken = ""
    expiresInSeconds = 0
    grantedScope = ""

    If Not TryGetJsonString(json, "access_token", accessToken) Then Exit Function
    If Len(accessToken) = 0 Then Exit Function
    If Not TryGetJsonNumber(json, "expires_in", expires) Then Exit Function
    If expires <= 0 Then Exit Function

    expiresInSeconds = CLng(expires)
    TryGetJsonString json, "scope", grantedScope
    ParseTokenResponse = True
End Function

' Returns "error: description" (or just the code) from an OAuth error response,
' an empty string if the body contains no error information.
Public Function ParseErrorResponse(ByVal json As String) As String
    Dim code As String
    Dim description As String

    If Not TryGetJsonString(json, "error", code) Then Exit Function
    If TryGetJsonString(json, "error_description", description) Then
        ParseErrorResponse = code & ": " & description
    Else
        ParseErrorResponse = code
    End If
End Function

' ---------------------------------------------------------------------------
'  Expiry
' ---------------------------------------------------------------------------

' Point in time when a token received at receivedAt should be renewed.
Public Function ComputeExpiry(ByVal receivedAt As Date, ByVal expiresInSeconds As Long) As Date
    Dim usable As Long

    usable = expiresInSeconds - TOKEN_EXPIRY_SKEW_SECONDS
    If usable < 0 Then usable = 0
    ComputeExpiry = DateAdd("s", usable, receivedAt)
End Function

Public Function IsTokenValid(ByVal token As String, ByVal expiresAt As Date, ByVal currentTime As Date) As Boolean
    IsTokenValid = (Len(token) > 0) And (currentTime < expiresAt)
End Function

' ---------------------------------------------------------------------------
'  Minimal JSON access (flat objects only)
' ---------------------------------------------------------------------------

Public Function TryGetJsonString(ByVal json As String, ByVal key As String, ByRef value As String) As Boolean
    Dim re As Object
    Dim matches As Object

    value = ""
    If Not IsPlainKey(key) Then Exit Function

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & key & """\s*:\s*""((?:[^""\\]|\\.)*)"""
    Set matches = re.Execute(json)
    If matches.Count = 0 Then Exit Function

    value = UnescapeJsonString(CStr(matches(0).SubMatches(0)))
    TryGetJsonString = True
End Function

Public Function TryGetJsonNumber(ByVal json As String, ByVal key As String, ByRef value As Double) As Boolean
    Dim re As Object
    Dim matches As Object

    value = 0
    If Not IsPlainKey(key) Then Exit Function

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & key & """\s*:\s*(-?[0-9]+(?:\.[0-9]+)?)"
    Set matches = re.Execute(json)
    If matches.Count = 0 Then Exit Function

    ' Val ignores the system decimal separator, which is what JSON needs.
    value = Val(CStr(matches(0).SubMatches(0)))
    TryGetJsonNumber = True
End Function

Private Function IsPlainKey(ByVal key As String) As Boolean
    Dim i As Long
    Dim c As String

    If Len(key) = 0 Then Exit Function
    For i = 1 To Len(key)
        c = Mid$(key, i, 1)
        If Not (c Like "[A-Za-z0-9_]") Then Exit Function
    Next i
    IsPlainKey = True
End Function

Private Function UnescapeJsonString(ByVal text As String) As String
    text = Replace(text, "\""", """")
    text = Replace(text, "\/", "/")
    text = Replace(text, "\\", "\")
    UnescapeJsonString = text
End Function
