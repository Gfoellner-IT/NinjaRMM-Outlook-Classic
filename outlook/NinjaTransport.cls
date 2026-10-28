VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NinjaTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' NinjaTransport v0.1.0
'
' INinjaTransport on top of MSXML2.ServerXMLHTTP.6.0 (late bound, no reference needed).
' Adds the bearer token from NinjaAuth, uses the system proxy settings, explicit
' timeouts and retries once with a fresh token after a 401.
' The response is decoded as UTF-8 (independent of the charset header).
Implements INinjaTransport
Option Explicit

' Timeouts in milliseconds: resolve, connect, send, receive.
Private Const TIMEOUT_RESOLVE As Long = 5000
Private Const TIMEOUT_CONNECT As Long = 5000
Private Const TIMEOUT_SEND As Long = 10000
Private Const TIMEOUT_RECEIVE As Long = 20000

Private Function INinjaTransport_HttpGet(ByVal path As String, ByRef httpStatus As Long, ByRef errorText As String) As String
    INinjaTransport_HttpGet = Send("GET", path, "", httpStatus, errorText)
End Function

Private Function INinjaTransport_HttpPostJson(ByVal path As String, ByVal jsonBody As String, ByRef httpStatus As Long, ByRef errorText As String) As String
    INinjaTransport_HttpPostJson = Send("POST", path, jsonBody, httpStatus, errorText)
End Function

Private Function Send(ByVal method As String, ByVal path As String, ByVal jsonBody As String, _
                      ByRef httpStatus As Long, ByRef errorText As String) As String
    Dim http As Object
    Dim token As String
    Dim attempt As Long

    httpStatus = 0
    errorText = ""

    On Error GoTo Failed

    For attempt = 1 To 2
        token = NinjaAuth.GetAccessToken()

        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.setTimeouts TIMEOUT_RESOLVE, TIMEOUT_CONNECT, TIMEOUT_SEND, TIMEOUT_RECEIVE
        http.Open method, NinjaAuth.NINJA_API_BASE_URL & path, False
        http.setRequestHeader "Authorization", "Bearer " & token
        http.setRequestHeader "Accept", "application/json"
        If method = "POST" Then
            http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
            http.send StringToUtf8(jsonBody)
        Else
            http.send
        End If

        httpStatus = CLng(http.Status)
        If httpStatus <> 401 Or attempt = 2 Then Exit For

        ' The token was rejected: drop it and try once more with a new one.
        NinjaAuth.InvalidateToken
    Next attempt

    Send = Utf8ToString(http.responseBody)
    Exit Function

Failed:
    httpStatus = 0
    errorText = Err.Description
    Send = ""
End Function

Private Function StringToUtf8(ByVal text As String) As Variant
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.Position = 0
    stream.Type = 1                ' adTypeBinary
    stream.Position = 3            ' skip the UTF-8 byte order mark
    StringToUtf8 = stream.Read
    stream.Close
End Function

Private Function Utf8ToString(ByVal bytes As Variant) As String
    Dim stream As Object

    If IsEmpty(bytes) Then Exit Function
    If Not IsArray(bytes) Then Exit Function

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1                ' adTypeBinary
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2                ' adTypeText
    stream.Charset = "utf-8"
    Utf8ToString = stream.ReadText
    stream.Close
End Function
