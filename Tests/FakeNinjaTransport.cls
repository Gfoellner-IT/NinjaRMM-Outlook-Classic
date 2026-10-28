VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeNinjaTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FakeNinjaTransport v0.1.0
'
' Test double for INinjaTransport: answers GET requests from a table of canned
' responses and records every requested path.
Implements INinjaTransport
Option Explicit

Private pResponses As Object   ' path -> Array(status, body, errorText)
Private pRequested As String

Public RequestCount As Long
Public LastPath As String
Public PostCount As Long
Public LastPostBody As String

Private Sub Class_Initialize()
    Set pResponses = CreateObject("Scripting.Dictionary")
End Sub

' Canned answer with an HTTP status.
Public Sub AddResponse(ByVal path As String, ByVal httpStatus As Long, ByVal body As String)
    pResponses.Item(path) = Array(httpStatus, body, "")
End Sub

' Simulates "no response at all" (offline, timeout, no token).
Public Sub AddFailure(ByVal path As String, ByVal errorText As String)
    pResponses.Item(path) = Array(0, "", errorText)
End Sub

' Canned answer for a POST with exactly this JSON body.
Public Sub AddPostResponse(ByVal path As String, ByVal requestBody As String, ByVal httpStatus As Long, ByVal body As String)
    pResponses.Item("POST " & path & vbLf & requestBody) = Array(httpStatus, body, "")
End Sub

Public Sub AddPostFailure(ByVal path As String, ByVal requestBody As String, ByVal errorText As String)
    pResponses.Item("POST " & path & vbLf & requestBody) = Array(0, "", errorText)
End Sub

Public Function WasRequested(ByVal path As String) As Boolean
    WasRequested = (InStr(pRequested, "|" & path & "|") > 0)
End Function

Private Function INinjaTransport_HttpPostJson(ByVal path As String, ByVal jsonBody As String, ByRef httpStatus As Long, ByRef errorText As String) As String
    Dim entry As Variant
    Dim key As String

    PostCount = PostCount + 1
    LastPostBody = jsonBody
    key = "POST " & path & vbLf & jsonBody

    If pResponses.Exists(key) Then
        entry = pResponses.Item(key)
        httpStatus = CLng(entry(0))
        errorText = CStr(entry(2))
        INinjaTransport_HttpPostJson = CStr(entry(1))
    Else
        httpStatus = 0
        errorText = "no fake response configured for POST " & path & " " & jsonBody
        INinjaTransport_HttpPostJson = ""
    End If
End Function

Private Function INinjaTransport_HttpGet(ByVal path As String, ByRef httpStatus As Long, ByRef errorText As String) As String
    Dim entry As Variant

    RequestCount = RequestCount + 1
    LastPath = path
    pRequested = pRequested & "|" & path & "|"

    If pResponses.Exists(path) Then
        entry = pResponses.Item(path)
        httpStatus = CLng(entry(0))
        errorText = CStr(entry(2))
        INinjaTransport_HttpGet = CStr(entry(1))
    Else
        httpStatus = 0
        errorText = "no fake response configured for " & path
        INinjaTransport_HttpGet = ""
    End If
End Function
