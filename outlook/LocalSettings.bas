Attribute VB_Name = "LocalSettings"
' LocalSettings v0.1.0
' Non-secret installation settings. Credentials remain in Windows Credential Manager.
Option Explicit

Private Declare PtrSafe Function ReadProfile Lib "kernel32" Alias "GetPrivateProfileStringW" ( _
    ByVal section As LongPtr, ByVal key As LongPtr, ByVal defaultValue As LongPtr, _
    ByVal buffer As LongPtr, ByVal size As Long, ByVal fileName As LongPtr) As Long

Private pDefaults As Boolean
Private pFile As String

Public Function FilePath() As String
    If Len(pFile) > 0 Then
        FilePath = pFile
    Else
        FilePath = Environ$("LOCALAPPDATA") & "\NinjaRMM-Outlook-Classic\settings.ini"
    End If
End Function

Public Function Value(ByVal key As String, ByVal fallback As String) As String
    Dim buffer As String
    Dim section As String
    Dim path As String
    Dim length As Long

    Value = fallback
    If pDefaults Then Exit Function
    section = "Settings"
    path = FilePath()
    buffer = String$(32768, vbNullChar)
    length = ReadProfile(StrPtr(section), StrPtr(key), StrPtr(fallback), StrPtr(buffer), Len(buffer), StrPtr(path))
    If length >= Len(buffer) - 2 Then Err.Raise vbObjectError + 11200, "LocalSettings", "Setting exceeds maximum length"
    Value = Left$(buffer, length)
End Function

Public Function ApiBaseUrl() As String
    Dim url As String
    url = Value("ApiBaseUrl", "https://tenant.example.invalid/v2/")
    ' Do not send a bearer token to an unencrypted or credential-bearing URL.
    If LCase$(Left$(url, 8)) <> "https://" Or InStr(url, "@") > 0 Or InStr(url, "?") > 0 Or InStr(url, "#") > 0 Then
        Err.Raise vbObjectError + 11201, "LocalSettings", "ApiBaseUrl must be an HTTPS base URL without credentials, query or fragment"
    End If
    If Right$(url, 1) <> "/" Then url = url & "/"
    ApiBaseUrl = url
End Function

Public Function AllTicketsBoard() As Long
    Dim text As String
    Dim i As Long
    text = Value("AllTicketsBoard", "2")
    If Len(text) = 0 Or Len(text) > 9 Then GoTo Invalid
    For i = 1 To Len(text)
        If Mid$(text, i, 1) < "0" Or Mid$(text, i, 1) > "9" Then GoTo Invalid
    Next i
    AllTicketsBoard = CLng(text)
    If AllTicketsBoard > 0 Then Exit Function
Invalid:
    Err.Raise vbObjectError + 11202, "LocalSettings", "AllTicketsBoard must be a positive integer"
End Function

' Test seams: each test starts with defaults and restores normal reading when it ends.
Public Sub UseDefaults()
    pDefaults = True
    pFile = ""
End Sub

Public Sub UseFile(ByVal path As String)
    pDefaults = False
    pFile = path
End Sub

Public Sub ResetState()
    pDefaults = False
    pFile = ""
End Sub
