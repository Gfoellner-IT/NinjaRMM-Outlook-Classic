Attribute VB_Name = "LocalSettingsTests"
' LocalSettingsTests v0.1.0
' Synthetic temporary INI files only. Never reads installation settings or credentials.
Option Explicit

Private Function WriteSettings(ByVal text As String) As String
    Dim fso As Object
    Dim stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    WriteSettings = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName())
    Set stream = fso.CreateTextFile(WriteSettings, True, True)
    stream.Write "[Settings]" & vbCrLf & text
    stream.Close
End Function

Private Sub RemoveSettings(ByVal path As String)
    LocalSettings.UseDefaults
    CreateObject("Scripting.FileSystemObject").DeleteFile path
End Sub

Public Sub Test_Settings_Defaults()
    TestAssert.AssertEqual "fallback", LocalSettings.Value("UnknownKey", "fallback")
    TestAssert.AssertEqual "https://tenant.example.invalid/v2/", LocalSettings.ApiBaseUrl()
    TestAssert.AssertEqual 2, LocalSettings.AllTicketsBoard()
End Sub

Public Sub Test_Settings_ReadsUnicodeAndNormalizesBase()
    Dim path As String
    path = WriteSettings("AllowedSender=Support B" & ChrW$(252) & "ro" & vbCrLf & _
                         "ApiBaseUrl=https://api.example.invalid/v2" & vbCrLf & "AllTicketsBoard=7")
    LocalSettings.UseFile path
    TestAssert.AssertEqual "Support B" & ChrW$(252) & "ro", TicketConfig.AllowedSender()
    TestAssert.AssertEqual "https://api.example.invalid/v2/", NinjaAuth.NINJA_API_BASE_URL
    TestAssert.AssertEqual 7, TicketConfig.NINJA_ALL_TICKETS_BOARD
    TestAssert.AssertEqual "fallback", LocalSettings.Value("Missing", "fallback")
    RemoveSettings path
End Sub

Public Sub Test_Settings_RejectsUnsafeApiBase()
    Dim url As Variant
    Dim path As String
    Dim value As String
    Dim errorNumber As Long
    For Each url In Array("http://api.example.invalid/v2/", "https://user:pass@api.example.invalid/", _
                          "https://api.example.invalid/?token=example", "https://api.example.invalid/#part")
        path = WriteSettings("ApiBaseUrl=" & CStr(url))
        LocalSettings.UseFile path
        On Error Resume Next
        value = LocalSettings.ApiBaseUrl()
        errorNumber = Err.Number
        Err.Clear
        On Error GoTo 0
        TestAssert.AssertTrue errorNumber <> 0, "unsafe base rejected"
        RemoveSettings path
    Next url
End Sub

Public Sub Test_Settings_RejectsInvalidBoard()
    Dim board As Variant
    Dim path As String
    Dim value As Long
    Dim errorNumber As Long
    For Each board In Array("0", "-1", "abc", "1.5", "1234567890")
        path = WriteSettings("AllTicketsBoard=" & CStr(board))
        LocalSettings.UseFile path
        On Error Resume Next
        value = LocalSettings.AllTicketsBoard()
        errorNumber = Err.Number
        Err.Clear
        On Error GoTo 0
        TestAssert.AssertTrue errorNumber <> 0, "invalid board rejected"
        RemoveSettings path
    Next board
End Sub

Public Sub Test_Settings_UnitDefaultsIgnoreFile()
    Dim path As String
    path = WriteSettings("AllowedSender=Local synthetic sender")
    LocalSettings.UseFile path
    TestAssert.AssertEqual "Local synthetic sender", TicketConfig.AllowedSender()
    LocalSettings.UseDefaults
    TestAssert.AssertEqual "IT Support | B" & ChrW$(252) & "ro", TicketConfig.AllowedSender()
    RemoveSettings path
End Sub
