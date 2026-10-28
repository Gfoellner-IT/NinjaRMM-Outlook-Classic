Attribute VB_Name = "TestAssert"
' TestAssert v0.1.0
'
' Minimal assertion helpers for the unit tests in this project.
' No references are required (late binding only).
'
' Usage (see TestRunner.bas):
'   TestAssert.BeginTest "Suite.TestName"
'   TestAssert.AssertEqual "expected", actual, "optional message"
'   TestAssert.EndTest Err.Number, Err.Description
'
' The runner calls every test with "On Error Resume Next" and hands the
' error state to EndTest, so one failing test never stops the whole run.

Option Explicit

Private pPassedTests As Long
Private pFailedTests As Long
Private pAssertions As Long
Private pCurrentTest As String
Private pCurrentFailed As Boolean
Private pLines As Collection

' ---------------------------------------------------------------------------
'  Lifecycle
' ---------------------------------------------------------------------------

Public Sub Reset()
    ' Tests never write the text log, the audit log or any other file of the real installation, and never
    ' read the real rules file: every runner starts with a Reset.
    TicketLog.DisableFileLogging
    pPassedTests = 0
    pFailedTests = 0
    pAssertions = 0
    pCurrentTest = ""
    pCurrentFailed = False
    Set pLines = New Collection
End Sub

Public Sub BeginTest(ByVal testName As String)
    If pLines Is Nothing Then Reset
    LocalSettings.ResetState
    If Left$(testName, 10) <> "LiveTests." Then LocalSettings.UseDefaults
    pCurrentTest = testName
    pCurrentFailed = False
End Sub

' Closes the current test. errNumber/errDescription are the values of the
' VBA Err object right after the test procedure returned.
Public Sub EndTest(ByVal errNumber As Long, ByVal errDescription As String)
    If errNumber <> 0 Then
        Fail "Unexpected error " & errNumber & ": " & errDescription
    End If
    Err.Clear
    LocalSettings.ResetState

    If pCurrentFailed Then
        pFailedTests = pFailedTests + 1
        AddLine "FAIL  " & pCurrentTest
    Else
        pPassedTests = pPassedTests + 1
        AddLine "PASS  " & pCurrentTest
    End If
    pCurrentTest = ""
End Sub

' ---------------------------------------------------------------------------
'  Results
' ---------------------------------------------------------------------------

Public Property Get PassedTests() As Long
    PassedTests = pPassedTests
End Property

Public Property Get FailedTests() As Long
    FailedTests = pFailedTests
End Property

Public Property Get AssertionCount() As Long
    AssertionCount = pAssertions
End Property

Public Function AllPassed() As Boolean
    AllPassed = (pFailedTests = 0 And pPassedTests > 0)
End Function

' Returns the complete report (one line per test and per failure).
Public Function Report() As String
    Dim result As String
    Dim item As Variant

    If pLines Is Nothing Then Reset
    For Each item In pLines
        result = result & CStr(item) & vbCrLf
    Next item
    Report = result
End Function

' ---------------------------------------------------------------------------
'  Assertions
' ---------------------------------------------------------------------------

Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal message As String = "")
    pAssertions = pAssertions + 1
    If Not condition Then Fail "Expected True. " & message
End Sub

Public Sub AssertFalse(ByVal condition As Boolean, Optional ByVal message As String = "")
    pAssertions = pAssertions + 1
    If condition Then Fail "Expected False. " & message
End Sub

Public Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
    pAssertions = pAssertions + 1
    If Not AreEqual(expected, actual) Then
        Fail "Expected <" & Describe(expected) & "> but was <" & Describe(actual) & ">. " & message
    End If
End Sub

Public Sub AssertNotEqual(ByVal notExpected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
    pAssertions = pAssertions + 1
    If AreEqual(notExpected, actual) Then
        Fail "Expected a value different from <" & Describe(notExpected) & ">. " & message
    End If
End Sub

' Asserts that two Date values are at most toleranceSeconds apart.
Public Sub AssertDateNear(ByVal expected As Date, ByVal actual As Date, Optional ByVal toleranceSeconds As Long = 1, Optional ByVal message As String = "")
    pAssertions = pAssertions + 1
    If Abs(DateDiff("s", expected, actual)) > toleranceSeconds Then
        Fail "Expected date near <" & Format$(expected, "yyyy-mm-dd hh:nn:ss") & "> but was <" & Format$(actual, "yyyy-mm-dd hh:nn:ss") & ">. " & message
    End If
End Sub

' Asserts that the error number captured by the caller equals expectedNumber.
' Typical use: On Error Resume Next / call / TestAssert.AssertError 11041, Err.Number
Public Sub AssertError(ByVal expectedNumber As Long, ByVal actualNumber As Long, Optional ByVal message As String = "")
    pAssertions = pAssertions + 1
    If expectedNumber <> actualNumber Then
        Fail "Expected error " & expectedNumber & " but was " & actualNumber & ". " & message
    End If
End Sub

' Records an explicit failure (for use in tests when no assertion fits).
Public Sub Fail(ByVal message As String)
    pCurrentFailed = True
    AddLine "      " & pCurrentTest & ": " & message
End Sub

' ---------------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------------

' Value comparison used by AssertEqual (public so it can be unit tested).
' Strings compare case-sensitive. Numbers, Booleans and Dates compare by value.
' Null equals Null, Empty equals Empty, objects compare by identity.
Public Function AreEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    On Error GoTo Mismatch

    If IsObject(a) Or IsObject(b) Then
        If IsObject(a) And IsObject(b) Then AreEqual = (a Is b)
        Exit Function
    End If

    If IsNull(a) Or IsNull(b) Then
        AreEqual = (IsNull(a) And IsNull(b))
        Exit Function
    End If

    If IsEmpty(a) Or IsEmpty(b) Then
        AreEqual = (IsEmpty(a) And IsEmpty(b))
        Exit Function
    End If

    If VarType(a) = vbString Or VarType(b) = vbString Then
        If VarType(a) = vbString And VarType(b) = vbString Then
            AreEqual = (StrComp(CStr(a), CStr(b), vbBinaryCompare) = 0)
        End If
        Exit Function
    End If

    AreEqual = (CDbl(a) = CDbl(b))
    Exit Function

Mismatch:
    AreEqual = False
End Function

Public Function Describe(ByVal value As Variant) As String
    On Error GoTo Fallback

    If IsObject(value) Then
        Describe = "[object " & TypeName(value) & "]"
    ElseIf IsNull(value) Then
        Describe = "Null"
    ElseIf IsEmpty(value) Then
        Describe = "Empty"
    Else
        Describe = TypeName(value) & ":" & CStr(value)
    End If
    Exit Function

Fallback:
    Describe = "[" & TypeName(value) & "]"
End Function

Private Sub AddLine(ByVal text As String)
    If pLines Is Nothing Then Reset
    pLines.Add text
End Sub
