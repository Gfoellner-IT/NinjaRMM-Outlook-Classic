<#
.SYNOPSIS
  Generates Tests\TestRunner.bas and Tests\IntegrationRunner.bas from the test modules.

.DESCRIPTION
  Outlook has no Application.Run, so every test must be called explicitly. This
  script scans Tests\SelfTests.bas and Tests\*Tests.bas for "Public Sub Test_*"
  procedures and writes one registration block per test. Tests\LiveTests.bas
  (opt-in, real network) is registered in RunLiveTests only.
  Tests\Outlook*Tests.bas use Outlook types (integration tests against the test PST); they are
  registered in the separate IntegrationRunner.bas so that TestRunner.bas still compiles without
  Outlook (headless run). Run this script after adding or renaming a test and commit the output.
  The output is ASCII with CRLF line endings.
#>
[CmdletBinding()]
param(
    [string]$TestDir,
    [string]$OutFile
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $TestDir) { $TestDir = Join-Path $root '..\Tests' }
if (-not $OutFile) { $OutFile = Join-Path $TestDir 'TestRunner.bas' }

function Get-TestNames([string]$path) {
    Select-String -Path $path -Pattern '^\s*Public Sub (Test_\w+)\s*\(\s*\)' |
        ForEach-Object { $_.Matches[0].Groups[1].Value }
}

# Suites: SelfTests first, then the other *Tests.bas modules alphabetically.
$suiteFiles = @()
$self = Join-Path $TestDir 'SelfTests.bas'
if (Test-Path $self) { $suiteFiles += Get-Item $self }
$suiteFiles += Get-ChildItem -Path $TestDir -Filter '*Tests.bas' |
    Where-Object { $_.Name -notin 'SelfTests.bas', 'LiveTests.bas' -and $_.Name -notlike 'Outlook*Tests.bas' } | Sort-Object Name

$live = Join-Path $TestDir 'LiveTests.bas'

$sb = New-Object System.Text.StringBuilder
function Add([string]$line) { [void]$sb.Append($line + "`r`n") }

function Add-Block([string]$module, [string]$test) {
    Add "    TestAssert.BeginTest ""$module.$test"""
    Add "    $module.$test"
    Add '    TestAssert.EndTest Err.Number, Err.Description'
    Add ''
}

Add 'Attribute VB_Name = "TestRunner"'
Add "' GENERATED FILE - do not edit. Regenerate with: tools\New-TestRunnerModule.ps1"
Add "'"
Add "' Entry points for all unit tests. Start RunAllTests from the VBA editor (F5)"
Add "' or headless with tools\Invoke-VbaTests.ps1."
Add "' Results go to the Immediate window and to %TEMP%\ninja_test_results.txt."
Add "'"
Add "' Outlook has no Application.Run, so every test is called explicitly. The calls are"
Add "' generated from the ""Public Sub Test_*"" procedures of the test modules."
Add "'"
Add "' Unit tests must not touch the mailbox or the network. Tests that use the"
Add "' Windows Credential Manager only use unique, temporary entries."
Add ''
Add 'Option Explicit'
Add ''
Add 'Public Const RESULT_FILE_NAME As String = "ninja_test_results.txt"'
Add ''
Add 'Public Sub RunAllTests()'
Add '    TestAssert.Reset'
Add ''
foreach ($f in $suiteFiles) { Add ("    Run" + [IO.Path]::GetFileNameWithoutExtension($f.Name)) }
Add ''
Add '    PublishResults "Unit tests", IIf(TestAssert.AllPassed, "RESULT: OK", "RESULT: FAILED")'
Add 'End Sub'
Add ''
Add "' Opt-in: real requests against NinjaOne (needs network and the Credential Manager entry)."
Add 'Public Sub RunLiveTests()'
Add '    TestAssert.Reset'
Add '    On Error Resume Next'
Add ''
if (Test-Path $live) { foreach ($t in Get-TestNames $live) { Add-Block 'LiveTests' $t } }
Add '    On Error GoTo 0'
Add '    PublishResults "Live tests", IIf(TestAssert.AllPassed, "RESULT: OK", "RESULT: FAILED")'
Add 'End Sub'
Add ''
Add "' Proves that the runner really reports failures: two deliberately failing"
Add "' tests and one passing test must give exactly ""2 failed, 1 passed""."
Add 'Public Sub RunFailureProbe()'
Add '    Dim ok As Boolean'
Add ''
Add '    TestAssert.Reset'
Add '    On Error Resume Next'
Add ''
Add '    TestAssert.BeginTest "Probe.AssertionFails"'
Add '    SelfTests.Probe_AssertionFails'
Add '    TestAssert.EndTest Err.Number, Err.Description'
Add ''
Add '    TestAssert.BeginTest "Probe.RaisesError"'
Add '    SelfTests.Probe_RaisesError'
Add '    TestAssert.EndTest Err.Number, Err.Description'
Add ''
Add '    TestAssert.BeginTest "Probe.Passes"'
Add '    SelfTests.Probe_Passes'
Add '    TestAssert.EndTest Err.Number, Err.Description'
Add ''
Add '    On Error GoTo 0'
Add '    ok = (TestAssert.FailedTests = 2 And TestAssert.PassedTests = 1)'
Add '    PublishResults "Failure probe (expected: 2 failed, 1 passed)", IIf(ok, "RESULT: PROBE OK", "RESULT: PROBE FAILED")'
Add 'End Sub'
Add ''
Add "' ---------------------------------------------------------------------------"
Add "'  Suites"
Add "' ---------------------------------------------------------------------------"

$count = 0
foreach ($f in $suiteFiles) {
    $module = [IO.Path]::GetFileNameWithoutExtension($f.Name)
    $tests = @(Get-TestNames $f.FullName)
    Add ''
    Add "Private Sub Run$module()"
    Add '    On Error Resume Next'
    Add ''
    foreach ($t in $tests) { Add-Block $module $t; $count++ }
    Add 'End Sub'
}

Add ''
Add "' ---------------------------------------------------------------------------"
Add "'  Output"
Add "' ---------------------------------------------------------------------------"
Add ''
Add 'Private Sub PublishResults(ByVal title As String, ByVal resultLine As String)'
Add '    Dim text As String'
Add ''
Add '    text = title & " - " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _'
Add '           TestAssert.Report() & _'
Add '           "Tests passed: " & TestAssert.PassedTests & _'
Add '           ", failed: " & TestAssert.FailedTests & _'
Add '           ", assertions: " & TestAssert.AssertionCount & vbCrLf & _'
Add '           resultLine & vbCrLf'
Add ''
Add '    Debug.Print text'
Add '    WriteResultFile text'
Add 'End Sub'
Add ''
Add 'Private Sub WriteResultFile(ByVal text As String)'
Add '    Dim fso As Object'
Add '    Dim stream As Object'
Add '    Dim path As String'
Add ''
Add '    On Error GoTo Done'
Add '    path = Environ$("TEMP") & "\" & RESULT_FILE_NAME'
Add '    Set fso = CreateObject("Scripting.FileSystemObject")'
Add '    Set stream = fso.CreateTextFile(path, True, False)'
Add '    stream.Write text'
Add '    stream.Close'
Add 'Done:'
Add 'End Sub'

[System.IO.File]::WriteAllText($OutFile, $sb.ToString(), (New-Object System.Text.ASCIIEncoding))
Write-Host "Wrote $OutFile ($count unit tests in $($suiteFiles.Count) suites)"

# ---------------------------------------------------------------------------
#  IntegrationRunner.bas: Outlook-typed tests (test PST only), started manually with F5
# ---------------------------------------------------------------------------
$integrationFiles = @(Get-ChildItem -Path $TestDir -Filter 'Outlook*Tests.bas' | Sort-Object Name)
$ib = New-Object System.Text.StringBuilder
function IAdd([string]$line) { [void]$ib.Append($line + "`r`n") }

IAdd 'Attribute VB_Name = "IntegrationRunner"'
IAdd "' GENERATED FILE - do not edit. Regenerate with: tools\New-TestRunnerModule.ps1"
IAdd "'"
IAdd "' Integration tests of the Outlook adapters. They use Outlook types and run ONLY in the"
IAdd "' test PST (see OutlookTestGuard); they cannot run headless. Start RunIntegrationTests"
IAdd "' from the VBA editor (F5). Result: Immediate window and %TEMP%\ninja_integration_results.txt."
IAdd ''
IAdd 'Option Explicit'
IAdd ''
IAdd 'Public Const INTEGRATION_RESULT_FILE_NAME As String = "ninja_integration_results.txt"'
IAdd ''
IAdd 'Public Sub RunIntegrationTests()'
IAdd '    TestAssert.Reset'
IAdd '    TicketLog.DisableFileLogging'
IAdd '    On Error Resume Next'
IAdd ''
$integrationCount = 0
foreach ($f in $integrationFiles) {
    $module = [IO.Path]::GetFileNameWithoutExtension($f.Name)
    foreach ($t in @(Get-TestNames $f.FullName)) {
        IAdd "    TestAssert.BeginTest ""$module.$t"""
        IAdd "    $module.$t"
        IAdd '    TestAssert.EndTest Err.Number, Err.Description'
        IAdd ''
        $integrationCount++
    }
}
IAdd '    On Error GoTo 0'
IAdd '    Publish "Integration tests (test PST NinjaTest)", IIf(TestAssert.AllPassed, "RESULT: OK", "RESULT: FAILED")'
IAdd 'End Sub'
IAdd ''
IAdd 'Private Sub Publish(ByVal title As String, ByVal resultLine As String)'
IAdd '    Dim text As String'
IAdd '    Dim fso As Object'
IAdd '    Dim stream As Object'
IAdd ''
IAdd '    text = title & " - " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _'
IAdd '           TestAssert.Report() & _'
IAdd '           "Tests passed: " & TestAssert.PassedTests & _'
IAdd '           ", failed: " & TestAssert.FailedTests & _'
IAdd '           ", assertions: " & TestAssert.AssertionCount & vbCrLf & _'
IAdd '           resultLine & vbCrLf'
IAdd ''
IAdd '    Debug.Print text'
IAdd '    On Error GoTo Done'
IAdd '    Set fso = CreateObject("Scripting.FileSystemObject")'
IAdd '    Set stream = fso.CreateTextFile(Environ$("TEMP") & "\" & INTEGRATION_RESULT_FILE_NAME, True, False)'
IAdd '    stream.Write text'
IAdd '    stream.Close'
IAdd 'Done:'
IAdd 'End Sub'

$integrationOut = Join-Path $TestDir 'IntegrationRunner.bas'
if ($integrationFiles.Count -gt 0) {
    [System.IO.File]::WriteAllText($integrationOut, $ib.ToString(), (New-Object System.Text.ASCIIEncoding))
    Write-Host "Wrote $integrationOut ($integrationCount integration tests)"
} elseif (Test-Path $integrationOut) {
    Remove-Item $integrationOut
}
