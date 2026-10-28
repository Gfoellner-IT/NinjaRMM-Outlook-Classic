<#
.SYNOPSIS
  Runs the VBA unit tests headless in a throw-away Excel instance (no Outlook, no mailbox).

.DESCRIPTION
  Creates a hidden Excel instance and an empty workbook, imports the modules
  listed in tools\unit-test-modules.txt and runs TestRunner.RunAllTests via
  Application.Run. The result is read from %TEMP%\ninja_test_results.txt and
  printed. The workbook is never saved and Excel is closed afterwards.

  Only modules without Outlook types can run here (pure logic such as parsing,
  JSON, decision functions, and the test framework itself). Outlook-dependent
  code is covered by the integration tests (milestone M4).

  Prerequisite: Excel must allow programmatic access to the VBA project
  (Trust Center -> Macro Settings -> "Trust access to the VBA project object
  model", registry HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1).
  This script does not change any security setting; it reports if the access is missing.

  Exit code: 0 = all tests passed, 1 = failures, 2 = environment problem.
#>
[CmdletBinding()]
param(
    [string]$ModuleList,
    # Runs TestRunner.RunFailureProbe instead: proves that failing tests are reported as failures.
    [switch]$Probe,
    # Runs TestRunner.RunLiveTests (real token request, needs network and the credential entry).
    [switch]$Live,
    # A VBA compile error opens a modal dialog in Excel that nobody can click. After this time Excel is killed.
    [int]$TimeoutSeconds = 180
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $ModuleList) { $ModuleList = Join-Path $root 'tools\unit-test-modules.txt' }
$resultFile = Join-Path $env:TEMP 'ninja_test_results.txt'

$modules = Get-Content $ModuleList |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -and -not $_.StartsWith('#') } |
    ForEach-Object { Join-Path $root $_ }
foreach ($m in $modules) { if (-not (Test-Path $m)) { Write-Host "ENV   module not found: $m"; exit 2 } }

if (Test-Path $resultFile) { Remove-Item $resultFile -Force }

$xl = $null
$wb = $null
$watchdog = $null

if (-not ('NinjaTestNative.Win' -as [type])) {
    Add-Type -Name Win -Namespace NinjaTestNative -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);'
}
try {
    try {
        $xl = New-Object -ComObject Excel.Application
    } catch {
        Write-Host "ENV   Excel could not be started: $($_.Exception.Message)"
        exit 2
    }
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    # Watchdog: separate process that kills this Excel instance if the run hangs.
    [uint32]$excelPid = 0
    [void][NinjaTestNative.Win]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$excelPid)
    if ($excelPid -gt 0) {
        $watchdog = Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -PassThru -ArgumentList @(
            '-NoProfile', '-Command', "Start-Sleep -Seconds $TimeoutSeconds; Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue")
    }
    $wb = $xl.Workbooks.Add()

    $project = $null
    try {
        $project = $wb.VBProject
        if ($null -eq $project) { throw 'VBProject is not accessible' }
        $null = $project.VBComponents.Count
    } catch {
        Write-Host 'ENV   No programmatic access to the VBA project in Excel.'
        Write-Host '      Enable "Trust access to the VBA project object model" for Excel'
        Write-Host '      (HKCU\Software\Microsoft\Office\16.0\Excel\Security, DWORD AccessVBOM = 1).'
        exit 2
    }

    foreach ($m in $modules) { $null = $project.VBComponents.Import($m) }
    Write-Host "Imported $($modules.Count) module(s) into a temporary workbook."

    $entry = if ($Probe) { 'RunFailureProbe' } elseif ($Live) { 'RunLiveTests' } else { 'RunAllTests' }
    $null = $xl.Run("'$($wb.Name)'!TestRunner.$entry")
}
catch {
    Write-Host "ENV   $($_.Exception.Message)"
    exit 2
}
finally {
    if ($watchdog) { try { Stop-Process -Id $watchdog.Id -Force -ErrorAction SilentlyContinue } catch {} }
    if ($wb) { try { $wb.Close($false) } catch {} ; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($xl) { try { $xl.Quit() } catch {} ; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if (-not (Test-Path $resultFile)) { Write-Host 'ENV   Test run produced no result file.'; exit 2 }
$text = Get-Content $resultFile -Raw
Write-Host $text
if ($text -match 'RESULT: (OK|PROBE OK)') { exit 0 } else { exit 1 }
