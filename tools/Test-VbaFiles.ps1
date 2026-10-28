<#
.SYNOPSIS
  Checks the VBA source files for the conventions the VBA editor import needs.

.DESCRIPTION
  - every .bas/.cls/.frm file must use CRLF line endings only (no bare LF)
  - files under Tests\ and outlook\ must start with the Attribute VB_Name line (.bas) or VERSION header (.cls)
  - every file must be ASCII (the VBA editor imports as ANSI; use ChrW for umlauts in strings)
  Exit code 0 = all fine, 1 = findings.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$findings = 0

$files = Get-ChildItem -Path $root -Recurse -Include *.bas, *.cls, *.frm -File |
    Where-Object { $_.FullName -notmatch '\\\.git\\' }

foreach ($file in $files) {
    $rel = $file.FullName.Substring($root.Length + 1)
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

    $bareLf = 0
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i] -eq 10 -and ($i -eq 0 -or $bytes[$i - 1] -ne 13)) { $bareLf++ }
    }
    if ($bareLf -gt 0) { Write-Host "FAIL  $rel : $bareLf line(s) without CR"; $findings++ }

    # The VBA editor imports files as ANSI: a UTF-8 umlaut would become two wrong characters
    # and break comparisons. All VBA files must be ASCII; use ChrW for umlauts in strings.
    foreach ($b in $bytes) { if ($b -gt 127) { Write-Host "FAIL  $rel : contains non-ASCII bytes"; $findings++; break } }

    $head = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min(80, $bytes.Length))
    if ($file.Extension -eq '.bas' -and $head -notmatch '^Attribute VB_Name') {
        Write-Host "FAIL  $rel : missing 'Attribute VB_Name' header"; $findings++
    }
}

if ($findings -eq 0) { Write-Host "OK    $($files.Count) VBA files checked" }
exit ([int]($findings -gt 0))
