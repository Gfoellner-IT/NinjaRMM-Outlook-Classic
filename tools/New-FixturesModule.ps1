<#
.SYNOPSIS
  Generates Tests\Fixtures.bas from Tests\fixtures\*.json.

.DESCRIPTION
  The JSON files are the single source of truth. Every file becomes a VBA
  function that returns its content, e.g. ticket_open.json -> Fixtures.TicketOpen().
  Run this script after changing a fixture and commit the generated module.
  Input files must be ASCII-only. The output uses CRLF line endings.
#>
[CmdletBinding()]
param(
    [string]$FixtureDir,
    [string]$OutFile
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $FixtureDir) { $FixtureDir = Join-Path $root '..\Tests\fixtures' }
if (-not $OutFile) { $OutFile = Join-Path $root '..\Tests\Fixtures.bas' }
$chunkSize = 160

function ConvertTo-PascalCase([string]$name) {
    ($name -split '[_\-\s]+' | Where-Object { $_ } | ForEach-Object { $_.Substring(0, 1).ToUpper() + $_.Substring(1) }) -join ''
}

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('Attribute VB_Name = "Fixtures"')
$lines.Add("' GENERATED FILE - do not edit. Source: Tests\fixtures\*.json")
$lines.Add("' Regenerate with: tools\New-FixturesModule.ps1")
$lines.Add("'")
$lines.Add("' Synthetic sample responses of the NinjaOne API used by the unit tests.")
$lines.Add('')
$lines.Add('Option Explicit')

$files = Get-ChildItem -Path $FixtureDir -Filter *.json | Sort-Object Name
if (-not $files) { throw "No fixtures found in $FixtureDir" }

foreach ($file in $files) {
    $raw = [System.IO.File]::ReadAllText($file.FullName)
    if ($raw -match '[^\x00-\x7F]') { throw "$($file.Name) contains non-ASCII characters" }

    # Minify: drop all whitespace outside of JSON strings.
    $sb = New-Object System.Text.StringBuilder
    $inString = $false
    $escaped = $false
    foreach ($ch in $raw.ToCharArray()) {
        if ($inString) {
            [void]$sb.Append($ch)
            if ($escaped) { $escaped = $false }
            elseif ($ch -eq '\') { $escaped = $true }
            elseif ($ch -eq '"') { $inString = $false }
        }
        elseif ($ch -eq '"') { $inString = $true; [void]$sb.Append($ch) }
        elseif (-not [char]::IsWhiteSpace($ch)) { [void]$sb.Append($ch) }
    }
    $compact = $sb.ToString()
    $func = ConvertTo-PascalCase ([System.IO.Path]::GetFileNameWithoutExtension($file.Name))

    $lines.Add('')
    $lines.Add("Public Function $func() As String")
    $lines.Add('    Dim s As String')
    for ($i = 0; $i -lt $compact.Length; $i += $chunkSize) {
        $chunk = $compact.Substring($i, [Math]::Min($chunkSize, $compact.Length - $i))
        $lines.Add('    s = s & "' + $chunk.Replace('"', '""') + '"')
    }
    $lines.Add("    $func = s")
    $lines.Add('End Function')
}

$text = ($lines -join "`r`n") + "`r`n"
[System.IO.File]::WriteAllText($OutFile, $text, (New-Object System.Text.ASCIIEncoding))
Write-Host "Wrote $OutFile ($($files.Count) fixtures)"
