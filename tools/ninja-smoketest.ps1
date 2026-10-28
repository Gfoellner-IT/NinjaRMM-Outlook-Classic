<#
.SYNOPSIS
  Read-only live check of the NinjaOne API with the client-credentials flow.

.DESCRIPTION
  Reads client id and secret from a generic entry of the Windows Credential
  Manager (user name = client id, password = client secret), requests a token
  and runs read-only GET calls. Secrets and tokens are never printed.
  Optionally checks one open and one closed ticket and derives the closing
  date the same way the planned NinjaApi module will: newest log entry whose
  status change goes to "Geschlossen"/"Closed".

  Exit code 0 = all checks passed, 1 = at least one check failed.

.EXAMPLE
  .\tools\ninja-smoketest.ps1 -CredentialTarget NinjaRMM-Outlook-Classic/API -ApiBase https://tenant.example.invalid/v2/ -OpenTicketId <id> -ClosedTicketId <id>
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$CredentialTarget,
    [string]$TokenUrl = 'https://eu.ninjarmm.com/ws/oauth/token',
    [Parameter(Mandatory=$true)][ValidatePattern('^https://[^@?#]+/$')][string]$ApiBase,
    [string]$Scope = 'monitoring',
    [long]$OpenTicketId = 0,
    [long]$ClosedTicketId = 0
)

$ErrorActionPreference = 'Stop'
$script:failed = 0

function Write-Check([bool]$ok, [string]$text) {
    if ($ok) { Write-Host "OK    $text" } else { Write-Host "FAIL  $text"; $script:failed++ }
}

if (-not ('NinjaSmokeCred' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class NinjaSmokeCred {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct CREDENTIAL {
        public uint Flags; public uint Type; public string TargetName; public string Comment;
        public long LastWritten; public uint CredentialBlobSize; public IntPtr CredentialBlob;
        public uint Persist; public uint AttributeCount; public IntPtr Attributes;
        public string TargetAlias; public string UserName;
    }
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool CredRead(string target, uint type, uint flags, out IntPtr cred);
    [DllImport("advapi32.dll")]
    static extern void CredFree(IntPtr cred);
    public static string[] Read(string target) {
        IntPtr p;
        if (!CredRead(target, 1, 0, out p)) return null;
        var c = (CREDENTIAL)Marshal.PtrToStructure(p, typeof(CREDENTIAL));
        string secret = Marshal.PtrToStringUni(c.CredentialBlob, (int)c.CredentialBlobSize / 2);
        string user = c.UserName;
        CredFree(p);
        return new[] { user, secret };
    }
}
'@
}

function ConvertTo-Utc([object]$unixSeconds) {
    [DateTimeOffset]::FromUnixTimeSeconds([long][Math]::Floor([double]$unixSeconds)).UtcDateTime
}

# --- 1. credential ----------------------------------------------------------
$cred = [NinjaSmokeCred]::Read($CredentialTarget)
Write-Check ($null -ne $cred -and $cred[0] -and $cred[1]) "Credential entry '$CredentialTarget' found (client id and secret present)"
if (-not $cred) { exit 1 }

# --- 2. token ---------------------------------------------------------------
$body = @{ grant_type = 'client_credentials'; client_id = $cred[0]; client_secret = $cred[1]; scope = $Scope }
try {
    $token = Invoke-RestMethod -Uri $TokenUrl -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded'
    Write-Check ($token.token_type -eq 'Bearer' -and $token.access_token) "Token request (type=$($token.token_type), expires_in=$($token.expires_in), scope=$($token.scope))"
} catch {
    Write-Check $false "Token request failed: HTTP $($_.Exception.Response.StatusCode.value__)"
    exit 1
}
$headers = @{ Authorization = "Bearer $($token.access_token)" }

function Get-Api([string]$path) {
    Invoke-RestMethod -Uri ($ApiBase.TrimEnd('/') + '/' + $path.TrimStart('/')) -Headers $headers
}

# --- 3. basic read calls ----------------------------------------------------
foreach ($path in 'organizations', 'ticketing/trigger/boards') {
    try { $r = Get-Api $path; Write-Check $true "GET $path (items: $(@($r).Count))" }
    catch { Write-Check $false "GET $path -> HTTP $($_.Exception.Response.StatusCode.value__)" }
}

# --- 4. ticket checks -------------------------------------------------------
function Test-Ticket([long]$id, [bool]$expectClosed) {
    try {
        $t = Get-Api "ticketing/ticket/$id"
        $isClosed = ($t.status.statusId -eq 6000)
        Write-Check ($isClosed -eq $expectClosed) "Ticket $id statusId=$($t.status.statusId) (expected closed=$expectClosed)"
        if ($isClosed) {
            $rawLog = Get-Api "ticketing/ticket/$id/log-entry?type=SAVE"
            $log = @($rawLog)
            $closing = $log | Where-Object {
                $_.changeDiff -and $_.changeDiff.status -and ($_.changeDiff.status.new -in @('Geschlossen', 'Closed'))
            } | Sort-Object { [double]$_.createTime } -Descending | Select-Object -First 1
            if ($closing) {
                Write-Check $true "Ticket $id closing date from log: $((ConvertTo-Utc $closing.createTime).ToString('yyyy-MM-dd HH:mm:ss')) UTC (automation entry: $([bool]$closing.automation))"
            } else {
                Write-Check $false "Ticket ${id}: no log entry with status change to closed ($($log.Count) entries)"
            }
        }
    } catch {
        Write-Check $false "Ticket $id -> HTTP $($_.Exception.Response.StatusCode.value__)"
    }
}

if ($OpenTicketId -gt 0) { Test-Ticket $OpenTicketId $false }
if ($ClosedTicketId -gt 0) { Test-Ticket $ClosedTicketId $true }

if ($script:failed -eq 0) { Write-Host 'RESULT: OK' } else { Write-Host "RESULT: FAILED ($script:failed)" }
exit ([int]($script:failed -gt 0))
