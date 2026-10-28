# Local configuration

The repository contains neutral examples only. Configure installation-specific values before importing into an existing environment. The `.invalid` example domain is not a production API endpoint.

## Location and format

1. Start with `config/settings.example.ini`.
2. Save the customized file as `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\settings.ini`. The VBA module `outlook/LocalSettings.bas` reads it without changing registry or Outlook security settings.
3. For non-ASCII characters, such as sender names, save as **UTF-16 LE with a byte order mark (BOM)**. The example itself is ASCII. The Windows INI API recognizes Unicode through the BOM; UTF-8 with accented characters is unsuitable for this file.
4. **Do not enter credentials.** The client ID and secret remain exclusively in Windows Credential Manager. `CredentialTarget` is only the name of the generic entry.

| Key | Meaning |
|---|---|
| `ApiBaseUrl` | Your HTTPS NinjaOne API base, with a trailing slash and no user information, query, or fragment |
| `CredentialTarget` | Name of the existing generic Credential Manager entry; default: `NinjaRMM-Outlook-Classic/API` |
| `AllowedSender` | Full display name of ticket notifications; this is not an SMTP-address filter |
| `TicketSubjectPattern` | Regular expression with two capture groups: ticket number and remaining subject |
| `TicketPrefixPattern` | `TICKET (...) / ` prefix to shorten; capture groups contain the department and separator |
| `OrganizationPrefixPattern` | Optional organization prefix to remove from free-form subjects |
| `ExcludedDomains` | Comma-separated excluded domains |
| `AllTicketsBoard` | Positive numeric ID of the board used for ticket queries |

The five matching levels and offline fallback remain configured in `TicketConfig.bas`. The token endpoint is still the public EU endpoint in `NinjaAuth.bas`; select an endpoint appropriate for the NinjaOne application when using another region. A `TokenUrl` INI key is not evaluated.

## Upgrade an existing installation

Back up existing non-secret settings locally before importing the cleaned source. Installation-specific backups do not belong in releases or ZIP files and must not be forcibly added to Git. Deliberately copy the template to the runtime path above, customize it, and verify it before a production import. This repository does not include private settings.

Import `LocalSettings.bas` as an additional module. Replace `NinjaAuth`, `TicketConfig`, `TicketMatcher`, and every changed test module with files from the same revision. Run **Debug > Compile**. Do not rename or recreate credentials if `CredentialTarget` already points to the existing entry. Restart Outlook after configuration changes so token and index caches use the new settings.

Each self-test uses neutral defaults and synthetic temporary files. Normal settings are read again after the test ends. `RunLiveTests` is separate and explicitly opt-in; its ticket IDs and external reference are specified through the optional `Live*` keys. Unit-test example IDs are not recommendations for live queries.

## PowerShell smoke test

The read-only smoke test requires an explicit API base and credential target; it has no default tenant:

```powershell
.\tools\ninja-smoketest.ps1 -ApiBase 'https://tenant.example.invalid/v2/' -CredentialTarget 'NinjaRMM-Outlook-Classic/API'
```

Replace the example URL with your own API base. Supply open/closed ticket IDs as parameters when needed. Do not commit references from internal support cases.
