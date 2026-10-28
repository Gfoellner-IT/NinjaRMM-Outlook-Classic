# VBA Outlook Ticket Management with NinjaRMM Integration

This VBA project automates the handling and archiving of tickets in Microsoft Outlook.
It also integrates with the NinjaRMM API to determine if a ticket is closed based on its status from the remote system, rather than parsing status text from emails.

> **Source preview (2026-10-28):** VBA automation for classic Outlook, including ticket sorting, archiving, matching, audit and confirmation dialogs. A separate .NET Framework 4.8 core starts the VSTO migration; a diagnostic Outlook-VSTO host is included; an installer and functional Outlook integration are not yet available. Configure the installation locally and validate the complete VBA project before production use.

Documentation is split by purpose: [CHANGELOG.md](CHANGELOG.md) records completed changes and historical evidence, this README describes current behavior and operation, and [ROADMAP.md](ROADMAP.md) contains open work and future milestones. Contributor conventions are in [CLAUDE.md](CLAUDE.md). All repository documentation is maintained in English (United States, en-US); runtime localization is unchanged.

Repository: [GitHub project](https://github.com/Gfoellner-IT/NinjaRMM-Outlook-Classic). Installation settings are local; see [configuration](docs/configuration.md).

The instructions below describe the source in this branch, including M6–M8. To upgrade an Outlook project that contains only M0–M5, follow [IMPORT.md](IMPORT.md). Back up the VBA project before importing; use the isolated test PST for acceptance.

## Overview

- **RunEmailRule**: Sorts all ticket mails of the inbox into ticket folders (`Tickets\#<number> (<subject>)`).
- **RunEmailRuleScript**: For the Outlook rule action "Run a script": sorts only the mail the rule hands over, if it is still in the inbox.
- **RunArchiveRule**: Moves folders of closed tickets into `Archiv\<year>\<month>`. Whether a ticket is closed comes from the NinjaRMM API, not from mail texts.
- **Offline fallback**: If the API cannot be reached, a folder is still archived when one of its mails says "Status: ... Geschlossen" (archive date = date of that mail). Switch: `TicketConfig.ENABLE_OFFLINE_FALLBACK`.

## Key Functions

### Structure

| Module | Purpose |
|---|---|
| `Ticketautomatismus.bas` | Outlook entry points only (`RunEmailRule`, `RunEmailRuleScript`, `RunArchiveRule`, `ProcessEmail` for old callers). Rule entry points log problems to `TicketLog`; the explicit progress command displays a form. |
| `TicketService.bas` | The logic without Outlook types: `ProcessMail` and `ArchiveClosedTickets`. |
| `TicketParsing.bas`, `TicketDecision.bas`, `TicketConfig.bas`, `TimeUtil.bas` | Pure helpers: ticket numbers and folder names, archive decision, configuration, UTC to local time. |
| `ITicketStore`, `ITicketFolder`, `IIncomingMail` + `OutlookTicketStore`, `OutlookTicketFolder`, `OutlookIncomingMail` | Interfaces to the mailbox and the thin Outlook adapters that implement them. |
| `NinjaApi.bas`, `NinjaTicketLogic.bas`, `NinjaTransport.cls`, `JsonLite.bas` | Read-only NinjaOne queries (see below). |
| `NinjaAuth.bas` and helpers | Authentication (see *Authentication*). |

### Archive decision (`TicketDecision.DecideArchive`)

| API state | Result |
|---|---|
| closed or deleted | archive (date from the ticket log) |
| not found (HTTP 404) | archive (date = today) |
| open | stays, mail texts are never consulted |
| unavailable (offline, timeout, 5xx, auth problem) | archive only if the offline fallback is on and a mail says "Status: ... Geschlossen"; otherwise nothing happens and the next run tries again |

After three "unavailable" answers in a row the API is not asked again in the same run, so an offline PC does not wait for timeouts once per folder.

### In `NinjaApi.bas` (with `NinjaTicketLogic.bas`, `NinjaTransport.cls`, `JsonLite.bas`)
- **GetTicketInfo / GetTicketState**
  - Reads `ticketing/ticket/<id>` (one request, state and subject) and returns open, closed (status id 6000 or deleted), not found (HTTP 404) or unavailable.
- **IsTicketClosed**
  - `True` for closed, deleted and not-found tickets, `False` for open tickets and whenever the state is unavailable.
- **GetTicketClosedDate**
  - Reads `ticketing/ticket/<id>/log-entry?type=SAVE` and takes the newest entry that changed the status to "Geschlossen"/"Closed" (or deleted the ticket). The time is UTC (converted to local time before the archive month is chosen); `0` means unknown.
- **GetTicketSubject**
  - Current subject of the ticket, used to keep the folder name up to date.

## How It Works

1. **RunEmailRule / RunEmailRuleScript**: a mail from the allowed sender with a subject like `[example] (#3621) ...` is moved into the active folder of that ticket. If the folder is in the archive it is moved back, otherwise it is created. The folder name is aligned with the current ticket subject from the API.
2. **RunArchiveRule**: every active ticket folder is checked against the API (see the decision table) and moved to `Archiv\<year>\<month>` if the ticket is closed.

## Matcher defaults and controls

All five matching rules default to level 1 (log only), capped by `MATCH_LEVEL_MAX = 3`. Use `RunMatchDryRun` to inspect the newest mail without moving or queuing it; `RunMatchReview` collects and reviews proposals when a rule is explicitly set to level 2. `RunManualAssignmentScan` records manual assignments, and `RunMatchRulesReview` creates the aggregate report and copies a review prompt. `REVIEW_DIALOG = 2` selects the list form with message-box fallback.

These settings affect foreign-mail matching. Existing ticket-notification sorting and archiving can still move mail. No matching level or exclusion was changed by the integration.

## Configuration (`TicketConfig.bas`)

- `FOLDER_TICKETS` / `FOLDER_ARCHIVE`: names of the Outlook folders (`Tickets` below the inbox, `Archiv` below `Tickets`).
- `AllowedSender()`: sender display name from local settings.
- `REGEX_SUBJECT_PATTERN`: pattern that finds the ticket number in the subject (`^\[example\]\s*\(#(\d+)\)(.*)`).
- `RegexTicketReplace()`: shortens `TICKET (B..ro - IT) / ` to `IT / ` in new folder names.
- `REGEX_STATUS_CLOSED`: mail text that marks a closed ticket (offline fallback only).
- `ENABLE_OFFLINE_FALLBACK`: `True` (default) or `False`.

## Local installation settings

This source uses neutral defaults and an intentionally invalid example API URL. Configure `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\settings.ini` before importing it into an existing installation. Import the new `LocalSettings.bas` module as well. See [configuration and migration](docs/configuration.md) and [publication scope](PUBLICATION.md). No production Outlook installation is modified by editing this repository.

## Requirements

1. **Classic Outlook for Windows** with VBA support (verify the Office version and bitness before use).
2. No further VBA references are needed (regular expressions and HTTP are late bound).
3. A NinjaOne machine-to-machine application (platform *API Services*, grant *Client Credentials*, scope `monitoring`) and a generic entry in the Windows Credential Manager (see *Authentication* below).

## Installation

1. Open Outlook's VBA editor (`Alt+F11`) and import these files (*File > Import File*):
   - **From `outlook\` (all production modules; exclude `OutlookTestGuard.bas`):** `AuditLogic.bas`, `AuditReport.bas`, `ClipboardHelper.bas`, `FileAuditSink.cls`, `FileRejectionStore.cls`, `FormMatchPrompt.cls`, `GermanText.bas`, `HttpTokenFetcher.cls`, `IAuditSink.cls`, `IIncomingMail.cls`, `IMatchPrompt.cls`, `INinjaTransport.cls`, `IProgressSink.cls`, `IRejectionStore.cls`, `ITicketFolder.cls`, `ITicketIndex.cls`, `ITicketStore.cls`, `ITokenFetcher.cls`, `JsonLite.bas`, `LearnedRules.bas`, `LocalSettings.bas`, `ManualScan.bas`, `MatchLevels.bas`, `MatchPromptText.bas`, `MatchProposal.cls`, `MatchProposals.bas`, `MatchRejections.bas`, `MatchReview.bas`, `MatchTicket.cls`, `MsgBoxMatchPrompt.cls`, `NinjaApi.bas`, `NinjaAuth.bas`, `NinjaCredentialStore.bas`, `NinjaTicketIndex.cls`, `NinjaTicketLogic.bas`, `NinjaTokenLogic.bas`, `NinjaTransport.cls`, `OutlookIncomingMail.cls`, `OutlookTicketFolder.cls`, `OutlookTicketStore.cls`, `ProgressHelper.bas`, `ProgressSinkForm.cls`, `TicketAudit.bas`, `Ticketautomatismus.bas`, `TicketConfig.bas`, `TicketDecision.bas`, `TicketLog.bas`, `TicketMailWatcher.cls`, `TicketMatcher.bas`, `TicketParsing.bas`, `TicketService.bas`, `TimeUtil.bas`
   - **From `forms\`:** `ProgressForm.frm` and `ReviewForm.frm` (each needs its matching `.frx` file in the same folder).
2. No VBA reference has to be enabled besides the default Outlook library. All files are plain ASCII, so the import cannot garble umlauts.
3. Create the non-secret local settings file using [docs/configuration.md](docs/configuration.md). Keep credentials in Windows Credential Manager. Check `TicketConfig.bas` for folder names, matching levels and offline fallback.
4. Create the Credential Manager entry (see *Authentication* below).
5. Make sure the folder `Tickets` exists below the inbox.
6. Optional: start the mail watcher from `Application_Startup`, add buttons for `RunArchiveRuleWithProgress` and `RunEmailRule` (*File > Options > Customize Ribbon*), see *Automatic Execution on Email Arrival*.

Problems are written to `%TEMP%\ninja_ticket_log.txt` (the rules show no dialogs).

### Installing the tests (optional)

1. Create a local, empty data file `NinjaTest.pst` (*New Items > More Items > Outlook Data File*, display name `NinjaTest`, no account) with the folder `Posteingang\Tickets`. The integration tests work only there; `OutlookTestGuard` refuses every other store.
2. Import everything from `Tests\` plus `outlook\OutlookTestGuard.bas`.
3. Run `TestRunner.RunAllTests` and `IntegrationRunner.RunIntegrationTests`; `TestRunner.RunLiveTests` is an opt-in test against the real NinjaOne API.

## Usage

- **RunEmailRule**: Manually invoked inbox sweep, also called when the mail watcher starts. For a script rule, use `RunEmailRuleScript`.
- **RunEmailRuleScript**: Helper macro for the "Run a script" Outlook rule action. It processes only the message supplied by the rule if it is still in the Inbox.
- **RunArchiveRuleWithProgress**: Same as `RunArchiveRule`, with a non-modal progress window (for buttons).
- **StartMailWatcher** / **StopMailWatcher**: Event-driven sorting of new mails (see *Automatic Execution on Email Arrival*).
- **RunArchiveRule**: Manually invoked or scheduled (e.g., daily) to move closed tickets into the archive subfolders.

## Notes

- The project uses API state (`statusId = 6000`) first. Mail text is consulted only when the API is unavailable and the offline fallback is enabled.
- For archiving, the date used is the time of the newest log entry that changed the status to closed (works for manual and automatic closing).
- The subfolder naming convention for the archive is `<Year><Month>` with the month as `mm - mmmm` (e.g. `2023\04 - April`, the month name follows the system language).

## Authentication

The API is accessed with the OAuth2 *client credentials* flow: no browser, no login dialog, no refresh token. The client id and secret are read from the Windows Credential Manager (*Control Panel -> Credential Manager -> Windows Credentials -> Add a generic credential*):

| Field | Value |
|---|---|
| Internet or network address | The name configured as `CredentialTarget` (default: `NinjaRMM-Outlook-Classic/API`) |
| User name | client id of the API Services application |
| Password | client secret |

The token endpoint is `https://eu.ninjarmm.com/ws/oauth/token`. A token is valid for one hour and is cached in memory until shortly before it expires. See `outlook/NinjaAuth.bas`.

## Tests

- **Headless unit tests (no Outlook):** `tools\Invoke-VbaTests.ps1` runs the pure-logic tests in a throw-away hidden Excel instance. It needs the Excel setting "Trust access to the VBA project object model" (`AccessVBOM = 1`) and never touches your mailbox or Outlook VBA project. Modules are listed in `tools\unit-test-modules.txt`.
- **Unit tests in the VBA editor:** import all test modules and their dependencies from the same source revision (the generated runner references more than the framework self-tests), then run `TestRunner.RunAllTests` (F5). The result is shown in the Immediate window and written to `%TEMP%\ninja_test_results.txt`.
- **Fixtures:** `Tests\fixtures\*.json` are the source; regenerate the VBA module with `tools\New-FixturesModule.ps1`.
- **Live smoke test (read-only):** `tools\ninja-smoketest.ps1 -ApiBase <https-base-url> -CredentialTarget <target> -OpenTicketId <id> -ClosedTicketId <id>` checks the API credentials stored in the Windows Credential Manager.
- **Integration tests of the Outlook adapters (manual, F5):** `IntegrationRunner.RunIntegrationTests` runs `Tests/OutlookAdapterTests.bas` against a local, empty data file `NinjaTest.pst` (display name `NinjaTest`, folder `Posteingang\Tickets`, not connected to any account). `OutlookTestGuard` refuses to run against any other store. The result is written to `%TEMP%\ninja_integration_results.txt`. These modules use Outlook types and are not part of the headless run.
- **File check:** `tools\Test-VbaFiles.ps1` verifies CRLF line endings and the ASCII-only rule for new files.

## Automatic Execution on Email Arrival

There are two ways to sort new mails automatically. Use the first one.

### Option A (recommended): mail watcher, no security change

`Ticketautomatismus.StartMailWatcher` listens to the inbox (event `Items.ItemAdd`) and sorts every arriving ticket mail within about a second, without any rule. Start it when Outlook starts: open the VBA editor, double-click **ThisOutlookSession** and add

```vb
Private Sub Application_Startup()
    Ticketautomatismus.StartMailWatcher
End Sub
```

`StartMailWatcher` also runs `RunEmailRule` once, because Outlook does not raise `ItemAdd` when a large number of mails arrive at once (Microsoft Learn). Run `RunEmailRule` now and then as a sweep (e.g. from a button) if you want to be sure nothing is left in the inbox. The wait time before a new mail is processed is `TicketConfig.EVENT_DELAY_MS`.

### Option B: rule action "Run a script"

The rule action *Run a script* requires that the registry value `HKCU\Software\Microsoft\Office\16.0\Outlook\Security\EnableUnsafeClientMailRules` (DWORD) is `1` and Outlook was restarted. This is a security setting, so it is up to you. Then:

1. Create a rule for incoming messages ("When a message is received").
2. As the action choose **Run a script** and select `RunEmailRuleScript`.
3. Move this rule to the very bottom of the list so that all other rules are applied first.

`RunEmailRuleScript` checks whether the passed message is still in the inbox and only then moves it to the ticket folder (`TicketService.ProcessMail`).

## Progress window

`Ticketautomatismus.RunArchiveRuleWithProgress` is `RunArchiveRule` with a small non-modal progress window (`UserForm.Show vbModeless`): Outlook stays usable while it runs. Outlook has no documented VBA access to its own status bar, so a form is used. Use it for buttons; scheduled or rule-triggered runs use `RunArchiveRule`, which shows nothing.

## Matching, audit and review (M6–M8)

The features below are included in this branch. Historical evidence is recorded in [CHANGELOG.md](CHANGELOG.md); remaining Outlook acceptance work is listed in [ROADMAP.md](ROADMAP.md). Import instructions are in [IMPORT.md](IMPORT.md).

- M6 matches foreign mail against existing open tickets using Ninja ID, external ticket number, conversation and normalized subject; learned rules add a fifth rule. The index caches API results and pages board queries. Conversation assignments are currently kept in memory.
- M7 records decisions in monthly JSONL audit files under `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\audit`, detects manual assignments and generates review reports. `RunMatchRulesReview` writes a report and copies an introductory prompt to the clipboard. New rule candidates are disabled until explicitly approved.
- M8 controls each rule with levels 0 (off), 1 (log only), 2 (confirm) and 3 (automatic), capped by `MATCH_LEVEL_MAX`. All five rule defaults are 1; the global cap is 3. `RunMatchReview` opens proposals using `ReviewForm` (`REVIEW_DIALOG = 2`), falling back to individual message boxes. `RunMatchDryRun` never moves mail or queues proposals.

Log-only applies to the foreign-mail matcher. The existing ticket-notification sorter and archive command can still move mail. Audit logs and reports can contain domains, ticket IDs and shortened matched subjects; they are not guaranteed anonymous. No mail bodies or credentials should be logged.

## M9 C# migration

The solution [dotnet/NinjaRmm.Outlook.sln](dotnet/NinjaRmm.Outlook.sln) contains five .NET Framework 4.8 projects:

- **Core:** archive policy, parsing, five matching rules, safety decisions, UTC conversion, API response parsing, and a session-only snapshot index.
- **Infrastructure:** non-secret local INI settings, read-only Windows Credential Manager access, OAuth client-credentials token caching, one refresh/retry after 401, ticket/log reads, and bounded board-query paging. It is not connected to the Outlook host.
- **Core.Tests / Infrastructure.Tests:** executable synthetic reference tests with shared VBA fixtures and simulated HTTP responses; no live API or real credentials.
- **AddIn:** a regular Visual Studio template-based Outlook VSTO host with a diagnostic Ribbon callback. It performs no mail processing or API access. Default build/clean compiles the DLL without generating VSTO deployment manifests, registering the add-in, or changing certificate trust.

VBA remains the working implementation. The full review/audit/persistence workflow, Outlook mail/folder/event adapters, installer, and business acceptance remain open. Build with `tools/Test-DotNetCore.ps1`; `-RuntimeOnly` remains a limited core diagnostic.

Use [M9 architecture](docs/m9/architecture.md), [reference cases](docs/m9/reference-cases.md), and the [development host procedure](docs/m9/development-host.md) for current boundaries and the separate signing/registration/host acceptance steps. Completed checks are recorded in [CHANGELOG.md](CHANGELOG.md).
