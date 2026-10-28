# M9 development host and acceptance procedure

## Current scope

The solution targets .NET Framework 4.8 and Visual Studio Community 2022. The Outlook project was created through Visual Studio's installed C# Outlook VSTO Add-in wizard, with its generated host/designer code. It references the pure core and exposes a **NinjaRMM M9 > Host status** diagnostic command for the Explorer ribbon.

The host has no mail watcher, mailbox enumeration, API calls, or assignment/archive commands. Startup only writes a fixed diagnostic marker. VBA remains the operational implementation. The infrastructure library is not connected to the host. An infrastructure object does not send requests on construction; explicit methods can read API data or the configured credential entry.

## Build and tests

From the current clean checkout:

```powershell
.\tools\Test-DotNetCore.ps1
```

This builds all five projects with Visual Studio MSBuild and the net48 targeting pack, then runs core and infrastructure tests. Tests use shared synthetic VBA JSON fixtures, in-memory HTTP handlers, and a unique temporary INI file. They do not call NinjaOne, read actual credentials, start Office, or modify the Outlook VBA project. The temporary INI is removed by the test runner.

Default host build/clean skips the installed VSTO targets that generate deployment manifests, manage certificates, register the add-in/form regions, and alter the VSTO inclusion list. It compiles/copies DLL outputs only. Publishing is blocked by default. This deliberately separates compilation from developer registration and trust; a DLL build is not a loadable signed deployment or an Outlook startup test.

`-RuntimeOnly` remains a limited core diagnostic and does not build the infrastructure/host or establish net48 project compatibility.

## First actual Outlook host test (requires developer setup)

1. Review this branch and configure a development signing certificate in Visual Studio's project signing settings. Keep keys/PFX files local and ignored. Choose development signing separately from the eventual organizational publisher trust. Do not change macro security or install trusted certificates merely to compile.
2. After reviewing the certificate/trust implications, explicitly enable the standard installed VSTO build lifecycle with the MSBuild property `EnableVstoDevelopmentSetup=true`. Configure signing before doing this. The flag enables the normal VSTO manifest, certificate, registration, and inclusion-list targets; it is not a no-side-effect build switch. Do not commit personal signing properties, `.csproj.user`, or private key files.
3. Save work in Outlook and restart it deliberately when ready. Do not terminate a running user instance or assume a second process is an isolated mailbox. The initial host performs no mail processing, so its diagnostic test does not require disabling the existing VBA watcher. Before later enabling overlapping processing, back up VBA and disable the overlapping entry points deliberately.
4. Verify the add-in is loaded in classic Outlook, the **NinjaRMM M9** Explorer tab is visible, and **Host status** reports the core reference check as OK. Check `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\vsto\host-diagnostics.log` for `startup-no-mail-processing`, `ribbon-requested`, `ribbon-loaded`, and `core-check-passed`. Fixed markers contain no message content. The log is diagnostic evidence, not a business audit report.
5. Record build revision, signing/development registration choices, Office bitness/version, cold/warm starts, and callback behavior. A visible template or successful DLL build alone does not satisfy this step.
6. For rollback, disable/remove this development add-in registration deliberately. Do not delete unrelated registry entries, inclusion-list records, or certificates. Preserve the working VBA project.

No certificate creation, trust change, add-in registration, Outlook restart, or host acceptance is performed by the default test command. Actual host startup remains unverified in this revision.

## Remaining migration work

The intended first functional VSTO release preserves M0-M8 behavior: notification sorting/archive restore, API archive decisions and offline fallback, five matching rules, confirmation/rejections, audit/report/manual-assignment scans, and approved learned rules. Source/reference cases define the starting specification; full business parity is not yet accepted.

- Add Outlook mail/folder adapters, a stable message-key strategy, event/sweep coordination, and guarded test-PST integration. Recheck ticket state immediately before every move.
- Connect API reads and snapshot refresh/cache behavior. Cancellation, HTTP failures, invalid/partial boards, and regex timeouts must not authorize a move. `GetTicketAsync` returns null for unavailable/not-found data or propagates transport/authentication errors; archive orchestration must use the four-state parser and explicitly map failures to Unavailable.
- Implement proposal queue/review UI, persistent rejections, audit/report generation, and rule-file loading. Learned rule matching is ported, but file parsing/reload and report-based rule management are not connected.
- Version persisted formats and test migration/recovery on copies before writing existing VBA data. Conversation assignments remain session-only; persistence across restart is an open decision.
- Select deployment/update sources, stable package identity/versioning, organizational signing, and pilot rollback before distribution. No installer or supported binary release is provided here.
- M10's wrapping read-only details text box remains after M9; this diagnostic Ribbon does not implement it.

## Deliberate port differences to review

- ASCII ticket digits and regex timeouts preserve bounded parsing; timeout exceptions are propagated to the caller rather than treated as permission to act.
- Unknown UTC dates use null. An explicitly supplied time zone handles month boundaries and daylight saving time.
- Board rows marked deleted are never open, even if their status ID is not 6000. Invalid row IDs are skipped; unreadable statuses are non-open. Repeated cursors and the 50-page cap fail with an incomplete-snapshot error rather than exposing a partial index as complete.
- Newly constructed learned-rule candidates are disabled unless explicitly enabled. This follows the approval policy; persistence/migration of older rules requires its own tests.

The Ribbon uses the documented [VSTO extensibility override](https://learn.microsoft.com/en-us/visualstudio/vsto/ribbon-xml). Reference parsing follows the repository's VBA implementation and synthetic fixtures, not a newly verified live tenant API schema.
