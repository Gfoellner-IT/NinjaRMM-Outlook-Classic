# M9: architecture and target environment

Target architecture for the started M9 milestone. The VBA implementation remains the functional reference; an Outlook VSTO host, installer, and deployment are not yet available.

## Target

- C# and **.NET Framework 4.8** are the initial shared target for the core and future VSTO host. An installed runtime does not replace a targeting pack.
- Classic Outlook on Windows; the first pilot targets **64-bit Office**. The core and reference tests are AnyCPU. Verify the host, interop, VSTO runtime, and future native dependencies against the target environment. 32-bit Office has not been accepted.
- Use a regular Outlook VSTO project from a Visual Studio template, generated host/designer code, and standard build/publish tools. Do not substitute a hand-written host or copy Office assemblies from local pilot projects.
- M0-M8 define the intended functional starting point: sorting, archiving, matching, safety levels, proposals, audit, and rules review. Partial ports do not enable production actions. M10's additional detail text box remains after the migration.

## Prerequisites and evidence limits

The first target environment is Windows with 64-bit classic Outlook. A regular build requires Visual Studio with Office/VSTO development tools and the .NET Framework 4.8 targeting pack. Installing Framework or VSTO runtimes alone does not provide those tools.

The core's limited runtime test path is independent of the Office host. It does not establish a regular net48 build, installation, or Outlook add-in startup. Recorded snapshot checks are in [PUBLICATION.md](../../PUBLICATION.md).

## Current code structure

```text
dotnet/NinjaRmm.Outlook.sln
  NinjaRmm.Outlook.Core          pure logic; no Office or network access
  NinjaRmm.Outlook.Core.Tests    executable reference tests; no NuGet/Office/credentials

not yet implemented:
  Infrastructure                API, tokens, Credential Manager, files, migration
  Outlook VSTO host             Outlook adapters, Ribbon, events, review dialog
```

The VBA files remain in place. The initial C# core implements `ArchivePolicy`, `MatchLevels`, and `TicketParsing`. The API, full matcher, audit, storage, proposals, and dialogs have not yet been ported.

API and file operations will use separate interfaces. Keep Office objects in the host/STA context. HTTP and data work require timeouts, cancellation, and controlled return to the UI context. Startup must prevent duplicate VBA/add-in processing.

## Explicit type and error-handling decisions

- Represent unknown closing dates as `DateTime? = null`, rather than VBA date value `0`. Archive functions receive local dates, the current day, and culture explicitly; UTC conversion remains separate work.
- Invalid numeric ticket states throw exceptions. VBA's fallback `Case Else` is not extended to invalid .NET enum values. The future API adapter must explicitly map failures to `Unavailable`.
- Configuration belongs to instances instead of global test overrides. All five rules start at `LogOnly`, with a global cap of `3`, as in VBA.
- Regex timeouts limit processing; a timeout is not a match. Future callers must log it safely. ASCII digits preserve ticket-number semantics; folder-prefix comparisons are ordinal.
- Version audit, rules, and rejection data before the first .NET write. Test migration with backups and recovery. This foundation does not read or change existing user files or credentials.
- Persistent conversation assignments across restarts remain an open decision. The current in-memory implementation is not persistent storage.

## Regular development tools

Use a Visual Studio edition whose license permits the intended use, with Office/VSTO tools. For example, Visual Studio 2022 with:

- Workload `Microsoft.VisualStudio.Workload.Office` (Office/SharePoint development).
- `Microsoft.VisualStudio.Component.TeamOffice` (VSTO), ClickOnce tools, and MSBuild.
- .NET Framework 4.8 SDK and targeting pack/developer tools.
- Appropriate Outlook interop and Office references from the regular development environment.

After setup, run `tools/Test-DotNetCore.ps1` without the diagnostic switch. Then add a regular Outlook VSTO add-in project to the solution and validate actual host/Ribbon startup separately.

Sources checked on October 27, 2026: [Microsoft: Office development environment](https://learn.microsoft.com/en-us/visualstudio/vsto/how-to-configure-a-computer-to-develop-office-solutions?view=visualstudio), [Microsoft: Visual Studio 2022 components](https://github.com/MicrosoftDocs/visualstudio-docs/blob/main/docs/install/includes/vs-2022/workload-component-id-vs-community.md). Linking the Community component list does not determine licensing eligibility for a particular organization or use case.
