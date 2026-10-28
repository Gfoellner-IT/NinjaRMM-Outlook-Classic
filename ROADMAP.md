# Roadmap

Planning date: **October 28, 2026**. This file contains open work, planned milestones, dependencies, and acceptance criteria. [README.md](README.md) describes current behavior; [CHANGELOG.md](CHANGELOG.md) records completed changes and historical evidence. Maintain all repository documentation in **English (United States, en-US)**.

## VBA: acceptance of the published source

These tasks apply to the integrated source described in the README. Publishing source does not establish business acceptance; the following checks remain necessary.

| Open task | Acceptance criteria |
|---|---|
| Outlook import and full compilation | Import the matching source revision, back up the VBA project, and complete **Debug > Compile Project1** without errors |
| M6: matching and restart behavior | Verify unique open matches, ambiguous results, excluded/closed/unreadable tickets, dry runs, and restarts in the isolated test PST |
| Persistent conversation assignments | Decide whether assignments must survive Outlook restarts; if required, implement storage, recovery, and tests |
| M7: audit and manual assignments | Verify stable message keys after moves, nonfatal audit write failures, and reports based on explicitly approved decisions |
| M7: first rules review | Evaluate sufficient confirmations/rejections, review candidates, and enable only explicitly approved rules |
| M8: Outlook dialogs | Test the form and message-box fallback: accept, reject, skip/cancel, repeated invocation, grouped notifications, and audit records |

The [import guide](IMPORT.md) applies only to its stated source scope. Tests that create Outlook folders or messages must use the local, non-Exchange `NinjaTest` PST, never the production mailbox.

## M9: migrate VBA to a .NET VSTO add-in (started)

The goal is a C# VSTO add-in for **classic Outlook on Windows** that preserves the VBA solution's business behavior and supports a regular build, installation, and update process. VSTO uses **.NET Framework**. The working target of .NET Framework 4.8 and its validation limits are documented in the [architecture](docs/m9/architecture.md). Regular development tools and host validation remain prerequisites. New Outlook does not support COM/VSTO and requires a separate architecture decision. Sources checked on October 27, 2026: [Microsoft: create VSTO add-ins](https://learn.microsoft.com/en-us/visualstudio/vsto/create-vsto-add-ins-for-office-by-using-visual-studio?view=visualstudio), [Microsoft: Outlook add-in platforms](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-add-ins-overview).

### M9.1: define the target environment and reference behavior

The [architecture](docs/m9/architecture.md) and [reference matrix](docs/m9/reference-cases.md) are available. The regular Visual Studio/VSTO toolchain, net48 targeting pack, project build evidence, and business approval of the reference scope remain open.

- Install the regular Visual Studio/VSTO toolchain and net48 targeting pack, verify the target matrix with the actual VSTO host, and demonstrate the regular project build.
- Define the first add-in version's scope using the existing VBA reference cases. Unaccepted VBA behavior requires at least a documented specification and test cases before porting; a merge alone does not establish acceptance.
- Define reference cases and data for sorting, archiving, offline fallback, matching, review, audit, and rules. Explicitly decide open issues such as persistent conversation assignments.
- Acceptance: documented target matrix, scope, and associated reference cases; no installation is required for this definition step.

### M9.2: port business logic and the API to C#

The initial core is described in [README.md](README.md). The full matcher and ticket index, API/token/credential adapters, UTC conversion, proposal/review logic, audit, rules, and data migration remain to be ported. The runtime diagnostic does not replace a build against net48 reference assemblies.

- Keep parsing, archive decisions, matching, safety levels, proposals, audit, and rules in a host-independent library. Carry over synthetic fixtures and protection cases.
- Preserve read-only NinjaOne operations and the client-credentials flow, including token caching, timeouts, cancellation, and Credential Manager access. Keep secrets out of packages and logs.
- Version audit, rule, and rejection formats. Test reuse or migration with backups and a recovery path.
- Acceptance: automated build, unit tests, negative cases, and comparison with approved VBA reference cases. This does not establish Outlook acceptance.

### M9.3: Outlook host, Ribbon, and review UI

- Create a regular Outlook VSTO project with a generated host, startup/shutdown handling, Ribbon commands for sorting, archiving, proposals, and reports, and an appropriate review UI.
- Encapsulate mail events, sweeps, ticket folders, and the archive behind Outlook adapters. Keep COM access in the required host/STA context. Run HTTP and data work without blocking Outlook. Handle reentrancy, duplicate events, cancellation, and errors.
- Acceptance: actual add-in/Ribbon startup and business tests in the test PST; cover cold/warm starts, multiple windows, and shutdown.

### M9.4: signing, installation, updates, and rollback

- Select deployment, such as ClickOnce or a managed installer, user/machine scope, a stable add-in identity, and a version scheme. Coordinate code/manifest signing and publisher trust with IT.
- Document prerequisites, update sources, and uninstallation. Do not embed developer or repository write tokens. Test update failures, offline behavior, and rollback, including registration and configuration.
- Back up the VBA project. Disable overlapping VBA watchers, rules, and automatic entry points before enabling the pilot add-in to prevent duplicate processing. For rollback, disable the add-in and deliberately reactivate the backed-up VBA version.
- Acceptance: installation, update, uninstallation, actual host startup, and demonstrated rollback on a representative target machine. Package/manifest checks alone are insufficient.

### M9.5: pilot and replacement

- Start with logging, then confirmed assignments, and enable automatic processing per rule only after business approval. Preserve existing folders, data, and access settings.
- Review sorting, archive month selection, offline behavior, safe matching, reports, and audit records with the user. Record startup time and Outlook responsiveness.
- Retire VBA only after successful business acceptance. Record the result, package version, target environment, and remaining limitations.

## M10: readable proposal details after the VSTO migration

**Dependency:** completion of M9. Implement this display improvement in the VSTO UI.

- Keep the default proposal list compact. Use a separate read-only text box for the focused entry's full ticket subject, email subject, sender domain, and matching rule.
- Enable word wrapping and vertical scrolling so long text remains readable.
- Optionally provide an expandable plain-text message preview. Clearly label truncated excerpts; the current VBA interface returns at most 4,000 characters.
- Keep focused and checked proposals independent: showing details must not change check marks or assignment decisions.
- Acceptance: long subjects and multiline content are readable without horizontal clipping; mouse/keyboard navigation updates the correct details. Accept, reject, and close retain their existing behavior.

## Sequence and decisions

First define the initial VSTO version's scope in M9.1. If M6-M8 are released as a VBA extension, complete their open acceptance checks first. A direct port requires approved reference cases and expected behavior; a prior VBA release is not mandatory.

Sequence: M9.1 -> M9.2 -> M9.3 -> M9.4 -> M9.5 -> M10. This is not a delivery-date commitment. Open decisions include the target environment/toolchain, initial add-in scope, persistent conversation assignments, data migration, deployment, and update sources.

## Cross-cutting acceptance criteria

- Preserve sorting, archive structure, and manual assignments. Do not add ticket creation or write operations on NinjaOne tickets.
- Test API states, closing dates, localized status names, offline fallback, and safe matching. Verify current tenant permissions and connectivity during the corresponding live acceptance run.
- Run automated tests and negative cases for changed logic. Associate results with a commit, environment, and date. Excel/parser success does not establish Outlook acceptance.
- Verify VBA signing and trust before release. For VSTO, also verify the package, installation, host/Ribbon startup, and rollback.
- Keep credentials out of the repository and logs. Shortened subjects can still contain personal data; review audit records and reports before sharing them.
