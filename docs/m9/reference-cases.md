# M9: reference behavior and porting evidence

The VBA specification comes from the existing code and tests at the same revision. Cases are synthetic; do not carry over production email text or credentials. Record completed checks in [CHANGELOG.md](../../CHANGELOG.md) and open work in [ROADMAP.md](../../ROADMAP.md).

## Test scope

VBA unit tests and C# reference tests use synthetic data. Separately registered Outlook integration tests must run only in the isolated test PST. Recorded results for the source snapshot are in [PUBLICATION.md](../../PUBLICATION.md). Business matching accuracy and actual Outlook UI behavior require separate evidence.

## Reference matrix

| Area | Required reference behavior | VBA source | C# port status |
|---|---|---|---|
| Archive decisions | Archive Closed/NotFound through the API; keep Open regardless of old status emails; use Unavailable fallback only when enabled and a message indicates closure | `Tests/TicketDecisionTests.bas` | Initial implementation: all 16 state/email/fallback combinations and need-for-mail checks |
| Safety levels | Five rules, default 1, global cap 3; take the minimum and clamp to 0-3; do not share global state between instances | `Tests/MatchLevelsTests.bas` | Initial implementation: every rule tested with configured levels and caps from -1 through 4 |
| Ticket/folder parsing | Subject tag at the start, case handling, cleared outputs on failure, prefix shortening, and no confusion between short and longer IDs | `Tests/TicketParsingTests.bas` | Initial reference cases implemented |
| Archive dates | Unknown date uses the supplied current day; month names follow the local culture | `Tests/TicketParsingTests.bas`, `Tests/TimeUtilTests.bas` | Local date/culture formatting implemented; UTC conversion remains open |
| API and tokens | States, closing dates, deleted/missing tickets, retry after 401, timeouts, token cache, and credential errors without secret disclosure | `Tests/NinjaApiTests.bas`, `Tests/NinjaTicketLogicTests.bas`, `Tests/NinjaAuthTests.bas`, `Tests/NinjaTokenLogicTests.bas` | Open |
| Ticket index and matcher | Paging, cache, rule priority, exclusions, ambiguity; never move closed or unreadable matches | `Tests/NinjaTicketIndexTests.bas`, `Tests/TicketMatcherTests.bas`, `Tests/TicketMatchServiceTests.bas` | Open |
| Proposals and decisions | A dry run never moves or queues messages; confirmation repeats protection checks; rejections persist; cancellation makes no decision | `Tests/MatchReviewTests.bas`, `Tests/MatchLevelsTests.bas` | Open |
| Audit and rules | Stable message keys, monthly files/retention, nonfatal write failures, explicit approval before rule activation | `Tests/AuditTests.bas`, `Tests/AuditReportTests.bas`, `Tests/ManualScanTests.bas`, `Tests/LearnedRulesTests.bas` | Open |
| Outlook and UI | Events, stop/shutdown, mail/folder adapters, archive restore, form/fallback, and no test changes to the production mailbox | `Tests/OutlookAdapterTests.bas`, `Tests/ReviewFormTests.bas` | Regular VSTO host and separate acceptance remain open |

## Execution

After installing the toolchain, run the regular project build and C# reference tests:

```powershell
.\tools\Test-DotNetCore.ps1
```

For a limited local diagnostic without Visual Studio or a targeting pack:

```powershell
.\tools\Test-DotNetCore.ps1 -RuntimeOnly
```

The diagnostic compiles only the core and its tests against installed Framework runtime assemblies. It does not validate net48 reference APIs, the MSBuild project, VSTO host, manifest, installation, or Outlook behavior. These tests do not start Office, call the API, or use real user files. Diagnostic outputs are stored in the ignored `artifacts/core-runtime` directory.
