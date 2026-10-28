# Import M6 through M8 into Outlook and run the tests

**Scope: the published VBA source in this revision.** This guide upgrades an older M0–M5 project with matching and the ticket index (M6), audit logs, reports, manual assignments and learned rules (M7), and safety levels and review dialogs (M8). Configure the installation locally and compile and validate the complete project in Outlook before production use. The VSTO migration is separate work in the [roadmap](ROADMAP.md).

**The matcher for other email starts in log-only mode:** all five rule levels default to 1. Existing ticket-notification sorting and archiving can still move messages or folders. Levels 2 and 3 enable actions only for the matcher.

Historical test results are in the [changelog](CHANGELOG.md). Headless tests cover the documented logic and simulated form cases in Excel; they do not validate Outlook types or actual dialog behavior in Outlook. Complete step 3 and record the module, line, and message for any compilation error.

## 0. Prepare

**For the cleaned publication:** create the [local configuration](docs/configuration.md) before importing. Also import `outlook\LocalSettings.bas` and replace `NinjaAuth`, `TicketConfig`, `TicketMatcher`, and the updated test modules. Neutral defaults replace the earlier installation settings. Repository changes do not automatically update an existing Outlook installation.

1. Close Outlook.
2. Back up the VBA project:

   ```powershell
   Copy-Item "$env:APPDATA\Microsoft\Outlook\VbaProject.OTM" "$env:APPDATA\Microsoft\Outlook\VbaProject.OTM.bak-$(Get-Date -Format yyyyMMdd)"
   ```

3. Start Outlook and open the VBA editor with `Alt+F11`. Use the source files from your repository checkout.

## 1. Remove the old modules

In Project Explorer, right-click each listed module, choose **Remove**, and answer **No** when asked whether to export it (export separately first if needed). Keep `ThisOutlookSession` and modules not listed here.

**From `outlook\` (12 changed modules):** `IIncomingMail`, `INinjaTransport`, `ITicketFolder`, `NinjaTicketLogic`, `NinjaTransport`, `OutlookIncomingMail`, `OutlookTicketFolder`, `TicketConfig`, `TicketLog`, `TicketMailWatcher`, `TicketService`, `Ticketautomatismus`.

**From `Tests\` (three fakes, required if tests are installed):** `FakeNinjaTransport`, `FakeMail`, `FakeTicketFolder`. Their interfaces have changed. Old versions can produce an error that `HttpPostJson` must be implemented for `INinjaTransport`.

Importing without first removing the existing module creates a duplicate such as `Name1` and prevents compilation.

## 2. Import the updated files

Use **File > Import File** (`Ctrl+M`), one file at a time. Import order does not matter.

**From `outlook\` (36 M6–M8 files; local-configuration modules are covered in step 0):**

| New (24) | Changed (12) |
|---|---|
| `AuditLogic.bas`, `AuditReport.bas`, `ClipboardHelper.bas`, `FileAuditSink.cls`, `FormMatchPrompt.cls`, `FileRejectionStore.cls`, `GermanText.bas`, `IAuditSink.cls`, `IMatchPrompt.cls`, `IRejectionStore.cls`, `ITicketIndex.cls`, `LearnedRules.bas`, `ManualScan.bas`, `MatchLevels.bas`, `MatchPromptText.bas`, `MatchProposal.cls`, `MatchProposals.bas`, `MatchRejections.bas`, `MatchReview.bas`, `MatchTicket.cls`, `MsgBoxMatchPrompt.cls`, `NinjaTicketIndex.cls`, `TicketAudit.bas`, `TicketMatcher.bas` | `IIncomingMail.cls`, `INinjaTransport.cls`, `ITicketFolder.cls`, `NinjaTicketLogic.bas`, `NinjaTransport.cls`, `OutlookIncomingMail.cls`, `OutlookTicketFolder.cls`, `TicketConfig.bas`, `TicketLog.bas`, `TicketMailWatcher.cls`, `TicketService.bas`, `Ticketautomatismus.bas` |

**From `Tests\` (three files):** `FakeNinjaTransport.cls`, `FakeMail.cls`, `FakeTicketFolder.cls`. For a complete test update, follow step 10.

**From `forms\` (one recommended form):** `ReviewForm.frm`, with `ReviewForm.frx` in the same directory. This provides the list-based review dialog (dialog B). If the form is absent or cannot be imported, Outlook falls back to individual message boxes (dialog A) and logs a notice.

## 3. Compile

Choose **Debug > Compile Project1**. Continue only after compilation succeeds. Record the module, line, and error message if it fails.

## 4. Run a dry run

In the Immediate window (`Ctrl+G`), enter:

```vb
Ticketautomatismus.RunMatchDryRun
```

This dry run does not move messages or queue proposals. Results go to `%TEMP%\ninja_ticket_log.txt` and `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\audit\match_audit_<year>-<month>.jsonl`. Unmatched messages are logged using a subject hash (`subject_hash=...`); matched or ambiguous results can include subjects. Older audit files may still contain subjects; review and remove them according to your retention requirements.

## 5. Try the review dialog (level 2)

1. In `TicketConfig`, set one rule to level 2, for example `MATCH_LEVEL_SUBJECT = 2`, and save. Levels are 0 = off, 1 = log only, 2 = confirm, 3 = automatic. `MATCH_LEVEL_MAX` defaults to 3.
2. Run `Ticketautomatismus.RunMatchReview` in the Immediate window. It checks messages from the last 14 days. With the imported form and the default `REVIEW_DIALOG = 2`, select messages using the checkboxes and choose the accept action to move them into the ticket folder, or reject to prevent another proposal. Select all checks every entry. Closing preserves undecided proposals. The current runtime captions remain German. Without the form, or with `REVIEW_DIALOG = 1`, each proposal has a message box: **Yes** accepts, **No** rejects, and **Cancel** ends the review while preserving remaining proposals.
3. Verify that an accepted message arrives in the correct folder (`#<number> (<subject>)`) and a rejected message is not proposed again on the next run.
4. When new email produces proposals, the watcher can offer to open the review, at most once every 15 minutes. The startup sweep does not ask; use the next arriving message or run `RunMatchReview` manually.

Integration tests must run only in the isolated `NinjaTest` PST, never in the production mailbox.

## 6. Add ribbon buttons

Open **File > Options > Customize Ribbon**:

1. Create a new tab or group, such as **Tickets**.
2. Under **Choose commands from**, select **Macros**.
3. Add `Project1.Ticketautomatismus.RunMatchReview`, rename it to a suitable caption such as **Review proposals**, and choose an icon.
4. Optionally add `RunMatchRulesReview`, `RunArchiveRuleWithProgress`, and `RunEmailRule`.

Only macros without parameters appear. `RunMatchDryRun` has an optional parameter, and `RunEmailRuleScript` is a rule action; neither appears in this list.

## 7. Create a report and review rules with Claude

`Ticketautomatismus.RunMatchRulesReview` first scans manually moved messages, writes `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\reports\match_review_<date>.md`, and copies an introductory prompt to the clipboard. Paste the prompt in a Claude session opened in the repository, or enter `/review-rules`.

The report includes counts, domains, ticket numbers, and subject stems of ticket-related messages. It does not copy separate sender names or message bodies, but subject stems can contain personal information. Review the report before sharing it.

## 8. Use learned rules

Rules are stored in `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\audit\match_rules.json`; the file does not exist until created. Reports propose JSON rule candidates with `"enabled":false`. Enable a reviewed rule explicitly by changing `enabled` to `true`.

Learned rules run last at `MATCH_LEVEL_LEARNED` (default 1, log only). A rule defines a search: domain plus subject stem identifies the single open ticket with that stem in its title. It does not bind messages to a fixed ticket number. The rules file is reloaded for every regular run and dry run.

## 9. Operational notes

- **Repeated audit entries:** `RunEmailRule` and the startup sweep recheck other messages from the last 14 days, producing audit rows on each run. The report deduplicates the same automatic decision for a message using the `msg` key; manual assignment detection records each message once. Set `MATCH_SWEEP_DAYS = 0` in `TicketConfig` to disable the sweep.
- **Manual assignments:** startup and report generation scan messages manually moved into ticket folders during the last 30 days. Each is recorded once as `manual_assignment`, with the hypothetical matcher result (`agree`, `differ`, `missed`, `excluded`, or `unchecked`).
- **Startup duration:** the sweep loads open NinjaOne tickets (approximately three pages in the reference environment) and checks ticket folders for new manual assignments. Duration depends on folder volume. Review `MANUAL_SCAN_DAYS`, `MATCH_SWEEP_DAYS`, and `INDEX_CACHE_MINUTES` if necessary.
- **Disable audit logging:** set `ENABLE_AUDIT_LOG = False`. **Disable all matching rules:** set `MATCH_LEVEL_MAX = 0`.
- **Signing:** importing invalidates the VBA project's signature. If required, sign it again using **Tools > Digital Signature** and your approved code-signing certificate.
- **Rollback:** close Outlook and restore `VbaProject.OTM.bak-<date>` over `VbaProject.OTM`.

## 10. Import the complete test suite into Outlook

Test modules must match the production source revision. Replacing only the three fakes from step 2 does not update the runner or other tests.

1. Back up the project as in step 0. Remove existing modules listed below from Project Explorer, exporting them first if needed. Keep production modules and `ThisOutlookSession`.
2. Import each of the following **39 files from `Tests\`** using **File > Import File** (`Ctrl+M`). Do not import `Tests\fixtures`: it contains JSON source data, already represented in `Fixtures.bas`.
3. Also import `outlook\OutlookTestGuard.bas`. Remove an existing copy first to avoid `OutlookTestGuard1`.
4. Install `ReviewForm` from step 2 or create it manually as described in step 12. `ReviewFormTests` instantiates the form; the runtime message-box fallback does not satisfy this test prerequisite.
5. Choose **Debug > Compile Project1**. Resolve all errors and duplicate modules such as `TestRunner1` or `FakeMail1` before running tests.

| File in `Tests\` | Module name in Project Explorer |
|---|---|
| `AuditReportTests.bas` | `AuditReportTests` |
| `AuditTests.bas` | `AuditTests` |
| `FakeAuditSink.cls` | `FakeAuditSink` |
| `FakeMail.cls` | `FakeMail` |
| `FakeMatchPrompt.cls` | `FakeMatchPrompt` |
| `FakeNinjaTransport.cls` | `FakeNinjaTransport` |
| `FakeProgressSink.cls` | `FakeProgressSink` |
| `FakeRejectionStore.cls` | `FakeRejectionStore` |
| `FakeTicketFolder.cls` | `FakeTicketFolder` |
| `FakeTicketIndex.cls` | `FakeTicketIndex` |
| `FakeTicketStore.cls` | `FakeTicketStore` |
| `FakeTokenFetcher.cls` | `FakeTokenFetcher` |
| `Fixtures.bas` | `Fixtures` |
| `IntegrationRunner.bas` | `IntegrationRunner` |
| `JsonLiteTests.bas` | `JsonLiteTests` |
| `LearnedRulesTests.bas` | `LearnedRulesTests` |
| `LiveTests.bas` | `LiveTests` |
| `LocalSettingsTests.bas` | `LocalSettingsTests` |
| `ManualScanTests.bas` | `ManualScanTests` |
| `MatchLevelsTests.bas` | `MatchLevelsTests` |
| `MatchPromptTextTests.bas` | `MatchPromptTextTests` |
| `MatchReviewTests.bas` | `MatchReviewTests` |
| `NinjaApiTests.bas` | `NinjaApiTests` |
| `NinjaAuthTests.bas` | `NinjaAuthTests` |
| `NinjaCredentialStoreTests.bas` | `NinjaCredentialStoreTests` |
| `NinjaTicketIndexTests.bas` | `NinjaTicketIndexTests` |
| `NinjaTicketLogicTests.bas` | `NinjaTicketLogicTests` |
| `NinjaTokenLogicTests.bas` | `NinjaTokenLogicTests` |
| `OutlookAdapterTests.bas` | `OutlookAdapterTests` |
| `ReviewFormTests.bas` | `ReviewFormTests` |
| `SelfTests.bas` | `SelfTests` |
| `TestAssert.bas` | `TestAssert` |
| `TestRunner.bas` | `TestRunner` |
| `TicketDecisionTests.bas` | `TicketDecisionTests` |
| `TicketMatcherTests.bas` | `TicketMatcherTests` |
| `TicketMatchServiceTests.bas` | `TicketMatchServiceTests` |
| `TicketParsingTests.bas` | `TicketParsingTests` |
| `TicketServiceTests.bas` | `TicketServiceTests` |
| `TimeUtilTests.bas` | `TimeUtilTests` |

**Self-test:** in `TestRunner`, place the cursor inside `RunAllTests` and press `F5`, or run `TestRunner.RunAllTests` in the Immediate window. Expect **303 passed tests, 0 failures** for this revision. Results appear in the Immediate window (`Ctrl+G`) and `%TEMP%\ninja_test_results.txt`. A result of only 117 tests indicates an old or incomplete runner. Fakes simulate failures deliberately; corresponding log messages belong to negative test cases.

**Separate Outlook integration tests:** run `IntegrationRunner.RunIntegrationTests` only against the local `NinjaTest` PST. There are **21 registered tests**, separate from the 303 unit tests. Results go to `%TEMP%\ninja_integration_results.txt`.

`TestRunner.RunLiveTests` is a separate opt-in command using a real NinjaOne connection. It is not part of the self-test.

## 11. Remaining acceptance work

- Reports become useful when there are enough explicit accept/reject decisions and manual assignments. Evaluate the installation's actual audit data before drawing conclusions.
- Headless logic tests and simulated form clicks in Excel do not establish actual Outlook dialog or adapter acceptance. Complete the outstanding checks in the [roadmap](ROADMAP.md) for the installed revision.

## 12. Create ReviewForm manually if import fails

1. In the VBA editor, choose **Insert > UserForm** and set `(Name)` to `ReviewForm`.
2. Create six controls with these exact names: `lblInfo` (Label), `lstProposals` (ListBox, `ListStyle = 1` / `fmListStyleOption`, `MultiSelect = 1` / `fmMultiSelectMulti`), and `btnSelectAll`, `btnAccept`, `btnReject`, `btnClose` (CommandButtons). Code sets their captions when the form opens.
3. Double-click the form to open its code window and replace its contents with `forms\ReviewForm.code.txt`.
4. Regenerate the exported form if needed using `tools\New-ReviewForm.ps1`.

For runtime use without a form, set `REVIEW_DIALOG = 1` in `TicketConfig`. Form tests still require `ReviewForm`.
