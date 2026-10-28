# Publication scope

This repository starts a new, independent Git history from a reviewed source snapshot. It contains no previous commit graph, tags, releases, PR discussions, installation configuration or private audit material. The existing MIT license and copyright notices are preserved.

## Included

- VBA sources, forms and synthetic fixtures/tests for classic Outlook ticket automation.
- A non-secret installation-settings template and configuration/import instructions.
- Initial host-independent C# core and reference tests for a planned .NET Framework 4.8 VSTO migration.

## Validation and limits

Snapshot verification on 2026-10-28:

- 303 VBA unit tests: 0 failures, 1,139 assertions, run in a disposable Excel workbook.
- 94 VBA source files pass ASCII/CRLF/header validation.
- C# runtime diagnostic: 9 test groups, 0 failures, 283 assertions.
- Synthetic fixtures and generated runners checked for consistency; local Markdown links and file encodings checked.
- No productive Outlook installation, live API test, VSTO host start, signing or installer test was performed for this snapshot. The C# runtime diagnostic is not a regular net48/VSTO build.

The 21 registered Outlook integration tests are a separate manual run against the local test PST. A source preview does not imply business acceptance, compatibility with every Office installation or a supported binary release.

## Future updates

Only reviewed source changes may enter this history. Do not merge or push a previous private development history, old tags, audit files, local configuration or generated build outputs into this repository. Gitignore does not protect against forced inclusion or unsafe archive creation.
