# Contributing

This source preview targets classic Outlook for Windows. Read [README.md](README.md) for current behavior and [ROADMAP.md](ROADMAP.md) for open work.

- Keep credentials and installation data outside the repository. Use synthetic fixtures and reserved example domains; do not paste real mail subjects or customer references into tests or issues.
- Preserve MIT and required third-party copyright/license notices.
- Keep VBA source ASCII with CRLF. Generate `Fixtures.bas`, `TestRunner.bas` and `IntegrationRunner.bas` through the repository scripts.
- Run `tools/Test-VbaFiles.ps1` and the relevant unit tests. Office folder/mail tests are limited to the isolated `NinjaTest` PST. Describe skipped checks explicitly.
- For the C# core, use `tools/Test-DotNetCore.ps1`; `-RuntimeOnly` is a limited diagnostic, not a VSTO build.
- Document completed changes in `CHANGELOG.md`, current behavior in `README.md`, and future work in `ROADMAP.md`.
- This publication history receives reviewed source snapshots. Never merge an older internal repository history or import its branches/tags. Submit changes against this repository's own history.
