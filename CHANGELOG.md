# Changelog

Completed changes and recorded checks belong here. Current behavior and installation are in [README.md](README.md); future work is in [ROADMAP.md](ROADMAP.md).

## 2026-10-28 - Initial source preview

- Published a standalone source snapshot with a new history under the existing MIT license. No previous branches, tags, pull-request history or installation data are included.
- Includes VBA ticket sorting, archive decisions and offline fallback, ticket matching, per-rule safety levels, proposals, audit reports and review dialogs.
- Installation-specific values are read from a local non-secret INI file. Credentials remain in Windows Credential Manager. Fixtures and examples are synthetic.
- Includes the initial C#/.NET Framework 4.8 core for archive decisions, safety levels and parsing, plus reference tests. A VSTO host and installer are not implemented.
- Verification is documented in [PUBLICATION.md](PUBLICATION.md). Source publication is not a productive Outlook acceptance or a supported binary release.
