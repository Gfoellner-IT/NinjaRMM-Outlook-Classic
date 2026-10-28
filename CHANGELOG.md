# Changelog

Completed changes and recorded checks belong here. Current behavior and installation are in [README.md](README.md); future work is in [ROADMAP.md](ROADMAP.md).

## 2026-10-28 - Documentation in en-US

- Translated the roadmap, import guide, configuration guide, M9 architecture/reference cases, and agent instructions into en-US.
- Established en-US for repository documentation and descriptions while preserving runtime localization and the past/present/future documentation split.
- Corrected the import guide's stale unit-test count reference and included LocalSettings.bas in the README installation list.
- Checked local documentation links, import inventories, UTF-8 encoding, and the documentation-only diff. No application tests were rerun for this documentation change.

## 2026-10-28 - Initial source preview

- Published a standalone source snapshot with a new history under the existing MIT license. No previous branches, tags, pull-request history or installation data are included.
- Includes VBA ticket sorting, archive decisions and offline fallback, ticket matching, per-rule safety levels, proposals, audit reports and review dialogs.
- Installation-specific values are read from a local non-secret INI file. Credentials remain in Windows Credential Manager. Fixtures and examples are synthetic.
- Includes the initial C#/.NET Framework 4.8 core for archive decisions, safety levels and parsing, plus reference tests. A VSTO host and installer are not implemented.
- Verification is documented in [PUBLICATION.md](PUBLICATION.md). Source publication is not a production Outlook acceptance or a supported binary release.
