# Changelog

Completed changes and recorded checks belong here. Current behavior and installation are in [README.md](README.md); future work is in [ROADMAP.md](ROADMAP.md).

## 2026-10-28 - M9 matcher, infrastructure, and diagnostic VSTO host

- Created a regular Outlook VSTO project with the installed Visual Studio Community 2022 wizard, preserving generated host/designer code. Added a diagnostic Explorer Ribbon and core-reference callback. No mail processing or API access is connected to startup.
- Ported the five matching rules, normalization/extraction, exclusions, ambiguity/priority behavior, session-only snapshot index, matching-level action guards, fresh-ticket confirmation checks, and UTC/DST conversion.
- Added API response/closing-date/board parsing against shared synthetic VBA fixtures. Added local INI settings, read-only Credential Manager access, OAuth token caching, one fresh-token retry after 401, read-only ticket/log operations, and bounded board queries. Runtime integration and complete feature parity remain open.
- Default build/clean compiles the host without VSTO deployment-manifest, registration, certificate, or inclusion-list targets. Publishing requires explicitly configured development/deployment setup. Private signing material is ignored.
- Verified the regular net48 Release build of all five projects with Visual Studio Community 2022 17.14.37824.6. Core: 19 groups, 404 assertions; infrastructure: 7 groups, 69 assertions; all passed. The limited RuntimeOnly core diagnostic also passed (19 groups, 404 assertions).
- Tests used simulated HTTP handlers, shared synthetic fixtures, and a unique temporary UTF-16 INI. No real credential read, live API call, Office restart, add-in registration, trust change, or actual Outlook host/Ribbon test was performed. VBA sources and their operational installation remain unchanged.
- Documented current scope and remaining signing, host acceptance, Outlook adapters, proposal/review/audit persistence, deployment, and rollback work in the M9 guides and roadmap.
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
