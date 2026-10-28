---
description: Review the ticket matching rules and the user's decisions from the audit log report, with the goal of proposing new or changed rules
argument-hint: "[path of a match_review report]"
---

Start the review of the mail-to-ticket matching rules (M6 to M8 of this project) together with the user.

1. **Find the report.** Use `$ARGUMENTS` if it names a file. Otherwise take the newest `match_review_*.md` in `%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\reports` (the Outlook macro `RunMatchRulesReview` writes it). If there is none, say so and explain that the macro creates it; do not read the raw audit log files unless the user asks for it, because they contain more detail than the report.
2. **Read the current state:** `outlook/TicketConfig.bas` (levels `MATCH_LEVEL_*`, thresholds `REPORT_*`, `EXCLUDED_DOMAINS`, subject limits), `outlook/TicketMatcher.bas` and `outlook/MatchLevels.bas`, and `README.md` (current matching/audit/review behavior); consult `ROADMAP.md` for open acceptance tasks.
3. **Summarize briefly in the user's preferred chat language:** the period and volume, the table per rule (hit rate, level today, recommendation), the hot spots (domains with many rejections), the candidates for new rules including the check against later decisions, and what the mails without a match look like (domains only).
4. **Discuss with the user**, one topic at a time, and for every proposal give the evidence from the report:
   - level changes (raise to 3 only with the numbers the report names, lower or tighten when the hit rate is poor),
   - entries for the exclusion list,
   - new rules. A rule describes the **search**, never fixed ticket numbers: for example "domain D and subject stem S: the single open ticket with S in its title". It starts at level 1 (log only) and goes through the same protection as every rule (open ticket only, unambiguous, never moves in a dry run). Check a candidate against the report's later-decision check before proposing it, and say what would have gone wrong in the past.
5. **Only change things the user agrees to.** Never change levels, lists or rules on your own. When a change is agreed: implement it in the repo style (CLAUDE.md: English code, ASCII-only, tests first or together, `tools\Test-VbaFiles.ps1` and `tools\Invoke-VbaTests.ps1`), commit locally on a branch, **never push** before the user has tested it in Outlook.
6. **Privacy:** the report holds domains, ticket numbers and subject stems of mails with ticket relevance. Do not quote more of it than needed, never write person names or mail texts into code, tests, fixtures or commit messages, and use anonymized examples in tests.

Keep any repository documentation and pull request descriptions produced during the review in en-US. Preserve runtime strings and configuration identifiers.
