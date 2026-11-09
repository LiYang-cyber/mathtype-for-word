# Upstream contribution drafts

This directory holds the drafts prepared for
[felimet/mathtype-for-word](https://github.com/felimet/mathtype-for-word). They are working notes
for this fork, not part of the skill.

Status (2026-11-09): upstream shipped v1.4.3 on 2026-11-08. The report behind PR #8 became the
MTEF-based translation check in upstream PR #11 (credited as "Reported in #8"), and upstream also
fixed the existing-number-count bug found while reproducing issue #10. PR #8 and #9 stay open with
changes requested; this fork synced the upstream features in v2.1.0.

| File | Purpose |
|---|---|
| `01-issue-in-place-editing.md` | Issue proposing in-place editing of existing equations (the biggest gap) |
| `02-pr-docs-texvc-typography.md` | Pull request description: measured Texvc subset and Chinese inline-math rules |
| `03-pr-normalize-checker.md` | Pull request description: inline-math layout checker |
| `04-change-inventory.md` | Per-file inventory of how this fork differs from upstream |

The two documentation changes are already prepared as branches on top of upstream `940fc37`:
`docs/mathtype-texvc-subset` and `feat/inline-math-layout-check`.

Published upstream:

- Issue #10 — in-place editing of existing equations: <https://github.com/felimet/mathtype-for-word/issues/10>
- PR #8 — Texvc subset and Chinese inline typography: <https://github.com/felimet/mathtype-for-word/pull/8>
- PR #9 — inline-math layout checker: <https://github.com/felimet/mathtype-for-word/pull/9>

