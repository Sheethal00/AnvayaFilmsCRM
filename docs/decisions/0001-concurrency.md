# 0001 — Concurrent Editing Model

**Status:** Decided (2026-07-23)

## Decision

The workbook is **single-user, local file** for all transactional sheets
(`02_CRM_Leads`, `03_Quotation_Manager`, `05_Bookings`, `07_Team_Allocation`,
`12_Payments`, `14_Expenses`, etc.). One person edits at a time; no
co-authoring, no OneDrive/SharePoint shared-file editing, no split
per-role workbooks.

The **only** sheet with more than one simultaneous editor is the team/employee
reference data (`19_Settings.TeamList`, referenced from `07_Team_Allocation`,
`10_Production_Pipeline`, `11_Album_Tracker`, `09_Shoot_Checklist` as a
lookup). Multiple people (e.g. a manager and an HR/admin person) may view or
update employee records.

## Why this resolves SPEC.md §0.1

Option (a) — single user, local file — applies workbook-wide. Options (b)
(SharePoint co-authoring) and (c) (split per-role workbooks + Power Query
sync) are **not needed**: they exist to handle simultaneous edits across
*different* transactional sheets (e.g. Sales entering leads while
Accountant records payments at the same time), and per the business owner
that scenario doesn't occur. The only real overlap is on `TeamList`, which
is low-frequency reference data, not a high-volume transactional log — a
"check before you edit" norm is sufficient, no macro/co-authoring tradeoff
needed.

## Consequences for schema/build

- No `LastEditedBy` / `LastEditedAt` / row-locking columns are needed on
  `02_CRM_Leads` or `12_Payments` — this was the specific risk called out
  in CLAUDE.md before scaffolding could start.
- VBA macros can run at full complexity on every sheet (no need to
  minimize heavy-macro sheets for co-authoring compatibility, since
  co-authoring is not in use).
- `19_Settings.TeamList` should still avoid destructive overwrites on
  save (e.g. append/update by `TeamMemberID` rather than blind row
  replace), since it's the one sheet where two people editing around the
  same time is expected. No schema change required — just a build/VBA
  note carried into `IDGenerator`/settings-write logic.
