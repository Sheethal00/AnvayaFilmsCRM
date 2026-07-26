# 0004 — Shoot Checklist: Check + Approve Workflow

**Status:** Decided (2026-07-26)

## Decision

`09_Shoot_Checklist` moves from a single `Checked` flag to a two-stage
workflow: the crew member on the booking checks an item off, then
someone else approves it. Three new columns —
`Approved, ApprovedBy, ApprovedAt` — mirror the existing
`Checked, CheckedBy, CheckedAt` pattern.

Two sub-decisions, both made directly with the business owner:

1. **`CheckedBy` is restricted to the booking's actual crew.** A macro
   (`vba/modules/ChecklistWorkflow.bas`) checks the named person against
   `07_Team_Allocation` for that `BookingID` and clears the cell (plus
   `Checked`/`CheckedAt`) if they're not on it — same revert-on-conflict
   pattern as the existing double-booking guard in
   `vba/modules/AvailabilityCheck.bas`.
2. **`ApprovedBy` auto-records the Windows username**
   (`Environ("Username")`) rather than being picked from a dropdown.
   There's no "Owner/Manager" role in `19_Settings.team_list`'s
   `TeamRoleList` (Photographer/Videographer/Drone Pilot/Assistant/
   Editor/Album Designer/Driver) to hang an approver dropdown off of,
   and adding one felt like solving a problem that didn't exist yet.

## Why

The business owner's stated requirement: "the pre check has to be done
by the member who is involved and it has to be approved by us." That's
two separate facts to enforce — *who* checked it, and that a second
party signed off — not just a single tick mark.

## Consequences

- `Approved` cannot be set `TRUE` until `Checked` is `TRUE` — the macro
  reverts it with a message box otherwise. Un-checking an item also
  revokes any existing approval, since "approved" implies "checked."
- **`ApprovedBy` is not real access control** — same caveat CLAUDE.md
  already states for `UserRolePermissions`. It records who was logged
  into Windows when the box was ticked; it does not restrict who is
  *allowed* to tick it. Anyone with the file and Enable Content can
  approve an item as if they were someone else by simply having that
  Windows session. Don't describe this to the client as securing
  sign-off — it's a paper trail, not a permission system.
- Like every other `[macro]`-tagged feature in this workbook, none of
  this runs until VBA is imported into a real `.xlsm`
  (`docs/decisions/0002-manual-vba-import.md`). The plain `.xlsx` still
  lets anyone type anything into `CheckedBy`/`Approved` with no
  enforcement at all.
