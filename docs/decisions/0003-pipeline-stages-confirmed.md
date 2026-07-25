# 0003 — Production Pipeline Stage List Confirmed

**Status:** Decided (2026-07-25)

## Decision

The client has reviewed `10_Production_Pipeline`'s 18-stage `CurrentStage`
dropdown (`schema/19_settings.yaml` → `lists.PipelineStageList`) and agreed
it matches their real workflow:

Inquiry → Consultation → Quoted → Booked → Pre-Production →
Shoot Scheduled → Shoot Completed → Footage Import → Culling/Selection →
Rough Cut → Editing → Color Grading → Client Review → Revisions →
Final Export → Delivery Prep → Delivered → Archived

## Why

This list was scaffolded as an unconfirmed placeholder (see the original
comment in `schema/19_settings.yaml` and step 3 of
`docs/client_testing_guide.md`) since SPEC.md §3 referenced "the 18-stage
list" without enumerating it. It was the single item flagged as most
worth the client's reaction during structure-only `.xlsx` testing. That
review is now done — no further sign-off needed on this list before it's
treated as final schema content.

## Consequences

- `schema/19_settings.yaml`'s `PipelineStageList` comment updated to
  reflect confirmed status instead of "do not treat this list as final."
- `docs/SPEC.md` §3 (`10_Production_Pipeline`) now points at this decision
  instead of leaving the list unenumerated.
- `docs/client_testing_guide.md` step 3 updated — future readers of that
  guide (e.g. testing a later rebuild) shouldn't be told to react to the
  pipeline stages as if they're still open.
- No schema/build change needed beyond the above — the dropdown's actual
  values were already correct in `19_Settings`, only their status was
  provisional.
