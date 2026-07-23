# CLAUDE.md

Guidance for Claude Code (or any agent) working in this repository.

## Project

Anvaya Films CRM + ERP — a wedding photography/videography business
management system, built as an Excel workbook (.xlsm) with VBA automation,
generated and maintained from source files in this repo rather than edited
by hand inside Excel directly. This repo is the source of truth; the .xlsm
in `dist/` is a build artifact.

Full functional/data spec: [`docs/SPEC.md`](docs/SPEC.md). Read it before
making schema or automation changes — it defines every sheet's columns,
relationships, and trigger logic. Don't invent fields or automation not
described there; propose a spec change first.

## Why generate the workbook from source instead of editing it directly

Editing .xlsm binaries directly is not diffable, not reviewable, and VBA
edited inside Excel has no version history. This repo keeps:
- sheet schemas as plain data (YAML) — reviewable in a PR
- VBA macros as `.bas`/`.cls` text files — reviewable, linted, testable
- a Python build script that assembles everything into the final `.xlsm`

## Repo Structure

```
/AnvayaFilmsCRM
│
├── CLAUDE.md                  # this file
├── README.md                  # human-facing quickstart
│
├── docs/
│   ├── SPEC.md                 # full schema + automation spec (source of truth)
│   └── decisions/               # one file per Open Decision from SPEC.md §0,
│       └── 0001-concurrency.md   # once decided, e.g. "single-user, local file"
│
├── schema/                     # machine-readable sheet definitions
│   ├── 02_crm_leads.yaml
│   ├── 03_quotation_manager.yaml
│   ├── 04_client_master.yaml
│   ├── 05_bookings.yaml
│   ├── ...                     # one file per sheet in docs/SPEC.md §3
│   └── settings.yaml            # dropdown lists, GST rate, ID counters
│
├── vba/                        # macros as plain text, imported at build time
│   ├── modules/
│   │   ├── IDGenerator.bas
│   │   ├── RefreshAllPivots.bas
│   │   ├── AvailabilityCheck.bas
│   │   └── PaymentReminder.bas
│   └── classes/
│       └── ThisWorkbook.cls
│
├── scripts/
│   ├── build_workbook.py       # reads schema/*.yaml + vba/*, writes dist/*.xlsm
│   ├── validate_schema.py      # lints schema files against docs/SPEC.md rules
│   └── seed_sample_data.py     # populates dist workbook with fake data for testing
│
├── sample_data/                # fixtures for local testing, not real client data
│   ├── leads.csv
│   ├── bookings.csv
│   └── payments.csv
│
├── tests/
│   ├── test_id_generation.py
│   ├── test_balance_formulas.py
│   └── test_schema_integrity.py   # e.g. every FK column has a matching PK sheet
│
├── dist/                       # build output, gitignored except last release
│   └── AnvayaFilmsCRM.xlsm
│
└── backup/                     # gitignored — local backups only, not committed
```

## Build Commands

```bash
# Build the workbook from schema + VBA source
python scripts/build_workbook.py --out dist/AnvayaFilmsCRM.xlsm

# Validate schema files against spec rules before building
python scripts/validate_schema.py

# Populate a built workbook with sample data for manual testing
python scripts/seed_sample_data.py dist/AnvayaFilmsCRM.xlsm

# Run tests
pytest tests/
```

## Working Conventions

- **Never hand-edit `dist/*.xlsm`.** Any change made by opening the file in
  Excel and typing will be lost on the next build. Change `schema/` or
  `vba/` instead, then rebuild.
- **Schema changes must match `docs/SPEC.md`.** If a change isn't reflected
  in the spec, update the spec in the same PR — it's the reference other
  contributors (and Claude, in future sessions) will read first.
- **IDs are never hardcoded or `=MAX()+1`.** Use the `IDGenerator` macro
  pattern described in SPEC.md §2/§6. Sequence counters live in
  `schema/settings.yaml`, not scattered across sheets.
- **New dropdown/reference lists go in `schema/settings.yaml`**, not inline
  in a sheet's own schema file — see SPEC.md §7.
- **Automation logic must have a defined trigger.** Before adding a macro,
  check SPEC.md §6 (Automation Trigger Table); add a row there if it's new.
- **Anything in SPEC.md §8 (Out of Scope) stays out of v1.** If asked to
  build one of those features, flag that it's a v2/roadmap item per the
  spec rather than building it into the current workbook.
- **Security/permissions caveat**: `UserRolePermissions` (SPEC.md §4) is
  UI-level tab hiding, not real access control. Don't describe it to users
  as securing financial data — say what it actually does (hides tabs by
  Windows username).

## Testing Philosophy

Since the deliverable is a binary Excel file, tests operate on the
intermediate representation, not the final .xlsm:
- `test_schema_integrity.py` checks every FK in `schema/*.yaml` resolves to
  a real PK in another schema file (catches typos like `BookngID`).
- `test_id_generation.py` runs the ID generator logic in isolation
  (extracted VBA logic reimplemented/tested in Python, or tested via a
  headless Excel COM call on Windows CI if available).
- `test_balance_formulas.py` checks formula strings written into cells
  match the expected formula for known inputs (e.g. `BalanceAmount`
  formula references the right range).

## Open Decisions Blocking Full Build

See `docs/SPEC.md` §0. Do not assume an answer and build around it silently —
check `docs/decisions/` first; if the relevant decision file doesn't exist yet,
stop and ask rather than guessing (this especially applies to the
multi-user/concurrency question in §0.1, which changes the schema for
`02_CRM_Leads` and `12_Payments` significantly).
