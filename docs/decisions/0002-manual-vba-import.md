# 0002 — Manual VBA Import (No Python Required)

**Status:** Decided (2026-07-23)

## Decision

The client (or anyone with Excel but no Python) gets a working `.xlsm` by
importing VBA **through Excel's built-in VBA editor**, not by running
`scripts/build_workbook.py` themselves. That script's COM-automation path
(`try_import_vba`) is only needed for *repeatable, scripted* builds on a
machine that also has Python + pywin32 installed — it was never a
requirement for getting a single working file into a non-technical
client's hands.

## Why

`build_workbook.py` always produces `dist/AnvayaFilmsCRM.xlsx` — a
structure-only file with no VBA — regardless of what platform runs it.
The client only needs that one `.xlsx` file plus Excel's own **Alt+F11**
VBA editor, which supports importing `.bas` modules and pasting code into
existing class modules natively. No terminal, no `pip install`, no
pywin32.

## How to do it (Windows + Excel, ~10 minutes, one time per rebuild)

1. Get `dist/AnvayaFilmsCRM.xlsx` onto the Windows machine (email, USB,
   shared drive — however files normally move there). This file is built
   on any machine (Linux/Mac/Windows, no Excel needed) by running
   `python scripts/build_workbook.py --out dist/AnvayaFilmsCRM.xlsm`
   — it always writes the `.xlsx` first regardless of the `--out` name.
2. Open it in Excel. Press **Alt+F11** to open the VBA editor.
3. **Import the standalone modules** (these become new, independent
   modules — use File → Import File for each):
   `vba/modules/IDGenerator.bas`, `LeadConversion.bas`,
   `PipelineHistory.bas`, `AvailabilityCheck.bas`, `PaymentReminder.bas`,
   `RefreshAllPivots.bas`, `ItemCatalog.bas`.
   In the VBA editor: right-click **VBAProject (AnvayaFilmsCRM.xlsx)** →
   **Import File...** → pick the `.bas` file. Repeat for all seven.
4. **Paste code into `ThisWorkbook`** (do NOT import this one — it must
   go into the workbook's existing built-in module, not a new
   component): in the Project Explorer (left pane), double-click
   **ThisWorkbook** under "Microsoft Excel Objects". Open
   `vba/classes/ThisWorkbook.cls` in a text editor, copy everything
   *below* the `Option Explicit` line (skip the `VERSION`/`BEGIN`/`END`/
   `Attribute` header lines at the top — those are metadata for creating
   a brand-new module, not for this workbook's own), and paste it into
   the empty code pane for ThisWorkbook.
5. **Paste code into each sheet's module** — same "paste, don't import"
   rule, because each sheet already has its own built-in code module
   with an Excel-assigned name (e.g. "Sheet7") distinct from its tab
   name. In the Project Explorer, double-click the sheet under
   "Microsoft Excel Objects" that matches each file below, and paste
   that file's full contents (no header to strip — these files have none):

   | File | Sheet to open in Project Explorer |
   |---|---|
   | `vba/classes/02_CRM_Leads.cls` | 02_CRM_Leads |
   | `vba/classes/03_Quotation_Manager.cls` | 03_Quotation_Manager |
   | `vba/classes/03a_Quotation_LineItems.cls` | 03a_Quotation_LineItems |
   | `vba/classes/04_Client_Master.cls` | 04_Client_Master |
   | `vba/classes/05_Bookings.cls` | 05_Bookings |
   | `vba/classes/07_Team_Allocation.cls` | 07_Team_Allocation |
   | `vba/classes/08_Equipment.cls` | 08_Equipment |
   | `vba/classes/09_Shoot_Checklist.cls` | 09_Shoot_Checklist |
   | `vba/classes/10_Production_Pipeline.cls` | 10_Production_Pipeline |
   | `vba/classes/12_Payments.cls` | 12_Payments |
   | `vba/classes/14_Expenses.cls` | 14_Expenses |
   | `vba/classes/15_Freelancers.cls` | 15_Freelancers |
   | `vba/classes/16_Vendors.cls` | 16_Vendors |

6. Press **Ctrl+S** (or File → Save). Excel will prompt that the
   workbook can't be saved macro-free — choose **"No"** on that prompt
   so it offers **Excel Macro-Enabled Workbook (.xlsm)**, then save it
   as `AnvayaFilmsCRM.xlsm`. (Equivalently: File → Save As → set "Save
   as type" to *Excel Macro-Enabled Workbook (\*.xlsm)* before saving.)
7. Close and reopen the `.xlsm`. Excel will show a yellow security
   banner ("Macros have been disabled") — click **Enable Content**.
   `Workbook_Open` should then run (pivot refresh, payment reminder
   check, role-based tab hiding).

No Trust Center changes are needed for this manual route — "Trust access
to the VBA project object model" is only required when something
*programmatic* (like `build_workbook.py`'s COM path) edits the VBA
project. A human pasting into the editor by hand doesn't need it.

## Consequences

- This is a **manual, one-time-per-build** step. Every time `schema/` or
  `vba/` changes and `build_workbook.py` is rerun, the `.xlsx` is
  regenerated from scratch and this whole import has to be repeated by
  hand — there's no incremental "just update the changed macro" mode.
- For anything beyond one-off client testing (ongoing development,
  frequent rebuilds), get Python + pywin32 on *some* Windows machine
  (doesn't need to be the client's) and use `build_workbook.py`'s
  automated path instead — see `docs/decisions/0001-concurrency.md`
  build commands and CLAUDE.md. The client never needs to touch that
  machine or know it exists; they just receive the finished `.xlsm`.
