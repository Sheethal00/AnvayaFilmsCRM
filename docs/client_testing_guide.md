# Testing Guide — `AnvayaFilmsCRM.xlsx` (structure-only build)

This is the current build for reviewing the **data model and layout**, not
the finished system. It has all 22 sheets, the dropdown lists, and a
handful of live formulas — but no macros yet (see "What's not here yet"
below). Use it to react to structure, not to judge finished automation.

## 1. Open the file

Double-click `AnvayaFilmsCRM.xlsx`. It opens like any normal Excel file —
no security warning, nothing to "enable," because this version has no
macros in it.

## 2. Look at the sheet tabs along the bottom

They're grouped in four zones:
- **CRM & Sales** (01–04) — Dashboard, Leads, Quotations, Clients
- **Bookings & Production** (05–11) — Bookings, Wedding Calendar, Staffing,
  Equipment, Shoot Checklist, Production Pipeline, Album Tracker
- **Finance** (12–18) — Payments, Income, Expenses, Freelancers, Vendors,
  Accounting, Reports
- **Config** (19–20) — Settings, Help

Skim the tab names first — does the overall structure match how the
studio actually organizes work?

## 3. Try the dropdowns

- `02_CRM_Leads` → click any cell in the `Source` or `Status` column → a
  small dropdown arrow appears → click it to see the options.
- `10_Production_Pipeline` → `CurrentStage` — the 18-stage list
  (Inquiry → Archived) was reviewed and confirmed already (2026-07-25),
  so this is just for reference now, not something to re-litigate.

## 4. Check the sample data

- `02_CRM_Leads` has 5 sample leads already filled in — do the fields
  look right?
- `05_Bookings` has one sample booking, `BK-0001`. Look at the
  `BalanceAmount` and `DeliveryDeadline` columns — these are computed
  automatically from the advance paid and wedding date. Try changing
  `PackageValue` or `WeddingDate` on that row and watch `BalanceAmount`/
  `DeliveryDeadline` update live.

## 5. Try adding a new row

- `05_Bookings` has a filled-in sample row (`BK-0001`). Type a new row
  directly under it — `BalanceAmount`/`DeliveryDeadline` should
  auto-fill for that new row too, not just the sample one.
- `03_Quotation_Manager` and `06_Wedding_Calendar` don't have sample data
  — row 2 there only holds a formula with blank inputs either side of
  it (so it may show `0` or a formula error, that's expected). Type
  values into row 2 itself, or a new row under it, and confirm the
  formula column (`Subtotal`/`GST%`/`GSTAmount`/`Total` on Quotation
  Manager, `CountdownDays` on Wedding Calendar) fills in and recalculates
  as you'd expect.

## 6. Check the Dashboard tab

`01_Dashboard` now shows real numbers pulled live from the other sheets —
Total Leads, Conversion Rate, Revenue (this month/year), Outstanding
Payments, Upcoming Weddings, Employee Workload, Equipment Utilization.
With only the sample data loaded, most of these will be small or zero —
that's expected, not a bug. Try changing a lead's `Status` to `Won`, or a
payment's `Status` to `Cleared`, and confirm the relevant Dashboard number
updates (may need to press F9 or wait for a recalculation).

## 7. Note what feels wrong or missing

Especially:
- The dropdown lists — packages, expense categories, lead sources — the
  master copies live on the `19_Settings` tab
- Whether the sheet layout matches actual day-to-day workflow
- Whether the Dashboard KPIs (step 6) are the numbers that actually
  matter day-to-day, or if something's missing/unnecessary

## What's not here yet — don't test these, they're not bugs

This build has no VBA (macros), so none of the following works yet:
- New leads/bookings/payments do **not** get an ID auto-generated —
  you'd need to type one in by hand for now
- No double-booking warning when a team member is allocated twice on the
  same date
- No auto-conversion of a lead to a client when Status is set to Won
- No stage-history logging when `CurrentStage` changes
- Picking an `ItemDescription` on `03a_Quotation_LineItems` does **not**
  auto-suggest `UnitPrice` yet — the dropdown itself works (it's a
  formula-free feature), but the price fill-in is macro-driven
- On `09_Shoot_Checklist`, `CheckedBy` isn't restricted to that
  booking's crew yet, and ticking `Approved` doesn't record who did it
  or timestamp it — both macro-driven

These all exist as written VBA code already (`vba/modules/`,
`vba/classes/`) — they just aren't embedded in this particular file. See
`docs/decisions/0002-manual-vba-import.md` for how a macro-enabled
version gets built once the layout above is signed off.
