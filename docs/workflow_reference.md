# Workflow Reference — What You Fill In, What Fills Itself In

The 15 working sheets, in the order the work actually happens (not sheet-
number order): a lead becomes a client, becomes a quote, becomes a
booking, gets shot, edited, delivered, and paid for.

Every automated field below is tagged one of two ways:

- **`[formula]`** — a live Excel formula. Works right now in the plain,
  macro-free `dist/AnvayaFilmsCRM.xlsx`.
- **`[macro]`** — VBA. Only runs once VBA is imported into a real
  `.xlsm` (see [`docs/decisions/0002-manual-vba-import.md`](decisions/0002-manual-vba-import.md)).
  Until that's done, these fields have to be typed in by hand — auto-IDs,
  the lead→client handoff, pipeline history logging, checklist
  timestamps, and the double-booking guards are all `[macro]`.

## 0. Settings & Team — `19_Settings`

One-time setup before day-to-day use.

- **You fill in:** `TeamMemberID, Name, Role, Phone, Email` (TeamList);
  adjust `GSTRate` / `DefaultDeliveryDays` or the dropdown lists if the
  studio's numbers differ.
- **Computed automatically:** each ID counter's `LastValue` `[macro]`.
- **Rule:** never hand-edit a counter's `LastValue` — every ID-generating
  sheet reads it live.

## 1. A lead comes in — `02_CRM_Leads`

Every booking starts here, before there's a client or a price.

- **You fill in:** `DateReceived, ClientName, Phone, Email, Source,
  EventType, EventDate (tentative), Budget, Status, LostReason,
  AssignedTo, NextFollowUpDate, Notes`.
- **Computed automatically:** `LeadID` `[macro]`; `ConvertedClientID` +
  a new row in `04_Client_Master` `[macro]`.
- **Trigger:** the client row and `ConvertedClientID` only appear the
  moment you set `Status` to `Won`.

## 2. Quote it — `03_Quotation_Manager` + `03a_Quotation_LineItems`

Line items priced separately (album, drone, highlight film…) instead of
one lump sum.

- **You fill in:** `LeadID, DateIssued, PackageSelected, ValidUntil,
  Status`; per line item: pick `ItemDescription` from the item catalog
  dropdown, then `Qty` and (if you want something other than the
  suggested price) `UnitPrice`.
- **Computed automatically:** `QuoteID` `[macro]`; `UnitPrice` —
  auto-suggested from `19_Settings`'s item catalog the moment you pick
  `ItemDescription`, but only while it's still blank — never overwrites
  a price you've already typed or edited `[macro]`; `LineTotal` = Qty ×
  UnitPrice `[formula]`; `Subtotal, GST%, GSTAmount, Total` `[formula]`.

## 3. Client record — `04_Client_Master`

Created for you the moment a lead is Won — you're only filling gaps.

- **You fill in:** `Address, FamilyContacts, ClientSince, Notes`.
- **Computed automatically:** `ClientID` + `PrimaryContactName, Phone,
  Email, Source` carried over from the lead `[macro]`;
  `TotalLifetimeValue` = sum of their bookings `[formula]`.

## 4. Confirm the booking — `05_Bookings`

Advance is in hand — the lead is now a real, dated commitment.

- **You fill in:** `ClientID, QuoteID, EventType, PackageValue,
  AdvanceAmount, AdvanceDate, WeddingDate, VenueName, VenueMapLink,
  Status`.
- **Computed automatically:** `BookingID` `[macro]`; `BalanceAmount` =
  Package − cleared payments `[formula]`; `DeliveryDeadline` = wedding
  date + Settings' delivery days `[formula]`.

## 5. Put it on the calendar — `06_Wedding_Calendar`

- **You fill in:** `BookingID, EventName, EventDate, StartTime,
  VenueName`.
- **Computed automatically:** `CountdownDays` `[formula]`.

## 6. Staff it — `07_Team_Allocation`

- **You fill in:** `BookingID, TeamMemberID, Role, Status`, and (for now)
  `EventDate` too.
- **Computed automatically:** `AllocationID` `[macro]`.
- **Guard, not a fill:** a macro checks for the same person on the same
  date across bookings and clears the cell if it finds a conflict,
  instead of writing a value.
- **Known gap:** `EventDate` is documented as auto-populating from the
  Wedding Calendar; today it doesn't. Type the wedding date in yourself.

## 7. Reserve equipment — `08_Equipment`

- **You fill in:** `Category, Model, SerialNumber, PurchaseDate, Status,
  CurrentBookingID, MaintenanceNotes`.
- **Computed automatically:** `EquipmentID` `[macro]`.
- **Guard, not a fill:** same conflict check as Team Allocation —
  overlapping bookings clear `CurrentBookingID` instead of allowing a
  double-assignment.

## 8. Pre/during/post-shoot checklist — `09_Shoot_Checklist`

Two-stage: the crew member checks an item off, someone else approves it
([`docs/decisions/0004-checklist-approval-workflow.md`](decisions/0004-checklist-approval-workflow.md)).

- **You fill in:** `BookingID, ChecklistPhase, ItemName`, `Checked`
  (tick it), `CheckedBy`, `Approved` (tick it once checked).
- **Computed automatically:** `CheckedAt` — stamped the instant `Checked`
  flips true `[macro]`; `ApprovedBy`/`ApprovedAt` — auto-recorded the
  instant `Approved` flips true `[macro]`.
- **Guard, not a fill:** `CheckedBy` must be crew actually allocated to
  that booking (`07_Team_Allocation`) — a macro clears it otherwise.
  `Approved` can't be ticked until `Checked` is true — a macro reverts
  it with a message instead.
- **Not real access control:** `ApprovedBy` records whoever's logged
  into Windows at the time, same as everywhere else in this workbook —
  it doesn't restrict who's *allowed* to tick `Approved`.

## 9. Edit & move through post-production — `10_Production_Pipeline` + `10a_Pipeline_History`

18 confirmed stages, Inquiry → Archived
([`docs/decisions/0003-pipeline-stages-confirmed.md`](decisions/0003-pipeline-stages-confirmed.md)).
You only ever touch the current-stage row.

- **You fill in:** `BookingID`, then move `CurrentStage` forward as work
  progresses; `AssignedEditor, Notes`.
- **Computed automatically:** `StageEnteredDate` — reset every time you
  change `CurrentStage` `[macro]`; `DaysInCurrentStage` `[formula]`; the
  entire `10a_Pipeline_History` table — logged on every stage change,
  never edited directly `[macro]`.

## 10. Album production — `11_Album_Tracker`

- **You fill in:** `BookingID, CurrentStage, PagesCount, DesignerID,
  PrintVendor, DispatchDate, TrackingNumber`.
- **Computed automatically:** nothing — fully manual sheet.

## 11. Record a payment — `12_Payments`

- **You fill in:** `BookingID, Amount, PaymentDate, Method, PaymentType,
  Status`.
- **Computed automatically:** `PaymentID` `[macro]`; `ReceiptNumber` —
  resets every calendar year `[macro]`.

## 12. Log an expense — `14_Expenses`

- **You fill in:** `Date, Category, Description, Amount,
  LinkedBookingID, PaidTo, PaymentMethod`.
- **Computed automatically:** `ExpenseID` — resets every calendar year
  `[macro]`.

## 13. Freelancers & Vendors — `15_Freelancers` + `16_Vendors`

Your contractor directory — pay them via `14_Expenses`, and their totals
catch up automatically.

- **You fill in:** `Name, Role/Category, Phone/ContactInfo,
  RatePerEvent, BankDetails`.
- **Computed automatically:** `FreelancerID` / `VendorID` `[macro]`;
  `TotalPaidYTD` — from matching Expenses rows `[formula]`.

## Never type into these

Pure outputs — every cell is a live formula reading the sheets above. If
a number here looks wrong, the fix is upstream, not here.

| Sheet | Contents |
|---|---|
| `01_Dashboard` | 8 KPI tiles |
| `13_Income` | Cleared payments by month |
| `17_Accounting` | Monthly P&L, cash flow |
| `18_Reports` | 7 report sections |
| `20_Help` | *(not built yet)* field defs, macro list, changelog |

## Before you rely on any of this

**Formulas work right now** in the plain `dist/AnvayaFilmsCRM.xlsx` —
balances, countdowns, GST, the Dashboard, Income, Accounting and Reports
sheets all recalculate live.

**Macros don't, yet.** Every auto-ID, the lead→client handoff, pipeline
history logging, checklist timestamps, and the double-booking guards
only run once VBA is imported into a macro-enabled `.xlsm` — a one-time
manual step per
[`docs/decisions/0002-manual-vba-import.md`](decisions/0002-manual-vba-import.md).
Until that's done, type IDs in by hand and expect the "Known gap" and
"Guard" notes above to be inactive too.

See also: [`docs/client_testing_guide.md`](client_testing_guide.md) for
how to hand this to a client for review.
