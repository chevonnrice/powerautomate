# Step 0: Create the SharePoint lists from CSV

You'll make five lists on your People Ops SharePoint site. Do them **in this order** (1 → 5). Each takes about 5 minutes.

## The one rule that matters

The CSV headers have **no spaces** (`EmployeeName`, not `Employee Name`). SharePoint turns a header into the column's
**internal name**, and every flow expression in this guide uses those internal names. Keep the headers exactly as they
are during import. You can change the **display name** afterwards (column settings → *Edit* → *Name*); the internal
name doesn't change when you do.

## How to import a CSV (same for each list)

1. Open the SharePoint site → **+ New → List → From CSV**, then upload the file.
2. In the preview, check each column's type against the tables below. The type dropdown is above each column.
3. Name the list exactly as shown below, then click **Create**.
4. Lists 4 and 5 contain one `SAMPLE` row, there only so SharePoint can detect the column types. Delete it after import.

> If SharePoint names the first column something other than **Title**, that's fine. The flows only use `Title` as a label.

---

## List 1: `HR Change Types` (csv/1-HR-Change-Types.csv)

Per-change-type settings (SLA, approval chain, notifications). Keep everything as text, except:

| Column | Type |
|---|---|
| `SLABusinessDays` | Number |
| `CloseLoopEmail`, `PostSummaryToIntake`, `NotifyPayrollOnResolve` | Yes/No |

After import, put real emails in `DefaultAssigneeEmail` (G — LOA is set to CJ; the rest are blank, meaning "nobody by default").

## List 2: `HR Change Templates` (csv/2-HR-Change-Templates.csv)

73 checklist steps (A1 through I), copied from your plan. Keep everything as text, except:

| Column | Type |
|---|---|
| `StepOrder` | Number |
| `Details` | Multiple lines of text |
| `BlockedUntilApproved`, `DueOnEffectiveDate`, `Active` | Yes/No |

> **Important:** the `ChangeType` values must match the Change Type choices on the Requests list character for
> character, including the long dash `—`. The CSV already has them right, so don't retype them.

## List 3: `HR Approval Matrix` (csv/3-HR-Approval-Matrix.csv)

Text everywhere, except `Active` = Yes/No. Replace the placeholder emails with your real Regional, HR and Payroll approvers.
`Region` = `*` means "any region". Add more rows for specific regions (e.g. `Role = Regional`, `Region = East`).

## List 4: `HR Change Requests` (csv/4-HR-Change-Requests.csv)

This is the system of record (replaces the Slack List). Set these types in the preview:

| Column (internal name) | Set type to | Choices / notes |
|---|---|---|
| `Title` | Single line of text | Set by the flow: [Change Type] — [Employee Name] |
| `RequestId` | Single line of text | Auto-generated: HRC-000123 |
| `EmployeeName` | Single line of text |  |
| `EmployeeId` | Single line of text |  |
| `ChangeType` | Choice | A1 — Offboarding (Voluntary), A2 — Offboarding (Involuntary), B — Job / Status Change, C — Comp Change, D — Pay / Banking Change, E — Personal Data Change, F — Time / PTO, G — LOA, H — Benefits / Life Event, I — Fractional FTE / Extra Shift Pickup, Other |
| `EffectiveDate` | Date (no time) |  |
| `Region` | Choice | East, Central, West |
| `Notes` | Multiple lines of text |  |
| `Status` | Choice | 0-New, 1-Assigned, 2-Waiting, 3-Resolved |
| `SLADueDate` | Date (no time) |  |
| `CompletionDate` | Date (no time) |  |
| `Outcome` | Choice | Completed, Rejected, Withdrawn |
| `JiraLink` | Single line of text | Offboarding only |
| `ApprovalStatus` | Choice | Not required, Pending, Approved, Rejected, Needs approver |
| `CurrentApprover` | Single line of text |  |
| `ApprovalHistory` | Multiple lines of text |  |
| `TasksTotal` | Number |  |
| `TasksDone` | Number |  |
| `ProgressPct` | Number |  |
| `LastDay` | Date (no time) |  |
| `SeparationReason` | Choice | Resignation, End of contract, Mutual separation, Termination, Reduction in force, End of contract (non-renewed), Performance-based |
| `EquipmentToRetrieve` | Multiple lines of text |  |
| `FinalPayType` | Single line of text |  |
| `NewTitle` | Single line of text |  |
| `NewDepartment` | Single line of text |  |
| `NewLocation` | Single line of text |  |
| `EmploymentStatus` | Choice | FT, PT, PRN |
| `NewRate` | Currency |  |
| `EffectivePayPeriod` | Single line of text |  |
| `PayChangeType` | Choice | Direct deposit, Tax, Garnishment |
| `VerificationAttached` | Yes/No |  |
| `FieldsChanging` | Single line of text |  |
| `NewValue` | Multiple lines of text |  |
| `TimeRequestType` | Choice | PTO, Sick, Unpaid time off, Bereavement, Jury duty, Other |
| `LeaveType` | Choice | FMLA, Medical, Parental, Military, Personal, Other |
| `StartDate` | Date (no time) |  |
| `EndDate` | Date (no time) |  |
| `Hours` | Number |  |
| `LifeEventType` | Choice | Marriage, Birth / adoption, Divorce, Loss of other coverage, Death of dependent, Other |
| `QualifyingDate` | Date (no time) |  |
| `DocumentationAttached` | Yes/No |  |
| `DatesWorked` | Single line of text |  |
| `ShiftType` | Choice | Training, Weekend clinic, Coverage, Other |
| `RateType` | Choice | Hourly, Flat / day |
| `RateAmount` | Currency |  |
| `HoursOrDays` | Number |  |
| `PayPeriod` | Single line of text | Set by the flow for change type I |
| `AfterCutoff` | Yes/No |  |

**After import:**

1. Delete the SAMPLE row.
2. **Choice columns**: SharePoint only adds the choice it saw in the sample row. Open each Choice column (column header →
   *Column settings → Edit*) and paste the full choice list from the table above, one per line. For `ChangeType`, also
   add `Other`. Set the defaults: `Status` = `0-New`, `ApprovalStatus` = `Not required`.
3. **Add three Person columns** by hand (*+ Add column → Person*). Type the name **without spaces** first, then rename the display name if you like:
   * `Manager` (required): manager approvals route here
   * `Assignee`: the case owner
   * `RequestedBy`: filled in by the flow
4. Make these required (*Column settings → Edit → Require that this column contains information*): `EmployeeName`,
   `EmployeeId`, `ChangeType`, `EffectiveDate`, `Manager`.
5. **Item-level permissions:** *Settings (gear) → List settings → Advanced settings*. Set *Read access* to "Read items that
   were created by the user" and *Create and Edit access* to "Create items and edit items that were created by the user".
   HR staff with Full Control still see everything.
6. Add the show/hide formulas so the form only shows the fields for the chosen change type: see
   [`../assets/form-conditional-formulas.md`](../assets/form-conditional-formulas.md).

## List 5: `HR Change Tasks` (csv/5-HR-Change-Tasks.csv)

| Column | Type |
|---|---|
| `StepOrder` | Number |
| `Details` | Multiple lines of text |
| `TaskStatus` | Choice: Not started, In progress, Done, Blocked - awaiting approval, Cancelled, N/A (default *Not started*) |
| `DueDate` | Date (no time) |
| everything else | Single line of text |

**After import:** delete the SAMPLE row, then add a **Lookup** column named `Request` (*+ Add column → Lookup*) →
source list **HR Change Requests** → show column **RequestId**.

## Optional: helpful views on HR Change Requests

* **Open by SLA:** filter `Status` is not equal to `3-Resolved`, sort by `SLADueDate` ascending
* **My cases:** same, plus `Assignee` is equal to `[Me]`
* **0-New (pick me up):** `Status` = `0-New`
* **Awaiting approval:** `ApprovalStatus` = `Pending`

✅ **Done when:** all five lists exist, and the New form on HR Change Requests shows the universal fields.
Next: [Step 1: Intake flow](01-intake-flow.md)
