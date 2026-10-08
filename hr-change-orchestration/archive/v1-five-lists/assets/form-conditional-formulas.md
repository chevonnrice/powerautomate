# HR Change Requests: form show/hide formulas

Your Slack form had **universal fields plus conditional fields per change type**. The SharePoint list form does the same thing with
[conditional formulas](https://learn.microsoft.com/sharepoint/dev/declarative-customization/list-form-conditional-show-hide).

**How:** open the list → **New** → **Edit form** (pencil, top right) → **Edit columns** → on each column below → **⋮ → Edit conditional formula** → paste.
Also drag the universal fields to the top, in the order of your Slack form.

## Universal fields (always shown)

Employee Name · Employee ID · Manager · Change Type · Effective Date · Region · Additional Notes

## Conditional fields

| Column | Shown for | Formula |
|---|---|---|
| Last Day (`LastDay`) | A1, A2 | `=if([$ChangeType] == 'A1 — Offboarding (Voluntary)' \|\| [$ChangeType] == 'A2 — Offboarding (Involuntary)', 'true', 'false')` |
| Reason for Separation (`SeparationReason`) | A1, A2 | `=if([$ChangeType] == 'A1 — Offboarding (Voluntary)' \|\| [$ChangeType] == 'A2 — Offboarding (Involuntary)', 'true', 'false')` |
| Equipment to Retrieve (`EquipmentToRetrieve`) | A1, A2 | `=if([$ChangeType] == 'A1 — Offboarding (Voluntary)' \|\| [$ChangeType] == 'A2 — Offboarding (Involuntary)', 'true', 'false')` |
| Final Pay Type (`FinalPayType`) | A1, A2 | `=if([$ChangeType] == 'A1 — Offboarding (Voluntary)' \|\| [$ChangeType] == 'A2 — Offboarding (Involuntary)', 'true', 'false')` |
| New Title (`NewTitle`) | B | `=if([$ChangeType] == 'B — Job / Status Change', 'true', 'false')` |
| New Department (`NewDepartment`) | B | `=if([$ChangeType] == 'B — Job / Status Change', 'true', 'false')` |
| New Location (`NewLocation`) | B | `=if([$ChangeType] == 'B — Job / Status Change', 'true', 'false')` |
| FT / PT / PRN (`EmploymentStatus`) | B | `=if([$ChangeType] == 'B — Job / Status Change', 'true', 'false')` |
| New Rate (`NewRate`) | C | `=if([$ChangeType] == 'C — Comp Change', 'true', 'false')` |
| Effective Pay Period (`EffectivePayPeriod`) | C | `=if([$ChangeType] == 'C — Comp Change', 'true', 'false')` |
| Pay / Banking Change Type (`PayChangeType`) | D | `=if([$ChangeType] == 'D — Pay / Banking Change', 'true', 'false')` |
| Verification Attached (`VerificationAttached`) | D | `=if([$ChangeType] == 'D — Pay / Banking Change', 'true', 'false')` |
| Field(s) Changing (`FieldsChanging`) | E | `=if([$ChangeType] == 'E — Personal Data Change', 'true', 'false')` |
| New Value (`NewValue`) | E | `=if([$ChangeType] == 'E — Personal Data Change', 'true', 'false')` |
| Time Request Type (`TimeRequestType`) | F | `=if([$ChangeType] == 'F — Time / PTO', 'true', 'false')` |
| Leave Type (`LeaveType`) | G | `=if([$ChangeType] == 'G — LOA', 'true', 'false')` |
| Start Date / First Date Worked (`StartDate`) | F, G, I | `=if([$ChangeType] == 'F — Time / PTO' \|\| [$ChangeType] == 'G — LOA' \|\| [$ChangeType] == 'I — Fractional FTE / Extra Shift Pickup', 'true', 'false')` |
| End Date / Expected Return (`EndDate`) | F, G | `=if([$ChangeType] == 'F — Time / PTO' \|\| [$ChangeType] == 'G — LOA', 'true', 'false')` |
| Hours (`Hours`) | F | `=if([$ChangeType] == 'F — Time / PTO', 'true', 'false')` |
| Life Event Type (`LifeEventType`) | H | `=if([$ChangeType] == 'H — Benefits / Life Event', 'true', 'false')` |
| Qualifying Date (`QualifyingDate`) | H | `=if([$ChangeType] == 'H — Benefits / Life Event', 'true', 'false')` |
| Documentation Attached (`DocumentationAttached`) | G, H | `=if([$ChangeType] == 'G — LOA' \|\| [$ChangeType] == 'H — Benefits / Life Event', 'true', 'false')` |
| Date(s) Worked (`DatesWorked`) | I | `=if([$ChangeType] == 'I — Fractional FTE / Extra Shift Pickup', 'true', 'false')` |
| Shift Type (`ShiftType`) | I | `=if([$ChangeType] == 'I — Fractional FTE / Extra Shift Pickup', 'true', 'false')` |
| Rate Type (`RateType`) | I | `=if([$ChangeType] == 'I — Fractional FTE / Extra Shift Pickup', 'true', 'false')` |
| Rate (hourly or day rate) (`RateAmount`) | I | `=if([$ChangeType] == 'I — Fractional FTE / Extra Shift Pickup', 'true', 'false')` |
| Hours or Days (`HoursOrDays`) | I | `=if([$ChangeType] == 'I — Fractional FTE / Extra Shift Pickup', 'true', 'false')` |

## Case fields (hidden on the New form, shown after submit)

Hide these from submitters on the New form. `RequestId` is filled in by the intake flow a few seconds after submission:

`=if([$RequestId] == '', 'false', 'true')`

| Column |
|---|
| Request (`Title`) |
| Request ID (`RequestId`) |
| Requested By (`RequestedBy`) |
| Assignee (`Assignee`) |
| Status (`Status`) |
| SLA Due Date (`SLADueDate`) |
| Completion Date (`CompletionDate`) |
| Outcome (`Outcome`) |
| Jira Ticket Link (`JiraLink`) |
| Approval Status (`ApprovalStatus`) |
| Current Approver (`CurrentApprover`) |
| Approval History (`ApprovalHistory`) |
| Tasks Total (`TasksTotal`) |
| Tasks Done (`TasksDone`) |
| Progress % (`ProgressPct`) |
| Pay Period (`PayPeriod`) |
| Submitted After Cutoff (`AfterCutoff`) |

> The formulas only change what the form shows. Submitters can still only *see their own* requests,
> because the provisioning script turns on item-level permissions (read and edit own items only).
