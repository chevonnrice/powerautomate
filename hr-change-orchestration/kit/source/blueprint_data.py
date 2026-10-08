"""Single source of truth for the HR Change + Orchestration build kit (v2: one SharePoint list).

Everything the generator writes (reference workbook, blueprint tracker, Word blueprint, list CSV)
comes from the structures in this file, so the documents never disagree with each other.
"""
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))

LIST_NAME = "HR Change Requests"
REF_FILE = "HR Change Reference.xlsx"
CR = "decodeUriComponent('%0D')"
LF = "decodeUriComponent('%0A')"

# ------------------------------------------------------------------ reference tables (flows read these)
SETTINGS = {
    "RequestPrefix": ("HRC-", "Prefix for Request IDs (HRC-000123)."),
    "TimeZone": ("Eastern Standard Time", "Windows time zone name used for dates, SLAs and the daily digest."),
    "CutoffDay1": (10, "Payroll cutoff day of month for the 1st–15th pay period."),
    "CutoffDay2": (25, "Payroll cutoff day of month for the 16th–end pay period (use 28 or lower)."),
    "ReminderBusinessDays": (2, "How many business days before each cutoff managers get the extra-shift reminder."),
    "IntakeFormUrl": ("https://<tenant>.sharepoint.com/sites/<site>/Lists/HRChangeRequests/NewForm.aspx", "The front door link (list New form). Pinned in #lane-people-talent."),
    "AdminEmail": ("cj@yourcompany.com", "Who receives automation failure alerts."),
}

CHANGE_TYPES = [
    # Code, ChangeType, SLA days, DueDateRule, SLATarget, ApprovalChain, DefaultAssigneeEmail, CloseLoopEmail, PostSummaryToPeopleOps, Sensitive, ConditionalFields, Notes
    ("A1", "A1 — Offboarding (Voluntary)", "", "Effective date", "Same day intake, executed by last day", "", "", "No", "Yes", "No",
     "Last day, reason for separation, equipment to retrieve, final pay type", "SLA = Last Day (or Effective Date)."),
    ("A2", "A2 — Offboarding (Involuntary)", "", "Effective date", "Same day intake, executed by last day", "", "", "No", "Yes", "Yes",
     "Last day, reason for separation, equipment to retrieve, final pay type", "Sensitive: employee name withheld from Slack and the digest."),
    ("B", "B — Job / Status Change", 3, "Business days", "2–3 business days", "", "", "Yes", "No", "No",
     "New title, new department, new location, FT/PT/PRN", "Planned chain (switch on at Sprint 3, after matrix sign-off): Manager > Regional"),
    ("C", "C — Comp Change", 5, "Business days", "3–5 business days", "", "", "Yes", "No", "No",
     "New rate, effective pay period", "Planned chain (switch on at Sprint 3): Manager > Regional > HR"),
    ("D", "D — Pay / Banking Change", 2, "Business days", "1–2 business days", "", "", "Yes", "No", "No",
     "Pay/banking change type (direct deposit, tax, garnishment), verification attached", "Manual verification (CJ or Johanna). No approval chain."),
    ("E", "E — Personal Data Change", 1, "Business days", "1 business day", "", "", "Yes", "No", "No",
     "Field(s) changing, new value", "Build and test this path first (Sprint 1)."),
    ("F", "F — Time / PTO", 2, "Business days", "1–2 business days", "", "", "Yes", "No", "No",
     "Request type, start/end dates, hours", "Planned chain (switch on at Sprint 3): Manager"),
    ("G", "G — LOA", "", "None", "Per policy", "", "cj@yourcompany.com", "No", "No", "Yes",
     "Leave type, start date, expected return, documentation attached", "Routes to CJ (case management). Stays open until return is confirmed. Sensitive."),
    ("H", "H — Benefits / Life Event", 3, "Business days", "2–3 business days", "", "", "Yes", "No", "No",
     "Event type, qualifying date, documentation attached", "Consider Sensitive = Yes if life-event details should stay out of Slack."),
    ("I", "I — Fractional FTE / Extra Shift Pickup", 2, "Business days", "Before the payroll cutoff", "HR", "", "No", "No", "No",
     "Date(s) worked, first date worked, shift type, rate type, rate, hours or days", "HR approval live from Sprint 2 (pilot). Payroll gets an FYI when resolved — verifier, not a gate."),
    ("Other", "Other", "", "None", "Needs manual scoping", "", "", "Yes", "No", "No", "—", "No checklist on purpose: the router posts 'needs manual scoping'."),
]
CHANGE_TYPE_COLS = ["Code", "ChangeType", "SLABusinessDays", "DueDateRule", "SLATarget", "ApprovalChain", "DefaultAssigneeEmail",
                    "CloseLoopEmail", "PostSummaryToPeopleOps", "Sensitive", "ConditionalFields", "Notes"]

with open(os.path.join(HERE, "checklists.json"), encoding="utf-8") as f:
    CHECKLISTS = json.load(f)
CHECKLIST_COLS = ["Code", "Step", "Task", "Details", "Owner", "Auto", "AfterApproval", "DueOnLastDay", "Active"]

APPROVERS = [
    ("Regional", "*", "regional.director@yourcompany.com", "Default Regional approver", "Yes", "Replace. Add one row per region (Region must match the list's Region choice)."),
    ("HR", "*", "cj@yourcompany.com", "HR approver", "Yes", "Used by C (stage 3) and I (pilot)."),
    ("Payroll", "*", "payroll@yourcompany.com", "Payroll approver", "No", "Only if Payroll ever becomes an approval stage (see decision D5). Inactive by default."),
]
APPROVER_COLS = ["Role", "Region", "ApproverEmail", "ApproverName", "Active", "Notes"]

# ------------------------------------------------------------------ the ONE list
CT_CHOICES = [c[1] for c in CHANGE_TYPES]
CODES_BY_CT = {c[1]: c[0] for c in CHANGE_TYPES}
FIELDS = [
    # internal, display, type, required, choices/default, group, showFor(codes), set by
    ("Title", "Request", "Single line of text", "No", "", "Universal", "", "Flow 1: '[Change Type] — [Employee Name]'"),
    ("RequestId", "Request ID", "Single line of text", "No", "", "Case", "", "Flow 1: HRC-000123"),
    ("EmployeeName", "Employee Name", "Single line of text", "Yes", "", "Universal", "", "Submitter"),
    ("EmployeeId", "Employee ID", "Single line of text", "Yes", "", "Universal", "", "Submitter"),
    ("Manager", "Manager", "Person", "Yes", "", "Universal", "", "Submitter (replaces 'Manager Name' so approvals can route)"),
    ("ChangeType", "Change Type", "Choice", "Yes", "; ".join(CT_CHOICES), "Universal", "", "Submitter"),
    ("EffectiveDate", "Effective Date", "Date only", "Yes", "", "Universal", "", "Submitter"),
    ("Region", "Region", "Choice", "No", "East; Central; West  (replace with yours)", "Universal", "", "Submitter (picks the Regional approver)"),
    ("Notes", "Additional Notes", "Multiple lines of text", "No", "", "Universal", "", "Submitter"),
    ("RequestedBy", "Requested By", "Person", "No", "", "Case", "", "Flow 1 (the submitter)"),
    ("Status", "Status", "Choice", "No", "0-New; 1-Assigned; 2-Waiting; 3-Resolved  (default 0-New)", "Case", "", "HR (approvals flow sets 2-Waiting / result)"),
    ("Assignee", "Assignee", "Person", "No", "", "Case", "", "HR (or default from Change Types)"),
    ("SLADueDate", "SLA Due Date", "Date only", "No", "", "Case", "", "Flow 1"),
    ("NextAction", "Next Action", "Single line of text", "No", "", "Case", "", "Flow 2 / Flow 6 (first open checklist step); HR can overwrite"),
    ("NextActionDue", "Next Action Due", "Date only", "No", "", "Case", "", "Flow 1 (= SLA); HR updates"),
    ("WaitingOn", "Waiting On", "Choice", "No", "Requester; Manager; Approver; Payroll; IT; Vendor; Other", "Case", "", "HR when Status = 2-Waiting (approvals flow sets 'Approver')"),
    ("Checklist", "Checklist", "Multiple lines of text (plain)", "No", "", "Case", "", "Flow 2 writes it; HR ticks [ ] → [x]"),
    ("ChecklistProgress", "Checklist %", "Number (0–100)", "No", "default 0", "Case", "", "Flow 6"),
    ("ApprovalStatus", "Approval Status", "Choice", "No", "Not required; Pending; Approved; Rejected; Needs approver  (default Not required)", "Case", "", "Flow 5"),
    ("CurrentApprover", "Current Approver", "Single line of text", "No", "", "Case", "", "Flow 5"),
    ("ApprovalHistory", "Approval History", "Multiple lines of text (plain)", "No", "", "Case", "", "Flow 5"),
    ("Outcome", "Outcome", "Choice", "No", "Completed; Rejected; Withdrawn", "Case", "", "Flow 4 (Completed) / Flow 5 (Rejected) / HR (Withdrawn)"),
    ("CompletionDate", "Completion Date", "Date only", "No", "", "Case", "", "Flow 4 (auto-stamp)"),
    ("JiraLink", "Jira Ticket Link", "Single line of text", "No", "", "Case", "", "HR (offboarding only; automated in Sprint 5)"),
    ("PayPeriod", "Pay Period", "Single line of text", "No", "", "Case", "", "Flow 1 (type I)"),
    ("AfterCutoff", "Submitted After Cutoff", "Yes/No", "No", "", "Case", "", "Flow 1 (type I)"),
    ("PayrollNotifiedOn", "Payroll Notified On", "Date only", "No", "", "Case", "", "Flow 3 (type I)"),
    ("AutomationNote", "Automation Note", "Single line of text", "No", "", "Case", "", "Catch blocks (exceptions)"),
    ("LastDay", "Last Day", "Date only", "No", "", "Conditional", "A1 A2", "Submitter"),
    ("SeparationReason", "Reason for Separation", "Choice", "No", "Resignation; End of contract; Mutual separation; Termination; Reduction in force; End of contract (non-renewed); Performance-based", "Conditional", "A1 A2", "Submitter"),
    ("EquipmentToRetrieve", "Equipment to Retrieve", "Multiple lines of text", "No", "", "Conditional", "A1 A2", "Submitter"),
    ("FinalPayType", "Final Pay Type", "Single line of text", "No", "", "Conditional", "A1 A2", "Submitter"),
    ("NewTitle", "New Title", "Single line of text", "No", "", "Conditional", "B", "Submitter"),
    ("NewDepartment", "New Department", "Single line of text", "No", "", "Conditional", "B", "Submitter"),
    ("NewLocation", "New Location", "Single line of text", "No", "", "Conditional", "B", "Submitter"),
    ("EmploymentStatus", "FT / PT / PRN", "Choice", "No", "FT; PT; PRN", "Conditional", "B", "Submitter"),
    ("NewRate", "New Rate", "Currency", "No", "", "Conditional", "C", "Submitter"),
    ("EffectivePayPeriod", "Effective Pay Period", "Single line of text", "No", "", "Conditional", "C", "Submitter"),
    ("PayChangeType", "Pay / Banking Change Type", "Choice", "No", "Direct deposit; Tax; Garnishment", "Conditional", "D", "Submitter"),
    ("VerificationAttached", "Verification Attached", "Yes/No", "No", "", "Conditional", "D", "Submitter"),
    ("FieldsChanging", "Field(s) Changing", "Single line of text", "No", "", "Conditional", "E", "Submitter"),
    ("NewValue", "New Value", "Multiple lines of text", "No", "", "Conditional", "E", "Submitter"),
    ("TimeRequestType", "Time Request Type", "Choice", "No", "PTO; Sick; Unpaid time off; Bereavement; Jury duty; Other", "Conditional", "F", "Submitter"),
    ("LeaveType", "Leave Type", "Choice", "No", "FMLA; Medical; Parental; Military; Personal; Other", "Conditional", "G", "Submitter"),
    ("StartDate", "Start Date / First Date Worked", "Date only", "No", "", "Conditional", "F G I", "Submitter"),
    ("EndDate", "End Date / Expected Return", "Date only", "No", "", "Conditional", "F G", "Submitter"),
    ("Hours", "Hours", "Number", "No", "", "Conditional", "F", "Submitter"),
    ("LifeEventType", "Life Event Type", "Choice", "No", "Marriage; Birth / adoption; Divorce; Loss of other coverage; Death of dependent; Other", "Conditional", "H", "Submitter"),
    ("QualifyingDate", "Qualifying Date", "Date only", "No", "", "Conditional", "H", "Submitter"),
    ("DocumentationAttached", "Documentation Attached", "Yes/No", "No", "", "Conditional", "G H", "Submitter"),
    ("DatesWorked", "Date(s) Worked", "Single line of text", "No", "", "Conditional", "I", "Submitter"),
    ("ShiftType", "Shift Type", "Choice", "No", "Training; Weekend clinic; Coverage; Other", "Conditional", "I", "Submitter"),
    ("RateType", "Rate Type", "Choice", "No", "Hourly; Flat / day", "Conditional", "I", "Submitter"),
    ("RateAmount", "Rate (hourly or day rate)", "Currency", "No", "", "Conditional", "I", "Submitter"),
    ("HoursOrDays", "Hours or Days", "Number", "No", "", "Conditional", "I", "Submitter"),
]
CODE_TO_CT = {c[0]: c[1] for c in CHANGE_TYPES}


def show_formula(field):
    if field[5] == "Conditional":
        conds = " || ".join(f"[$ChangeType] == '{CODE_TO_CT[c]}'" for c in field[6].split())
        return f"=if({conds}, 'true', 'false')"
    if field[5] == "Case":
        return "=if([$RequestId] == '', 'false', 'true')"
    if field[0] == "Title":
        return "=if([$RequestId] == '', 'false', 'true')"
    return "(always shown)"


CSV_SAMPLE = {
    "Title": "SAMPLE - delete after import", "RequestId": "HRC-000000", "EmployeeName": "Sample Employee", "EmployeeId": "000000",
    "ChangeType": CT_CHOICES[5], "EffectiveDate": "2026-10-15", "Region": "East", "Notes": "Sample row - delete",
    "Status": "0-New", "SLADueDate": "2026-10-16", "NextAction": "01. Confirm field(s) changing and new value", "NextActionDue": "2026-10-16",
    "WaitingOn": "Requester", "Checklist": "[ ] 01. Sample step", "ChecklistProgress": "0", "ApprovalStatus": "Not required",
    "CurrentApprover": "none", "ApprovalHistory": "none", "Outcome": "Completed", "CompletionDate": "2026-10-16", "JiraLink": "https://jira.example/ABC-1",
    "PayPeriod": "2026-10 (1-15)", "AfterCutoff": "No", "PayrollNotifiedOn": "2026-10-16", "AutomationNote": "none",
    "LastDay": "2026-10-15", "SeparationReason": "Resignation", "EquipmentToRetrieve": "Laptop", "FinalPayType": "Regular",
    "NewTitle": "Title", "NewDepartment": "Dept", "NewLocation": "Location", "EmploymentStatus": "FT", "NewRate": "25.00",
    "EffectivePayPeriod": "2026-10 (1-15)", "PayChangeType": "Direct deposit", "VerificationAttached": "No", "FieldsChanging": "Address",
    "NewValue": "12 Elm St", "TimeRequestType": "PTO", "LeaveType": "FMLA", "StartDate": "2026-10-15", "EndDate": "2026-10-16", "Hours": "8",
    "LifeEventType": "Marriage", "QualifyingDate": "2026-10-15", "DocumentationAttached": "No", "DatesWorked": "Oct 15",
    "ShiftType": "Training", "RateType": "Hourly", "RateAmount": "100.00", "HoursOrDays": "8",
}

STATUSES = [
    ("0-New", "Submitted, not yet picked up", "Column default at intake; approvals flow returns here if approved with no Assignee", "Pick it up: set Assignee, move to 1-Assigned"),
    ("1-Assigned", "Has an owner; work in progress", "HR, or the approvals flow after approval when an Assignee exists", "Work the Checklist; keep Next Action current"),
    ("2-Waiting", "Owner acted; waiting on someone", "HR (set Waiting On), or the approvals flow (Waiting On = Approver)", "Chase whoever 'Waiting On' names"),
    ("3-Resolved", "All tasks complete; case closed", "HR, or the approvals flow on rejection (Outcome = Rejected)", "Flow 4 stamps Completion Date + ✅ comment + close-the-loop"),
]

VIEWS = [
    ("New / Unassigned", "Status = 0-New", "Created ascending", "RequestId, Title, SLADueDate, Created", "Nobody owns these yet (Tag, You're It)."),
    ("My Active", "Assignee = [Me] AND Status ≠ 3-Resolved", "NextActionDue ascending", "RequestId, Title, Status, NextAction, NextActionDue, ChecklistProgress", "Your workbench."),
    ("Needs Attention Today", "Status ≠ 3-Resolved AND NextActionDue ≤ [Today]", "NextActionDue ascending", "RequestId, Title, Assignee, NextAction, NextActionDue", "Every active case has a next action and a date."),
    ("Waiting", "Status = 2-Waiting", "Group by WaitingOn", "RequestId, Title, Assignee, WaitingOn, CurrentApprover, Modified", "Who we're waiting on."),
    ("Awaiting Approval", "ApprovalStatus = Pending", "Created ascending", "RequestId, Title, CurrentApprover, ApprovalHistory", "Approvals in flight."),
    ("All Open by SLA", "Status ≠ 3-Resolved", "SLADueDate ascending; group by ChangeType", "RequestId, Title, Status, Assignee, SLADueDate, ChecklistProgress", "Dashboard-lite (Phase 3)."),
    ("Recently Resolved", "Status = 3-Resolved AND CompletionDate ≥ [Today]-14", "CompletionDate descending", "RequestId, Title, Outcome, CompletionDate, Assignee", "Review completed work."),
    ("Exceptions", "AutomationNote is not empty", "Modified descending", "RequestId, Title, AutomationNote, Modified", "Anything an automation could not finish."),
]

# ------------------------------------------------------------------ decisions, inspiration review, plan traceability
DECISIONS = [
    ("D1", "System of record", "Replace the HR Change Slack List with ONE SharePoint list: HR Change Requests.", "Your direction. One place per case; filterable views; works with Power Automate natively.", "Decided"),
    ("D2", "Reference data", "SLAs, approval chains, checklists, approvers and settings live as tables in one Excel workbook (HR Change Reference.xlsx) that the flows read.", "Manage them like a reference sheet, without extra lists or flow edits. Contract: never rename a table or a column header.", "Decided"),
    ("D3", "Front door", "The list's New form is the front door, pinned in #lane-people-talent. Microsoft Forms is reserved for the Phase 3 manager self-service form.", "A Slack Workflow Builder form cannot write to SharePoint without a bot token or webhook app. Principle update: 'Microsoft-native for the record, Slack for visibility'.", "Confirm (CJ)"),
    ("D4", "Checklist / subtasks", "The checklist is posted to the item's comment thread (the work log) AND written to an editable Checklist field. HR ticks [ ] → [x]; a flow keeps Checklist % and Next Action current.", "Delivers Phase 1 (thread checklist) and the Phase 2 goal (progress visibility) without a second list.", "Decided"),
    ("D5", "Template I — Payroll role", "Payroll is a verifier, not a gate: a one-way Slack FYI when People Ops resolves an I case.", "The plan's Flow 3 section says this explicitly; the master reference's 'Payroll final approval' line conflicts. To make Payroll a gate later, add 'Payroll' to I's ApprovalChain.", "Confirm (CJ)"),
    ("D6", "Manager field", "'Manager Name' becomes a Manager person field.", "Approvals need an email to route to.", "Confirm (Johanna)"),
    ("D7", "Comp 'approval confirmed (Y/N)'", "Retired. Replaced by the in-system approval chain.", "One source of truth for approval.", "Decided"),
    ("D8", "Approvals channel", "Power Automate Approvals (email, Teams, Power Automate app). In-Slack approve buttons wait for the Phase 3 bot token.", "Standard connector, no IT app approval.", "Decided"),
    ("D9", "Slack connection", "Power Automate Slack connector (OAuth) under a People Ops service account. Posts go to channels.", "Answers the plan's open question: no bot token needed. DMs need member IDs, so channels for now.", "Decided"),
    ("D10", "Status model", "Keep Tag, You're It (0-New, 1-Assigned, 2-Waiting, 3-Resolved). Add Waiting On, Next Action, Next Action Due, Outcome.", "Keeps continuity; adds 'who/what/when' for every active case (from the inspiration).", "Decided"),
    ("D11", "Sensitive change types", "A2 and G (and optionally H) withhold the employee name from Slack and the digest. Item-level permissions on the list.", "Limit sensitive content in notifications.", "Decided"),
    ("D12", "SLA calendar", "Business days skip weekends, not holidays.", "Simple first. Holiday table is on the Run-phase backlog.", "Accepted"),
    ("D13", "Approval rollout", "I's HR approval goes live in Sprint 2 (pilot). B, C, F chains are switched on in Sprint 3 by filling ApprovalChain in Excel after the matrix is signed off.", "Matches the dependency 'Approval matrix signed off by stakeholders'.", "Decided"),
    ("D14", "Plan typo", "'deparHRent' → 'department' (looks like a find/replace of 'tm' → 'HR'). Check the source doc for other damaged words.", "", "FYI"),
    ("D15", "Template I pilot cases", "The ICS live examples (DVA $100/hr, $800/day, first days Aug 15) have passed. Pilot with the next real I submission.", "", "FYI"),
]

INSPIRATION = [
    ("Email-first intake (employees keep emailing HR)", "Reject", "Breaks 'One front door' and 'Not in the intake, does not exist'.", "Adapt instead: an auto-reply on the HR mailbox that points people to the form (Sprint 2, no flow needed)."),
    ("Ticket-marker matching on email replies", "Reject (for now)", "Only needed with email intake.", "The earlier hr-ticketing-triage solution keeps this if email intake is ever wanted."),
    ("One accountable owner + Next action + Next action date on every active case", "Adopt", "Directly improves Tag, You're It: 'In progress' alone doesn't say what happens next.", "Next Action / Next Action Due fields; Flow 6 keeps Next Action = first open checklist step."),
    ("Richer status set (waiting on employee / third party)", "Adapt", "Keep the 0–3 labels for continuity.", "Add a Waiting On field used with 2-Waiting."),
    ("First human response date", "Defer", "Your SLAs measure completion, not first response.", "Revisit in Run phase for reporting."),
    ("Action-based workbench views", "Adopt", "Fast daily triage.", "8 views (see Views tab), incl. Needs Attention Today and Exceptions."),
    ("Daily digest instead of per-update emails", "Adopt", "Less noise; one place to see risk.", "Flow 7 (Sprint 4), sensitive names withheld."),
    ("Separate restricted lane for sensitive casework", "Adapt", "HR change requests are mostly routine; LOA/offboarding need care, ER investigations are out of scope.", "Sensitive flag per change type + item-level permissions now; restricted ER area is a Run-phase decision."),
    ("Copy email bodies/attachments into the ticket; Ticket Messages list", "Reject", "One-list rule; the comment thread is the log.", "Attachments go on the list item directly when needed."),
    ("Send HR responses from list fields", "Reject", "Complexity and risk for little gain.", "Reply in Slack/Outlook; close-the-loop email is automated at resolution."),
    ("Auto-close after inactivity", "Reject", "LOA must stay open; closure is a deliberate HR act in your plan.", "Daily digest surfaces stale items instead."),
    ("Exception path — never lose a request", "Adopt", "A failed step must stay visible.", "Try/Catch in every flow: admin email with run link + Automation Note + Exceptions view."),
    ("Loop guards with trigger conditions", "Adopt", "Flows that update the item they trigger on can loop.", "Flows 3, 4, 6 use trigger conditions / compare-before-write."),
    ("Don't promise resolution dates in auto-acknowledgements", "Adopt", "Sets the right expectation.", "Confirmation email says 'target completion' and 'a team member will pick it up'."),
    ("Pilot scenario test list (duplicates, out-of-office, failures…)", "Adapt", "Email-specific cases don't apply.", "22-case Test Pack written for HR change paths, approvals, permissions and failures."),
    ("Test access with a non-HR account before launch", "Adopt", "Filtered views are not security.", "Gate 0 / test T21."),
]

TRACE = [
    ("Guiding principle — One front door", "List New form pinned in #lane-people-talent; HR mailbox auto-reply points to it", "S0, S2"),
    ("Guiding principle — Not in the intake, does not exist", "Every case is an item; digest + views only show list items", "S2"),
    ("Guiding principle — Slack native first", "Updated (D3): Microsoft-native for the record, Slack for visibility", "S0"),
    ("Guiding principle — The list is the system of record", "ONE SharePoint list: HR Change Requests", "S0"),
    ("Guiding principle — Templates do the thinking", "Checklists table (A1–I) → Flow 2 Checklist Router", "S1"),
    ("Guiding principle — The thread is the work log", "Item Comments: intake, checklist, approvals, resolution all logged", "S1"),
    ("Guiding principle — Offboarding is a path", "A1/A2 are change types on the same list and engine", "S2"),
    ("Layer 1 — Intake form + confirmation to submitter", "List form + Flow 1 Intake Notification (confirmation email, People Ops ping)", "S1"),
    ("Layer 2 / Phase 1 — Thread-based checklist", "Flow 2 posts the checklist as an item comment", "S1"),
    ("Layer 2 / Phase 2 — True subtasks + progress", "Checklist field with [ ]/[x] + Flow 6 Checklist % and Next Action", "S3"),
    ("Master template items never completed", "Checklists table in the reference workbook (SOP source of truth)", "S0"),
    ("Layer 3 — Email sends", "Flow 1 confirmation, Flow 4 close-the-loop", "S1"),
    ("Layer 3 — Slack notifications", "Flows 1, 2, 3, 5, 6, 7", "S1–S4"),
    ("Layer 3 — Jira, Word templates, Monday.com", "Sprint 5, each behind its dependency gate", "S5"),
    ("Intake fields — universal + conditional", "List columns + show/hide formulas (List Fields tab)", "S0"),
    ("List fields — Request ID … Completion date", "List columns (List Fields tab)", "S0"),
    ("Status labels 0–3", "Status choice (Statuses tab)", "S0"),
    ("SLA targets by change type", "Change Types table → Flow 1 SLA Due Date (business days / last day / per policy)", "S1"),
    ("Completion auto-comment (✅ Resolved on [date] by [assignee] + stamp date)", "Flow 4 Completion Auto-Comment", "S1"),
    ("Flow 1 — Intake Notification", "Flow 1", "S1"),
    ("Flow 2 — Checklist Router (10 paths + Other catch-all)", "Flow 2 (one generic path reads the table; Other → 'needs manual scoping')", "S1"),
    ("Flow 3 — Payroll Notification (Template I only, not a gate)", "Flow 3", "S2"),
    ("Template I — HR approval → Paycor earning → resolve", "Flow 5 (HR stage) + checklist I + Flow 3", "S2"),
    ("Template I — log against the right pay period; flag after cutoff", "Flow 1 Pay Period + Submitted After Cutoff", "S2"),
    ("Template G — route to CJ, stays open until return", "Default Assignee = CJ; Flow 6 never auto-resolves; digest flags returns", "S2, S4"),
    ("Approval routing (B, C, F per approval matrix)", "Flow 5 + Approvers table; switched on after matrix sign-off", "S3"),
    ("Nice to add — cutoff reminder to managers", "Flow 7 cutoff reminder", "S4"),
    ("Phase 3 — Dashboard view", "Views (All Open by SLA, Recently Resolved) + daily digest; Power BI later", "S4, S6"),
    ("Phase 3 — Manager self-service form", "Microsoft Forms → same list (scoped types)", "S6"),
    ("Phase 3 — Bot token / in-Slack approvals; Paycor iPaaS", "Run-phase backlog", "S6"),
    ("Build order — E first, then I, then G", "Sprint 1 = E end-to-end; Sprint 2 = I pilot then G", "S1, S2"),
    ("Open question — Slack app vs Power Automate connector", "Answered: connector (D9)", "S0"),
]

DEPENDENCIES = [
    ("HR Change Reference.xlsx stored on the People Ops site", "CJ", "Ready (this kit)", "S0", "Flows read it."),
    ("Slack connection under a People Ops service account", "CJ / IT", "To confirm", "S0", "Messages appear from this account."),
    ("Field list signed off", "Johanna", "To confirm", "S0", "Gate 0."),
    ("Approval matrix signed off by stakeholders", "CJ", "To confirm", "S3", "Gate 3: switches on B, C, F approvals."),
    ("Jira connector access in Power Automate + HR Jira project", "IT", "To confirm", "S5", "Plan open question."),
    ("Shared offboarding email distro exists and is current", "CJ / Johanna", "To confirm", "S5", "Plan open question."),
    ("Word termination letter template (content controls)", "CJ", "To build", "S5", "'Populate a Microsoft Word template' is a premium connector action."),
    ("Monday backfill request form URL", "Talent", "To confirm", "S5", ""),
    ("Payroll Monday board — column/item to post to", "Sydney / Sandra", "To confirm", "S5", "Plan open question."),
    ("Paycor bot token / iPaaS", "IT / Paycor CSM", "Pending", "S6", "Run phase."),
]

# ------------------------------------------------------------------ roadmap
SPRINTS = [
    ("S0", "Foundations", 1, "Crawl", "Reference workbook, the one list, permissions, form, views, connections.",
     "Plan approved; this kit downloaded.",
     "G0 — Ready to build", ["D3, D5, D6 confirmed by CJ", "Field list signed off with Johanna", "Reference workbook uploaded and Settings filled",
                             "List built: all columns, choices, required fields, show/hide formulas, views", "T21 passed (non-HR account sees only own requests)"]),
    ("S1", "Crawl — Path E end to end", 2, "Crawl", "Flow 1 Intake, Flow 2 Checklist Router, Flow 4 Completion Auto-Comment. Proven on E.",
     "G0 passed.",
     "G1 — Plumbing proven", ["T01–T06 pass", "One live E submission by Johanna closed end to end", "A failure alert email received from a Catch block (any test)", "Flows owned by the HR/service account, not a personal one"]),
    ("S2", "Crawl complete — Template I pilot + front door live", 2, "Crawl", "Flow 5 Approval Routing (HR stage for I), Flow 3 Payroll Notification, all checklists live, front door launched.",
     "G1 passed.",
     "G2 — Front door live", ["T07–T13 pass", "First real I case approved, earning entered, resolved, Payroll FYI posted", "Form pinned in #lane-people-talent; manager comms sent", "HR mailbox auto-reply points to the form"]),
    ("S3", "Walk — Approvals B/C/F + checklist progress", 2, "Walk", "Approval chains switched on from Excel; Flow 6 Checklist Progress; UAT with Johanna.",
     "G2 passed; approval matrix signed off.",
     "G3 — Approvals live", ["Approval matrix signed off; Approvers table filled per region", "T14–T17 and T22 pass", "UAT with Johanna on every path A1–I"]),
    ("S4", "Walk — Daily digest, cutoff reminder, hardening", 2, "Walk", "Flow 7; exception drill; two weeks of steady operation.",
     "G3 passed.",
     "G4 — Operational", ["T18–T20 pass", "Two weeks of digests reviewed", "No request handled outside the list in those two weeks", "SLA baseline reviewed (All Open by SLA / Recently Resolved)"]),
    ("S5", "Walk — Downstream actions (dependency-driven)", 2, "Walk", "Jira, offboarding distro, Word letters, Monday board — each only when its dependency is confirmed.",
     "G4 passed; each item's dependency confirmed.",
     "G5 — per item", ["Dependency confirmed in writing", "Built and tested on a test request", "Checklist step updated to show it is automated (Auto column)"]),
    ("S6", "Run — Self-service + full audit", 0, "Run", "Manager self-service form, dashboard, in-Slack approvals, Paycor iPaaS, holiday-aware SLA.",
     "G5 items as available.",
     "Ongoing", ["Each item ships with its own test cases and a rollback note"]),
]

# ------------------------------------------------------------------ build steps
# Each step: id, sprint, area, title, where, do (list), paste (list of (label, text)), check, owner
STEPS = []


def step(id, area, title, where="", do=(), paste=(), check="", owner="Johanna"):
    STEPS.append(dict(id=id, sprint=id.split("-")[0], area=area, title=title, where=where, do=list(do), paste=list(paste), check=check, owner=owner))


# ---- reusable blocks (documented once, referenced by id)
OPENING = [
    ("Settings row  (Excel Online (Business) → List rows present in a table)", f"Location: your People Ops site · Document Library: Documents · File: HR Change/{REF_FILE} · Table: tblSettings"),
    ("Settings  (Compose)", "first(outputs('Settings_row')?['body/value'])"),
    ("Code  (Compose)", "first(split(triggerBody()?['ChangeType']?['Value'], ' '))"),
    ("Change type row  (Excel → List rows present in a table) — Table: tblChangeTypes · Filter Query:", "Code eq '@{outputs('Code')}'"),
    ("Config  (Compose)", "first(outputs('Change_type_row')?['body/value'])"),
    ("Request ID  (Compose)", "concat(outputs('Settings')?['RequestPrefix'], formatNumber(triggerBody()?['ID'], '000000'))"),
]
COMMENT = [
    ("<Name> text  (Compose) — type the message; insert @{…} parts with fx", "e.g. 📥 @{outputs('Request_ID')} received."),
    ("Post <name>  (SharePoint → Send an HTTP request to SharePoint) — Method", "POST"),
    ("Uri", f"_api/web/lists/getbytitle('{LIST_NAME}')/items(@{{triggerBody()?['ID']}})/Comments"),
    ("Headers", "Accept = application/json;odata=nometadata   ·   Content-Type = application/json;odata=nometadata"),
    ("Body (fx)", "setProperty(json('{}'), 'text', take(outputs('<Name>_text'), 1900))"),
]
CATCH = [
    ("Catch  (Control → Scope) → ⋯ → Configure run after", "Try: ☑ has failed  ☑ has timed out  (☐ is successful)"),
    ("Failure alert  (Office 365 Outlook → Send an email (V2)) — To", "your admin email (type it — don't depend on Settings here)"),
    ("Subject", "[HR Change] @{workflow()?['tags']?['flowDisplayName']} failed"),
    ("Body", "Open the failed run: @{concat('https://make.powerautomate.com/environments/', workflow()?['tags']?['environmentName'], '/flows/', workflow()?['name'], '/runs/', workflow()?['run']?['name'])}"),
    ("Flag the item  (SharePoint → Update item, item-triggered flows only) — AutomationNote (fx)", "concat('⚠️ ', workflow()?['tags']?['flowDisplayName'], ' failed ', formatDateTime(utcNow(), 'yyyy-MM-dd HH:mm'), ' UTC — see admin email')"),
]
BLOCKS = {"OPENING": ("Standard opening block", OPENING), "COMMENT": ("Comment pattern (post to the item's thread)", COMMENT), "CATCH": ("Try / Catch pattern", CATCH)}

# ================= S0
step("S0-01", "Decisions", "Confirm the open design decisions", "Decisions tab",
     ["Review D3 (front door), D5 (Template I: Payroll verifier, not gate), D6 (Manager person field).", "Mark each Confirmed / Changed in the Decisions tab."],
     check="D3, D5, D6 no longer say 'Confirm'.", owner="CJ")
step("S0-02", "Excel", "Store the reference workbook", "People Ops SharePoint site → Documents",
     [f"Create a folder 'HR Change' and upload {REF_FILE}.", "Folder → Manage access: HR owners can edit; everyone else no access.",
      "Read the About sheet: never rename a table (tblSettings, tblChangeTypes, tblChecklist, tblApprovers) or a column header."],
     check="The file opens in Excel for the web and Table Design shows the four table names.", owner="CJ")
step("S0-03", "Excel", "Fill in Settings", f"{REF_FILE} → Settings",
     ["Set TimeZone, CutoffDay1, CutoffDay2 (your real payroll cutoffs), ReminderBusinessDays and AdminEmail.", "Leave IntakeFormUrl for S0-14."],
     check="No '<' placeholders left except IntakeFormUrl.")
step("S0-04", "Excel", "Review Change Types", f"{REF_FILE} → Change Types",
     ["Check SLA days and the Due Date Rule for every type (Business days / Effective date / None).",
      "Set DefaultAssigneeEmail where a type always goes to the same person (G → CJ is pre-filled).",
      "Leave ApprovalChain blank for B, C, F until Sprint 3 (decision D13). I = 'HR' (pilot).",
      "Decide Sensitive for H (A2 and G are Yes)."],
     check="Every row has Code, ChangeType, DueDateRule; Yes/No columns contain only Yes or No.")
step("S0-05", "Excel", "Review the checklists against the SOP", f"{REF_FILE} → Checklists",
     ["Walk A1–I with Johanna. Edit Task / Details / Owner as needed. Keep Step as two digits (01, 02 …).",
      "Auto: 'approval' = ticked by the approvals flow; 'resolution' = ticked when the case is resolved. AfterApproval = Yes shows 🔒."],
     check="Johanna signs off each change type's steps.")
step("S0-06", "Excel", "Draft the Approvers table", f"{REF_FILE} → Approvers",
     ["One row per Role (Regional, HR, Payroll) and Region. Region '*' = default for any region.", "Region values must match the list's Region choices exactly."],
     check="An HR row with Region '*' and Active = Yes exists (needed for the I pilot).", owner="CJ")
step("S0-07", "SharePoint", "Create the list from CSV", "Site → + New → List → From CSV → HR-Change-Requests.csv",
     ["Keep the headers exactly (no spaces = internal names the flows use).", "In the preview, set each column type per the List Fields tab.",
      f"Name the list '{LIST_NAME}' → Create. Delete the SAMPLE row."],
     check="List settings → Columns shows internal names without spaces (hover a column: Field=EmployeeName).")
step("S0-08", "SharePoint", "Finish the choice columns", "List → column header → Column settings → Edit",
     ["For every Choice column, paste the full choice list from the List Fields tab (one per line).", "Defaults: Status = 0-New, ApprovalStatus = Not required."],
     check="New form shows all 11 Change Type values.")
step("S0-09", "SharePoint", "Add the three Person columns", "List → + Add column → Person",
     ["Manager (required), Assignee, RequestedBy — type the name without spaces first; rename the display name afterwards if you like."],
     check="All three appear in the New form / list settings.")
step("S0-10", "SharePoint", "Required fields + permissions", "List settings",
     ["Required: EmployeeName, EmployeeId, Manager, ChangeType, EffectiveDate.",
      "Advanced settings → Read access: 'Read items that were created by the user'; Create and Edit access: 'Create items and edit items that were created by the user'.",
      "HR team: Full Control (or Edit + Override List Behaviors) so they see everything."],
     check="Test T21 with a non-HR colleague.", owner="CJ")
step("S0-11", "SharePoint", "Show/hide formulas on the form", "List → + New → Edit form (pencil) → Edit columns → column ⋯ → Edit conditional formula",
     ["Paste the 'Show formula' from the List Fields tab for every Conditional and Case column.",
      "Drag the universal fields to the top in Slack-form order: Employee Name, Employee ID, Manager, Change Type, Effective Date, Region, Additional Notes."],
     check="Choosing 'C — Comp Change' shows New Rate + Effective Pay Period only; case fields are hidden on New.")
step("S0-12", "SharePoint", "Create the workbench views", "List → All Items → Create new view",
     ["Create the 8 views on the Views tab (filters, sort/group, columns)."],
     check="'Needs Attention Today' and 'Exceptions' exist.")
step("S0-13", "Power Automate", "Create the connections", "make.powerautomate.com → Data → Connections",
     ["Slack: sign in as the People Ops service account (D9).", "SharePoint, Office 365 Outlook, Excel Online (Business), Approvals: the HR flow-owner account (Full Control on the site)."],
     check="Five connections show 'Connected'.", owner="CJ")
step("S0-14", "Excel", "Record the front door link", "List → + New",
     ["Copy the browser address of the New form into Settings → IntakeFormUrl."],
     check="The link opens the New form for a non-HR user.")

# ================= S1 — Flow 1
step("S1-01", "Flow 1", "Create 'HR Change - 1 Intake Notification'", "My flows → + New flow → Automated cloud flow",
     [f"Trigger: SharePoint — When an item is created · Site: People Ops · List: {LIST_NAME}.",
      "Add Control → Scope, rename it 'Try'. Build every following action INSIDE Try."],
     check="Saved with trigger + empty Try.")
step("S1-02", "Flow 1", "Add the standard opening block", "Inside Try",
     ["Add the six actions of the Standard opening block (Appendix A), renamed exactly."],
     paste=OPENING, check="Run once (Test → Manually, then create an E request): Config shows the E row.")
step("S1-03", "Flow 1", "SLA due date", "Inside Try, after Request ID",
     ["Four Compose actions, renamed as shown."],
     paste=[("Submitted Local", "convertFromUtc(triggerBody()?['Created'], outputs('Settings')?['TimeZone'])"),
            ("SLA Start  (weekend submissions start Monday)", "addDays(startOfDay(outputs('Submitted_Local')), if(equals(dayOfWeek(outputs('Submitted_Local')), 6), 2, if(equals(dayOfWeek(outputs('Submitted_Local')), 0), 1, 0)))"),
            ("SLA Days", "if(empty(outputs('Config')?['SLABusinessDays']), 0, int(outputs('Config')?['SLABusinessDays']))"),
            ("SLA Due Date", "if(equals(outputs('Config')?['DueDateRule'], 'Effective date'), formatDateTime(coalesce(triggerBody()?['LastDay'], triggerBody()?['EffectiveDate']), 'yyyy-MM-dd'), if(equals(outputs('SLA_Days'), 0), null, formatDateTime(addDays(outputs('SLA_Start'), add(add(mul(div(outputs('SLA_Days'), 5), 7), mod(outputs('SLA_Days'), 5)), if(greater(add(dayOfWeek(outputs('SLA_Start')), mod(outputs('SLA_Days'), 5)), 5), 2, 0))), 'yyyy-MM-dd')))")],
     check="E submitted Monday → SLA Due Date = Tuesday.")
step("S1-04", "Flow 1", "Pay period + cutoff flag (used by type I)", "Inside Try",
     ["Four Compose actions. Harmless for other change types."],
     paste=[("Pay Period Date", "coalesce(triggerBody()?['StartDate'], triggerBody()?['EffectiveDate'])"),
            ("Pay Period", "concat(formatDateTime(outputs('Pay_Period_Date'), 'yyyy-MM'), if(lessOrEquals(dayOfMonth(outputs('Pay_Period_Date')), 15), ' (1-15)', ' (16-EOM)'))"),
            ("Cutoff Date", "addDays(startOfMonth(outputs('Pay_Period_Date')), sub(int(if(lessOrEquals(dayOfMonth(outputs('Pay_Period_Date')), 15), outputs('Settings')?['CutoffDay1'], outputs('Settings')?['CutoffDay2'])), 1))"),
            ("After Cutoff", "and(equals(outputs('Code'), 'I'), greater(ticks(outputs('Submitted_Local')), ticks(addDays(outputs('Cutoff_Date'), 1))))")],
     check="For E: After Cutoff = false.")
step("S1-05", "Flow 1", "Write ID, title and SLA back to the item", "Inside Try → SharePoint → Update item",
     [f"List: {LIST_NAME} · Id: ID (trigger).", "Required columns the action asks for: pick the SAME field from the trigger (it writes back what is already there).",
      "Do not set Status or ApprovalStatus here — column defaults cover intake and Flow 5 owns approvals."],
     paste=[("Title (fx)", "concat(triggerBody()?['ChangeType']?['Value'], ' — ', triggerBody()?['EmployeeName'])"),
            ("RequestId (fx)", "outputs('Request_ID')"),
            ("RequestedBy Claims (fx)", "triggerBody()?['Author']?['Email']"),
            ("SLADueDate (fx)", "outputs('SLA_Due_Date')"),
            ("NextActionDue (fx)", "outputs('SLA_Due_Date')"),
            ("PayPeriod (fx)", "if(equals(outputs('Code'), 'I'), outputs('Pay_Period'), '')"),
            ("AfterCutoff (fx)", "outputs('After_Cutoff')")],
     check="Item shows Request = 'E — Personal Data Change — <name>' and Request ID HRC-00000N.")
step("S1-06", "Flow 1", "Default assignee (from Change Types)", "Inside Try → Control → Condition 'Default assignee?'",
     ["Left (fx) and(empty(triggerBody()?['Assignee']), not(empty(coalesce(outputs('Config')?['DefaultAssigneeEmail'], '')))) · is equal to · (fx) true",
      "True → Update item 'Set default assignee' (required columns from trigger)."],
     paste=[("Assignee Claims (fx)", "outputs('Config')?['DefaultAssigneeEmail']")],
     check="A G request gets CJ as Assignee.")
step("S1-07", "Flow 1", "Intake comment on the thread", "Inside Try — Comment pattern (Appendix B)",
     ["Add 'Display Name' Compose first (used by Slack too). Then the Comment pattern named 'Intake'."],
     paste=[("Display Name (Compose)", "if(equals(outputs('Config')?['Sensitive'], 'Yes'), '(confidential — see list)', triggerBody()?['EmployeeName'])"),
            ("Intake text (Compose)", "📥 @{outputs('Request_ID')} received from @{triggerBody()?['Author']?['DisplayName']}. SLA target: @{coalesce(outputs('SLA_Due_Date'), outputs('Config')?['SLATarget'])}.@{if(outputs('After_Cutoff'), ' ⚠️ Submitted after the payroll cutoff for this pay period.', '')}"),
            ("Post intake — Body (fx)", "setProperty(json('{}'), 'text', take(outputs('Intake_text'), 1900))")],
     check="Comment appears in the item's Comments pane.")
step("S1-08", "Flow 1", "Ping the People Ops room", "Inside Try → Slack → Post message (V2)",
     ["Channel: People Ops room. Message text below."],
     paste=[("Message text", ":inbox_tray: *New HR change request <@{triggerBody()?['{Link}']}|@{outputs('Request_ID')}>*\n@{triggerBody()?['ChangeType']?['Value']} — @{outputs('Display_Name')}\nEffective @{formatDateTime(triggerBody()?['EffectiveDate'], 'MMM d, yyyy')} • SLA target @{coalesce(outputs('SLA_Due_Date'), outputs('Config')?['SLATarget'])}@{if(outputs('After_Cutoff'), ' • :warning: after payroll cutoff', '')}")],
     check="Message posts; for A2/G the name shows '(confidential — see list)'.")
step("S1-09", "Flow 1", "Confirmation to the submitter", "Inside Try → Office 365 Outlook → Send an email (V2)",
     ["Wording sets expectations without promising a date."],
     paste=[("To (fx)", "triggerBody()?['Author']?['Email']"),
            ("Subject", "We've got your request: @{outputs('Request_ID')}"),
            ("Body", "Hi @{triggerBody()?['Author']?['DisplayName']},\n\nThanks for using the HR front door. Your request @{outputs('Request_ID')} (@{triggerBody()?['ChangeType']?['Value']}) is in the People Ops queue.\n\nTarget completion: @{coalesce(outputs('SLA_Due_Date'), outputs('Config')?['SLATarget'])}. This is an automatic confirmation — a team member will pick it up and keep you posted.\n\nFollow along here: @{triggerBody()?['{Link}']}\n\nPeople Ops")],
     check="Email arrives.")
step("S1-10", "Flow 1", "Add the Catch block", "Below Try (not inside)",
     ["Add the Try / Catch pattern (Appendix C)."], paste=CATCH, check="Configure run after shows only 'has failed' and 'has timed out'.")
step("S1-11", "Test", "Test Flow 1", "Test → Manually → create an E request",
     ["Run T01 and T02 from the Test Pack."], check="T01, T02 = Pass.")

# ---- Flow 2
step("S1-12", "Flow 2", "Create 'HR Change - 2 Checklist Router'", "Automated cloud flow",
     [f"Trigger: SharePoint — When an item is created · List: {LIST_NAME}. (Separate from Flow 1 on purpose: one failing can't stop the other.)",
      "Add Scope 'Try' and the Standard opening block inside it."],
     paste=OPENING, check="Saved.")
step("S1-13", "Flow 2", "Read this change type's checklist", "Inside Try",
     ["Excel → List rows present in a table 'Checklist rows': Table tblChecklist, Filter Query and Order By below.", "Data Operation → Filter array 'Active steps', advanced mode."],
     paste=[("Checklist rows — Filter Query", "Code eq '@{outputs('Code')}'"), ("Checklist rows — Order By", "Step"),
            ("Active steps — From (fx)", "outputs('Checklist_rows')?['body/value']"), ("Active steps — advanced mode", "@equals(item()?['Active'], 'Yes')")],
     check="For E: Active steps returns 5 rows.")
step("S1-14", "Flow 2", "Build the checklist and post it", "Inside Try → Condition 'Has checklist'",
     ["Left (fx) length(body('Active_steps')) · is greater than · 0.", "True branch: two Selects (Map in text mode — the T icon), one Compose, Update item, Comment pattern 'Checklist'."],
     paste=[("Checklist lines — Select From (fx)", "body('Active_steps')"),
            ("Checklist lines — Map (fx)", "concat('[ ] ', item()?['Step'], '. ', item()?['Task'], if(equals(item()?['AfterApproval'], 'Yes'), ' 🔒', ''), if(equals(item()?['DueOnLastDay'], 'Yes'), concat(' 📅 ', formatDateTime(coalesce(triggerBody()?['LastDay'], triggerBody()?['EffectiveDate']), 'MMM d')), ''), if(empty(item()?['Auto']), '', concat(' ⚙ ', item()?['Auto'])))"),
            ("Comment lines — Select From (fx)", "body('Active_steps')"),
            ("Comment lines — Map (fx)", "concat('☐ ', item()?['Step'], '. ', item()?['Task'], if(empty(item()?['Details']), '', concat(' — ', item()?['Details'])))"),
            ("Checklist text (Compose, fx)", f"join(body('Checklist_lines'), {LF})"),
            ("Save checklist (Update item) — Checklist (fx)", "outputs('Checklist_text')"),
            ("Save checklist — ChecklistProgress", "0"),
            ("Save checklist — NextAction (fx)", "concat(first(body('Active_steps'))?['Step'], '. ', first(body('Active_steps'))?['Task'])"),
            ("Checklist text for the thread — 'Checklist comment text' (Compose)", "📋 @{outputs('Code')} checklist — @{length(body('Active_steps'))} steps. Tick them in the Checklist field by changing [ ] to [x].\n@{join(body('Comment_lines'), decodeUriComponent('%0A'))}"),
            ("Post checklist comment — Body (fx)", "setProperty(json('{}'), 'text', take(outputs('Checklist_comment_text'), 1900))")],
     check="E item: Checklist field has 5 '[ ]' lines; Next Action = '01. Confirm field(s) changing and new value'; checklist comment posted.")
step("S1-15", "Flow 2", "Catch-all for 'Other'", "Condition 'Has checklist' → False branch",
     ["Comment pattern 'No template' + Slack post to the People Ops room."],
     paste=[("No template text (Compose)", "⚠️ No checklist for change type @{outputs('Code')} — needs manual scoping. Set an Assignee and write the steps into the Checklist field."),
            ("Slack message", ":warning: <@{triggerBody()?['{Link}']}|@{outputs('Request_ID')}> (@{triggerBody()?['ChangeType']?['Value']}) has no checklist — needs manual scoping.")],
     check="T04 passes.")
step("S1-16", "Flow 2", "Catch block + test", "Below Try",
     ["Add the Catch pattern. Run T03 and T04."], paste=CATCH, check="T03, T04 = Pass.")

# ---- Flow 4
step("S1-17", "Flow 4", "Create 'HR Change - 4 Completion Auto-Comment'", "Automated cloud flow",
     [f"Trigger: SharePoint — When an item is created or modified · List: {LIST_NAME}.",
      "Trigger ⋯ → Settings → Trigger conditions (two lines). They make it fire once per resolution — the flow stamps Completion Date, so its own update can't re-fire it.",
      "Add Scope 'Try' + opening block (you can skip 'Request ID')."],
     paste=[("Trigger condition 1", "@equals(triggerBody()?['Status']?['Value'], '3-Resolved')"), ("Trigger condition 2", "@empty(triggerBody()?['CompletionDate'])")] + OPENING[:5],
     check="Saved.")
step("S1-18", "Flow 4", "Work out outcome, resolver and the ticked checklist", "Inside Try",
     ["Compose × 4, Select × 1, Filter array × 1."],
     paste=[("Outcome", "coalesce(triggerBody()?['Outcome']?['Value'], 'Completed')"),
            ("Resolver", "coalesce(triggerBody()?['Assignee']?['DisplayName'], triggerBody()?['Editor']?['DisplayName'])"),
            ("Today", "convertFromUtc(utcNow(), outputs('Settings')?['TimeZone'])"),
            ("Checklist before", f"split(replace(coalesce(triggerBody()?['Checklist'], ''), {CR}, ''), {LF})"),
            ("Checklist after — Select From (fx) / Map (fx)", "outputs('Checklist_before')   ·   if(and(startsWith(item(), '[ ]'), or(contains(item(), '⚙ resolution'), not(equals(outputs('Outcome'), 'Completed')))), concat(if(equals(outputs('Outcome'), 'Completed'), '[x]', '[–]'), substring(item(), 3)), item())"),
            ("Still open — Filter array From (fx) / advanced", "body('Checklist_after')   ·   @startsWith(item(), '[ ]')")],
     check="On an E resolve: Checklist after shows steps 03–05 as [x].")
step("S1-19", "Flow 4", "Stamp completion and post ✅", "Inside Try",
     ["Update item 'Stamp completion' (required columns from trigger), then Comment pattern 'Resolved'."],
     paste=[("CompletionDate (fx)", "formatDateTime(outputs('Today'), 'yyyy-MM-dd')"),
            ("Outcome Value (fx)", "outputs('Outcome')"),
            ("Checklist (fx)", f"if(empty(triggerBody()?['Checklist']), '', join(body('Checklist_after'), {LF}))"),
            ("NextAction (fx)", "''"),
            ("Resolved text (Compose)", "✅ Resolved on @{formatDateTime(outputs('Today'), 'MMM d, yyyy')} by @{outputs('Resolver')}@{if(equals(outputs('Outcome'), 'Completed'), '', concat(' (Outcome: ', outputs('Outcome'), ')'))}@{if(greater(length(body('Still_open')), 0), concat(' — note: ', string(length(body('Still_open'))), ' checklist step(s) were still open'), '')}")],
     check="Comment reads '✅ Resolved on <date> by <assignee>'.")
step("S1-20", "Flow 4", "Close-the-loop + People Ops summary", "Inside Try",
     ["Condition 'Close the loop?': Left (fx) or(equals(outputs('Config')?['CloseLoopEmail'], 'Yes'), not(equals(outputs('Outcome'), 'Completed'))) · is equal to · (fx) true → Send an email (V2).",
      "Condition 'Summary to People Ops?': Left (fx) equals(outputs('Config')?['PostSummaryToPeopleOps'], 'Yes') · is equal to · (fx) true → Slack Post message (V2) to the People Ops room."],
     paste=[("Email To (fx)", "coalesce(triggerBody()?['RequestedBy']?['Email'], triggerBody()?['Author']?['Email'])"),
            ("Email Subject", "Update on your HR request @{triggerBody()?['RequestId']}"),
            ("Email Body", "Hi @{coalesce(triggerBody()?['RequestedBy']?['DisplayName'], triggerBody()?['Author']?['DisplayName'])},\n\n@{if(equals(outputs('Outcome'), 'Completed'), concat('Good news — your request ', triggerBody()?['RequestId'], ' (', triggerBody()?['ChangeType']?['Value'], ') is complete.'), concat('Your request ', triggerBody()?['RequestId'], ' (', triggerBody()?['ChangeType']?['Value'], ') was closed as ', outputs('Outcome'), '.'))}\n@{if(empty(triggerBody()?['ApprovalHistory']), '', concat('Approval notes: ', triggerBody()?['ApprovalHistory']))}\n\nQuestions? Just reply in #lane-people-talent or open the request: @{triggerBody()?['{Link}']}\n\nPeople Ops"),
            ("Slack summary", ":white_check_mark: *<@{triggerBody()?['{Link}']}|@{triggerBody()?['RequestId']}>* @{outputs('Code')} resolved by @{outputs('Resolver')} (@{outputs('Outcome')}).@{if(greater(length(body('Still_open')), 0), concat(' :warning: ', string(length(body('Still_open'))), ' step(s) still open.'), '')}")],
     check="T05 passes.")
step("S1-21", "Flow 4", "Catch block + test", "Below Try", ["Catch pattern. Run T05."], paste=CATCH, check="T05 = Pass.")
step("S1-22", "Test", "Live E submission", "Real request",
     ["Johanna submits one real Personal Data Change, works the checklist and resolves it (T06)."], check="Gate G1 checklist complete.", owner="Johanna")

# ================= S2
step("S2-01", "Flow 5", "Create 'HR Change - 5 Approval Routing'", "Automated cloud flow",
     [f"Trigger: When an item is created · List: {LIST_NAME}.",
      "BEFORE the Try scope (variables can't live inside a scope): Initialize variable 'Outcome' (String) = Approved; Initialize variable 'History' (String) = empty.",
      "Then Scope 'Try' + opening block."], paste=OPENING, check="Saved.")
step("S2-02", "Flow 5", "Read the approval chain", "Inside Try",
     ["Filter array 'Stages'; Condition 'Needs approval' (Left fx length(body('Stages')) · is greater than · 0). Everything below goes in its True branch."],
     paste=[("Stages — From (fx)", "split(replace(coalesce(outputs('Config')?['ApprovalChain'], ''), ' ', ''), '>')"), ("Stages — advanced", "@not(empty(item()))")],
     check="For I: Stages = [\"HR\"]. For E: empty → nothing happens.")
step("S2-03", "Flow 5", "Mark the request as waiting", "True branch",
     ["Excel List rows 'Approvers rows' (Table tblApprovers, no filter).", "Update item 'Mark pending' (required columns from trigger)."],
     paste=[("ApprovalStatus Value", "Pending"), ("Status Value", "2-Waiting"), ("WaitingOn Value", "Approver")],
     check="Item shows 2-Waiting / Pending / Approver.")
step("S2-04", "Flow 5", "Loop the stages one at a time", "True branch → Apply to each 'Each stage'",
     ["Select an output (fx): body('Stages').", "⚠️ Each stage ⋯ → Settings → Concurrency control ON, Degree of parallelism = 1.",
      "Inside the loop: Condition 'Still approved' — Left (fx) variables('Outcome') · is equal to · Approved. Steps S2-05 to S2-07 go in its True branch."],
     check="Concurrency shows 1.")
step("S2-05", "Flow 5", "Find the approver for this stage", "Still approved → True",
     ["Filter array 'This region' and 'Any region' (From fx outputs('Approvers_rows')?['body/value']), then Compose 'Approver'."],
     paste=[("This region — advanced", "@and(equals(item()?['Role'], items('Each_stage')), equals(item()?['Active'], 'Yes'), equals(toLower(trim(item()?['Region'])), toLower(trim(coalesce(triggerBody()?['Region']?['Value'], '')))))"),
            ("Any region — advanced", "@and(equals(item()?['Role'], items('Each_stage')), equals(item()?['Active'], 'Yes'), equals(trim(item()?['Region']), '*'))"),
            ("Approver (Compose)", "if(equals(items('Each_stage'), 'Manager'), coalesce(triggerBody()?['Manager']?['Email'], ''), coalesce(first(body('This_region'))?['ApproverEmail'], first(body('Any_region'))?['ApproverEmail'], ''))")],
     check="For I: Approver = the HR row's email.")
step("S2-06", "Flow 5", "Ask, record, and stop on reject", "Condition 'Approver found' (Left fx empty(outputs('Approver')) · is equal to · fx false) → True",
     ["Update item 'Show approver'; Approvals → Start and wait for an approval 'Approval' (Approve/Reject – First to respond); Append to string variable 'History'; Comment pattern 'Decision'; Condition 'Rejected?' (Left fx outputs('Approval')?['body/outcome'] · is equal to · Reject) → True: Set variable Outcome = Rejected."],
     paste=[("Show approver — CurrentApprover", "@{items('Each_stage')}: @{outputs('Approver')}"),
            ("Approval — Title", "@{items('Each_stage')} approval: @{triggerBody()?['ChangeType']?['Value']} — @{triggerBody()?['EmployeeName']} [@{outputs('Request_ID')}]"),
            ("Approval — Assigned to (fx)", "outputs('Approver')"),
            ("Approval — Details", "**Employee:** @{triggerBody()?['EmployeeName']} (@{triggerBody()?['EmployeeId']})\n**Change type:** @{triggerBody()?['ChangeType']?['Value']}\n**Effective date:** @{formatDateTime(triggerBody()?['EffectiveDate'], 'MMM d, yyyy')}\n**Job change:** @{triggerBody()?['NewTitle']} @{triggerBody()?['NewDepartment']} @{triggerBody()?['NewLocation']} @{triggerBody()?['EmploymentStatus']?['Value']}\n**Comp:** @{triggerBody()?['NewRate']} @{triggerBody()?['EffectivePayPeriod']}\n**Extra shift:** @{triggerBody()?['DatesWorked']} @{triggerBody()?['ShiftType']?['Value']} @{triggerBody()?['RateType']?['Value']} @{triggerBody()?['RateAmount']} × @{triggerBody()?['HoursOrDays']} — pay period @{triggerBody()?['PayPeriod']}\n**Notes:** @{triggerBody()?['Notes']}"),
            ("Approval — Item link (fx) / description", "triggerBody()?['{Link}']   ·   Open @{outputs('Request_ID')}"),
            ("History — value (end with Enter)", "@{items('Each_stage')}: @{if(equals(outputs('Approval')?['body/outcome'], 'Approve'), 'Approved', 'Rejected')} by @{first(outputs('Approval')?['body/responses'])?['responder']?['displayName']} on @{formatDateTime(convertFromUtc(utcNow(), outputs('Settings')?['TimeZone']), 'MMM d, yyyy')}@{if(empty(first(outputs('Approval')?['body/responses'])?['comments']), '', concat(' — ', first(outputs('Approval')?['body/responses'])?['comments']))}"),
            ("Decision text (Compose)", "@{if(equals(outputs('Approval')?['body/outcome'], 'Approve'), '👍', '👎')} @{items('Each_stage')} @{if(equals(outputs('Approval')?['body/outcome'], 'Approve'), 'approved', 'rejected')} by @{first(outputs('Approval')?['body/responses'])?['responder']?['displayName']}@{if(empty(first(outputs('Approval')?['body/responses'])?['comments']), '', concat(': ', first(outputs('Approval')?['body/responses'])?['comments']))}")],
     check="Approval email arrives; decision comment appears.")
step("S2-07", "Flow 5", "No approver found", "Condition 'Approver found' → False",
     ["Set variable Outcome = NoApprover; Slack post to the People Ops room."],
     paste=[("Slack message", ":rotating_light: <@{triggerBody()?['{Link}']}|@{outputs('Request_ID')}> needs a *@{items('Each_stage')}* approver but none was found (Manager empty, or no active Approvers row for this role/region).")],
     check="T16 passes (Sprint 3).")
step("S2-08", "Flow 5", "Save the result after the loop", "Needs approval → True, AFTER the loop",
     ["SharePoint Get item 'Request now' (Id = trigger ID). Select 'Approval ticks'. Update item 'Save approval result' (required columns from Request now). Comment pattern 'Result'."],
     paste=[("Approval ticks — From (fx)", f"split(replace(coalesce(outputs('Request_now')?['body/Checklist'], ''), {CR}, ''), {LF})"),
            ("Approval ticks — Map (fx)", "if(and(startsWith(item(), '[ ]'), contains(item(), '⚙ approval')), concat('[x]', substring(item(), 3)), item())"),
            ("ApprovalStatus Value (fx)", "if(equals(variables('Outcome'), 'Approved'), 'Approved', if(equals(variables('Outcome'), 'Rejected'), 'Rejected', 'Needs approver'))"),
            ("Status Value (fx)", "if(equals(variables('Outcome'), 'Rejected'), '3-Resolved', if(equals(variables('Outcome'), 'NoApprover'), '2-Waiting', if(empty(outputs('Request_now')?['body/Assignee']), '0-New', '1-Assigned')))"),
            ("Outcome Value (fx)", "if(equals(variables('Outcome'), 'Rejected'), 'Rejected', outputs('Request_now')?['body/Outcome']?['Value'])"),
            ("WaitingOn Value (fx)", "if(equals(variables('Outcome'), 'NoApprover'), 'Approver', null)"),
            ("CurrentApprover (fx) / ApprovalHistory (fx)", "''   ·   variables('History')"),
            ("Checklist (fx)", f"if(or(equals(variables('Outcome'), 'NoApprover'), empty(outputs('Request_now')?['body/Checklist'])), outputs('Request_now')?['body/Checklist'], join(body('Approval_ticks'), {LF}))"),
            ("Result text (Compose)", "@{if(equals(variables('Outcome'), 'Approved'), concat('🔓 Approved (', join(body('Stages'), ' → '), '). 🔒 steps can start now.'), if(equals(variables('Outcome'), 'Rejected'), '⛔ Rejected — case closed as Rejected; the requester will be told.', '⚠️ Approval stopped: no approver found. Fix the Manager field or the Approvers table, then approve offline and log it here.'))}")],
     check="T07: after approving, ApprovalStatus = Approved and step 01 is [x].")
step("S2-09", "Flow 5", "Catch block", "Below Try", ["Catch pattern."], paste=CATCH, check="Saved.")
step("S2-10", "Flow 3", "Create 'HR Change - 3 Payroll Notification' (Template I only)", "Automated cloud flow",
     [f"Trigger: When an item is created or modified · List: {LIST_NAME}. Trigger conditions (4 lines) below.",
      "Try: Settings row + Settings; Slack post to the Payroll room; Update item 'Mark payroll notified'; Comment pattern 'Payroll'. Then Catch.",
      "One-way FYI: it never waits for Payroll (decision D5)."],
     paste=[("Trigger condition 1", "@equals(triggerBody()?['Status']?['Value'], '3-Resolved')"),
            ("Trigger condition 2", "@startsWith(triggerBody()?['ChangeType']?['Value'], 'I ')"),
            ("Trigger condition 3", "@empty(triggerBody()?['PayrollNotifiedOn'])"),
            ("Trigger condition 4", "@not(equals(triggerBody()?['Outcome']?['Value'], 'Rejected'))"),
            ("Slack message (Payroll room)", ":moneybag: *Extra shift earning submitted in Paycor* — <@{triggerBody()?['{Link}']}|@{triggerBody()?['RequestId']}>\nEmployee: @{triggerBody()?['EmployeeName']} (@{triggerBody()?['EmployeeId']})\nDates worked: @{coalesce(triggerBody()?['DatesWorked'], '-')} • Shift: @{coalesce(triggerBody()?['ShiftType']?['Value'], '-')}\nRate: @{coalesce(triggerBody()?['RateType']?['Value'], '-')} @{coalesce(string(triggerBody()?['RateAmount']), '')} • Hours/days: @{coalesce(string(triggerBody()?['HoursOrDays']), '-')}\nPay period: @{coalesce(triggerBody()?['PayPeriod'], '-')}@{if(equals(triggerBody()?['AfterCutoff'], true), ' :warning: submitted after cutoff', '')}\n_FYI — please verify the earning in Paycor. No response needed._"),
            ("Mark payroll notified — PayrollNotifiedOn (fx)", "formatDateTime(convertFromUtc(utcNow(), outputs('Settings')?['TimeZone']), 'yyyy-MM-dd')"),
            ("Payroll text (Compose)", "💰 Payroll notified to verify the Paycor earning (FYI — not a gate).")],
     check="T10 passes.")
step("S2-11", "Test", "Template I pilot tests", "Test Pack", ["Run T07, T08, T09, T10."], check="All pass.")
step("S2-12", "Test", "LOA and offboarding paths", "Test Pack", ["Run T11, T12, T13."], check="All pass.")
step("S2-13", "Launch", "Open the front door", "Slack + Outlook",
     ["Pin the IntakeFormUrl in #lane-people-talent as the official front door (use the announcement text in Appendix D).",
      "HR mailbox: set an automatic reply that points to the form (inspiration: email → form, no flow needed).",
      "Tell managers: 'Not in the intake, does not exist.'"], check="Pinned; auto-reply on; announcement sent.", owner="CJ")
step("S2-14", "Test", "First real Template I case", "Live",
     ["Process the next real extra-shift submission end to end: HR approval → Paycor single event earning → resolve → Payroll FYI."], check="Gate G2 checklist complete.", owner="Johanna")

# ================= S3
step("S3-01", "Decisions", "Approval matrix sign-off", "Stakeholders",
     ["Get written sign-off on who approves B, C, F by region."], check="Dependency marked Confirmed.", owner="CJ")
step("S3-02", "Excel", "Switch on B, C, F approvals", f"{REF_FILE} → Change Types + Approvers",
     ["ApprovalChain: B = 'Manager > Regional', C = 'Manager > Regional > HR', F = 'Manager'.", "Add Approvers rows per region."],
     check="Next B request starts a Manager approval (no flow edit needed).", owner="CJ")
step("S3-03", "Flow 6", "Create 'HR Change - 6 Checklist Progress'", "Automated cloud flow",
     [f"Trigger: When an item is created or modified · List: {LIST_NAME}. Trigger condition below. Trigger Settings → Concurrency control ON, degree 1.",
      "Try: Compose 'Lines', Filter 'Open steps', Filter 'Done steps', Compose 'Progress', Compose 'Next action', Condition 'Changed' → Update item + Condition 'Just finished' → Comment pattern 'Ready' + Slack. Then Catch.",
      "Compare-before-write: the flow only updates when something changed, so its own update ends the loop after one extra run."],
     paste=[("Trigger condition", "@not(empty(triggerBody()?['Checklist']))"),
            ("Lines (Compose)", f"split(replace(triggerBody()?['Checklist'], {CR}, ''), {LF})"),
            ("Open steps — From / advanced", "outputs('Lines')   ·   @startsWith(trim(item()), '[ ]')"),
            ("Done steps — From / advanced", "outputs('Lines')   ·   @startsWith(toLower(trim(item())), '[x]')"),
            ("Progress (Compose)", "if(equals(add(length(body('Open_steps')), length(body('Done_steps'))), 0), 0, div(mul(length(body('Done_steps')), 100), add(length(body('Open_steps')), length(body('Done_steps')))))"),
            ("Next action (Compose)", "if(empty(body('Open_steps')), '', take(trim(substring(trim(first(body('Open_steps'))), 3)), 250))"),
            ("Changed — Left (fx) · is equal to · true", "or(not(equals(float(outputs('Progress')), float(coalesce(triggerBody()?['ChecklistProgress'], 0)))), not(equals(outputs('Next_action'), coalesce(triggerBody()?['NextAction'], ''))))"),
            ("Update item — ChecklistProgress (fx) / NextAction (fx)", "outputs('Progress')   ·   outputs('Next_action')"),
            ("Just finished — Left (fx) · is equal to · true", "and(equals(outputs('Progress'), 100), less(float(coalesce(triggerBody()?['ChecklistProgress'], 0)), 100), not(equals(triggerBody()?['Status']?['Value'], '3-Resolved')))"),
            ("Ready text (Compose)", "☑️ All checklist steps are done — set Status to 3-Resolved to close the case (LOA: only once the return is confirmed)."),
            ("Slack message", ":ballot_box_with_check: All steps done on <@{triggerBody()?['{Link}']}|@{triggerBody()?['RequestId']}> — @{coalesce(triggerBody()?['Assignee']?['DisplayName'], 'owner')}, ready to resolve.")],
     check="T17 passes.")
step("S3-04", "Test", "Approval and progress tests", "Test Pack", ["Run T14, T15, T16, T17, T22."], check="All pass.")
step("S3-05", "Test", "UAT with Johanna", "Every path A1–I",
     ["Submit one of each change type; work and resolve each. Log issues in the Notes column."], check="Gate G3 checklist complete.", owner="Johanna")

# ================= S4
step("S4-01", "Flow 7", "Create 'HR Change - 7 Daily Digest + Cutoff Reminder'", "Scheduled cloud flow — every 1 day at 8:00 (your time zone)",
     ["Try: Settings row + Settings; Compose 'Today' and 'Today key'; Excel List rows 'Change types' (tblChangeTypes); Filter 'Sensitive types'; Select 'Sensitive codes'; SharePoint Get items 'Open requests'."],
     paste=[("Today / Today key", "convertFromUtc(utcNow(), outputs('Settings')?['TimeZone'])   ·   formatDateTime(outputs('Today'), 'yyyyMMdd')"),
            ("Sensitive types — From / advanced", "outputs('Change_types')?['body/value']   ·   @equals(item()?['Sensitive'], 'Yes')"),
            ("Sensitive codes — Select From / Map (text mode)", "body('Sensitive_types')   ·   item()?['Code']"),
            ("Open requests — Filter Query / Order By / Top Count", "Status ne '3-Resolved'   ·   SLADueDate asc   ·   5000")],
     check="Run manually: Open requests returns your test items.")
step("S4-02", "Flow 7", "Digest sections", "Inside Try",
     ["Five Filter arrays named exactly as below (From fx outputs('Open_requests')?['body/value']), and one Select per filter named '<filter> lines' (e.g. 'SLA overdue lines', From fx body('SLA_overdue')) using the same Map. Then Condition 'Weekday with news' → Slack post to the People Ops room."],
     paste=[("Unassigned", "@equals(item()?['Status']?['Value'], '0-New')"),
            ("SLA overdue", "@and(not(empty(item()?['SLADueDate'])), less(formatDateTime(item()?['SLADueDate'], 'yyyyMMdd'), outputs('Today_key')))"),
            ("Action due", "@and(not(empty(item()?['NextActionDue'])), lessOrEquals(formatDateTime(item()?['NextActionDue'], 'yyyyMMdd'), outputs('Today_key')))"),
            ("Exceptions", "@not(empty(item()?['AutomationNote']))"),
            ("LOA returns", "@and(startsWith(coalesce(item()?['ChangeType']?['Value'], ''), 'G '), not(empty(item()?['EndDate'])), lessOrEquals(ticks(item()?['EndDate']), ticks(addDays(utcNow(), 7))))"),
            ("Line Map for every Select (text mode)", "concat('• <', item()?['{Link}'], '|', item()?['RequestId'], '> ', first(split(item()?['ChangeType']?['Value'], ' ')), ' — ', if(contains(body('Sensitive_codes'), first(split(item()?['ChangeType']?['Value'], ' '))), '(confidential)', item()?['EmployeeName']), ' — ', coalesce(item()?['Assignee']?['DisplayName'], 'unassigned'), if(empty(item()?['NextAction']), '', concat(' — next: ', item()?['NextAction'])))"),
            ("Weekday with news — Left (fx) · is equal to · true", "and(greater(dayOfWeek(outputs('Today')), 0), less(dayOfWeek(outputs('Today')), 6), greater(add(add(add(length(body('Unassigned')), length(body('SLA_overdue'))), add(length(body('Action_due')), length(body('Exceptions')))), length(body('LOA_returns'))), 0))"),
            ("Slack message", ":bar_chart: *HR Change — daily check* (@{length(outputs('Open_requests')?['body/value'])} open)\n@{if(empty(body('Unassigned_lines')), '', concat('*:white_circle: 0-New, not picked up*', decodeUriComponent('%0A'), join(body('Unassigned_lines'), decodeUriComponent('%0A')), decodeUriComponent('%0A')))}@{if(empty(body('SLA_overdue_lines')), '', concat('*:red_circle: SLA overdue*', decodeUriComponent('%0A'), join(body('SLA_overdue_lines'), decodeUriComponent('%0A')), decodeUriComponent('%0A')))}@{if(empty(body('Action_due_lines')), '', concat('*:large_yellow_circle: Next action due*', decodeUriComponent('%0A'), join(body('Action_due_lines'), decodeUriComponent('%0A')), decodeUriComponent('%0A')))}@{if(empty(body('Exceptions_lines')), '', concat('*:warning: Automation exceptions*', decodeUriComponent('%0A'), join(body('Exceptions_lines'), decodeUriComponent('%0A')), decodeUriComponent('%0A')))}@{if(empty(body('LOA_returns_lines')), '', concat('*:calendar: LOA returns within 7 days*', decodeUriComponent('%0A'), join(body('LOA_returns_lines'), decodeUriComponent('%0A'))))}")],
     check="T18 passes.")
step("S4-03", "Flow 7", "Payroll cutoff reminder", "Inside Try",
     ["Compose 'Cutoff days'; Select 'This month cutoffs' and 'Next month cutoffs' (Map text mode); Filter 'Remind today'; Condition (length > 0) → Slack post to the managers channel. Then Catch (email only)."],
     paste=[("Cutoff days (Compose)", "createArray(int(outputs('Settings')?['CutoffDay1']), int(outputs('Settings')?['CutoffDay2']))"),
            ("This month cutoffs — Map", "formatDateTime(addDays(startOfMonth(outputs('Today')), sub(item(), 1)), 'yyyy-MM-dd')"),
            ("Next month cutoffs — Map", "formatDateTime(addDays(startOfMonth(addToTime(outputs('Today'), 1, 'Month')), sub(item(), 1)), 'yyyy-MM-dd')"),
            ("Remind today — From (fx)", "union(body('This_month_cutoffs'), body('Next_month_cutoffs'))"),
            ("Remind today — advanced", "@equals(formatDateTime(outputs('Today'), 'yyyy-MM-dd'), formatDateTime(addDays(addDays(item(), mul(int(outputs('Settings')?['ReminderBusinessDays']), -1)), if(equals(dayOfWeek(addDays(item(), mul(int(outputs('Settings')?['ReminderBusinessDays']), -1))), 6), -1, if(equals(dayOfWeek(addDays(item(), mul(int(outputs('Settings')?['ReminderBusinessDays']), -1))), 0), -2, 0))), 'yyyy-MM-dd'))"),
            ("Slack message (managers)", ":alarm_clock: *Payroll cutoff is @{formatDateTime(first(body('Remind_today')), 'dddd, MMM d')}.* Any extra shifts or fractional FTE pickups to report for this pay period? Submit them as *I — Fractional FTE / Extra Shift Pickup*: <@{outputs('Settings')?['IntakeFormUrl']}|HR change request form>")],
     check="T19 passes.")
step("S4-04", "Test", "Exception drill", "Flow 1",
     ["Run T20: make one action fail on purpose, confirm the alert, the Automation Note and the digest Exceptions line, then undo."], check="T20 = Pass.")
step("S4-05", "Operate", "Two weeks of steady operation", "Views + digest",
     ["Use the digest every morning. Redirect any request raised in DMs/email to the form.", "End of week 2: review All Open by SLA and Recently Resolved for the SLA baseline."],
     check="Gate G4 checklist complete.", owner="CJ")

# ================= S5 (dependency-driven; outline-level)
step("S5-01", "Downstream", "Jira ticket for offboarding (A1/A2)", "New flow or a branch in Flow 1",
     ["When the IT dependency is confirmed: on A1/A2 create → Jira 'Create a new issue' (access revocation + equipment retrieval) → write the issue URL to JiraLink → comment.",
      "In the Checklists table set Auto = 'jira' on the 'Create Jira ticket' steps and tick them in the same flow."], check="A test A1 creates a Jira issue and fills Jira Ticket Link.", owner="CJ")
step("S5-02", "Downstream", "Offboarding distro email", "Flow 1 branch",
     ["When the distro address is confirmed: A1 → email the distro at intake. A2 stays manual (timing-sensitive access revocation)."], check="Test A1 emails the distro.", owner="CJ")
step("S5-03", "Downstream", "Word letters", "Word Online (Business) → Populate a Microsoft Word template (premium)",
     ["Termination letter (A2), comp change confirmation (C), job change confirmation (B). Save to the HR file; attach to the item."], check="Generated letter opens correctly.", owner="CJ")
step("S5-04", "Downstream", "Monday.com payroll board + backfill link", "Monday connector",
     ["When the board/column is confirmed: post A1/A2 final pay and C rate changes to the payroll board; send the backfill form link to the manager on A1."], check="Item appears on the board.", owner="CJ")

# ================= S6 (Run)
step("S6-01", "Run", "Manager self-service form", "Microsoft Forms → flow → same list", ["Scoped types (e.g. E, F, I), readiness checklist first."], owner="CJ")
step("S6-02", "Run", "Dashboard", "List views → Power BI (optional)", ["Open by age, SLA compliance, change-type breakdown."], owner="CJ")
step("S6-03", "Run", "In-Slack approvals", "Bot token", ["Replace email approvals with Slack buttons once the bot token is approved."], owner="CJ")
step("S6-04", "Run", "Paycor write-back", "Paycor iPaaS", ["Auto-write approved changes instead of manual entry."], owner="CJ")
step("S6-05", "Run", "Holiday-aware SLA", "Reference workbook", ["Add a tblHolidays table and skip those dates in Flow 1."], owner="CJ")

# ------------------------------------------------------------------ test pack
TESTS = [
    ("T01", "S1", "E — basic intake", "Submit an E request on a weekday.", "Request = 'E — Personal Data Change — <name>'; Request ID HRC-00000N; SLA = next business day; 📥 comment; People Ops Slack ping; confirmation email."),
    ("T02", "S1", "Weekend SLA", "Submit an E request on Saturday (or check SLA Start in the run of a Friday-evening test).", "Saturday submission → SLA = Tuesday (counting starts Monday)."),
    ("T03", "S1", "E checklist", "Look at the T01 item.", "Checklist has 5 '[ ] 0n.' lines; 03–05 end with '⚙ resolution'; Next Action = '01. Confirm field(s) changing and new value'; 📋 comment lists the steps."),
    ("T04", "S1", "Other — catch-all", "Submit Change Type = Other.", "No checklist; '⚠️ … needs manual scoping' comment; Slack warning."),
    ("T05", "S1", "Resolve E", "Set Assignee = you, Status = 3-Resolved.", "Completion Date = today; Outcome = Completed; 03–05 ticked [x]; '✅ Resolved on <date> by <you>' (+ note if 01–02 were open); close-the-loop email."),
    ("T06", "S1", "Live E", "Johanna submits and works a real E request.", "Same as T01–T05 with no manual fixes."),
    ("T07", "S2", "I — HR approval (approve)", "Submit I: Rate Type Hourly, $100, first date worked in the current pay period before cutoff.", "2-Waiting / Pending / Waiting On Approver; HR approval email; approve → Approved, step 01 [x], 🔓 comment; Status 0-New (no Assignee) or 1-Assigned."),
    ("T08", "S2", "I — after cutoff", "Submit I with first date worked in a pay period whose cutoff has passed.", "Pay Period label set; Submitted After Cutoff = Yes; ⚠️ in comment and Slack."),
    ("T09", "S2", "I — HR rejects", "Submit I; reject with a comment.", "Status 3-Resolved; Outcome Rejected; open steps [–]; rejection email with approval notes; NO Payroll FYI."),
    ("T10", "S2", "I — Payroll FYI", "Resolve the T07 request.", "Payroll room message once; Payroll Notified On = today; editing the item again does not re-post."),
    ("T11", "S2", "G — LOA", "Submit G with an expected return date.", "SLA blank ('Per policy' in comment); Assignee = CJ; Slack shows '(confidential — see list)'; ticking every step does not resolve it."),
    ("T12", "S2", "A1 — voluntary offboarding", "Submit A1 with Last Day next Friday.", "SLA Due Date = Last Day; 16 steps; 'Terminate in Paycor' line shows 📅 <last day>."),
    ("T13", "S2", "A2 — involuntary (sensitive)", "Submit A2.", "Name withheld in Slack and the digest; SLA = Last Day; 12 steps."),
    ("T14", "S3", "B — two-stage approval", "Submit B (Manager = you, Region with a Regional row).", "Manager approval first; Regional only after it is approved; both lines in Approval History."),
    ("T15", "S3", "C — reject at stage 2", "Submit C; approve Manager, reject Regional.", "HR stage never sent; Rejected path as in T09."),
    ("T16", "S3", "No approver", "Submit B with a Region that has no Regional row and with the '*' row set Active = No.", "Approval Status = Needs approver; Status 2-Waiting; Slack alert. (Restore the '*' row.)"),
    ("T17", "S3", "Checklist progress", "On any open request change two '[ ]' to '[x]'; then tick the rest.", "Checklist % updates; Next Action = next open step; at 100%: '☑️ … ready to resolve' comment + Slack; Status unchanged."),
    ("T18", "S4", "Daily digest", "Run Flow 7 manually with open test items (one overdue, one 0-New, one exception).", "One Slack post with the right sections; A2/G names show '(confidential)'."),
    ("T19", "S4", "Cutoff reminder", "Temporarily set CutoffDay1 to a date ReminderBusinessDays business days from today; run Flow 7.", "Managers channel reminder with the form link. Restore the setting."),
    ("T20", "S4", "Exception drill", "Temporarily point Flow 1's Slack step at a channel the service account can't post to; submit E.", "Admin email with run link; Automation Note filled; item in the Exceptions view and digest. Undo."),
    ("T21", "S0", "Permissions", "A non-HR colleague submits a request, then browses the list and searches for another employee's request.", "They see only their own item."),
    ("T22", "S3", "Config without flow edits", "Change E's SLABusinessDays to 2 in Excel; submit E.", "SLA = 2 business days. Restore to 1."),
]

ANNOUNCEMENT = (
    "📣 *New: one front door for HR changes*\n"
    "From today, every HR change request (offboarding, job/status, comp, pay/banking, personal data, time/PTO, LOA, benefits, extra shifts) "
    "goes through this form: <IntakeFormUrl>\n"
    "• You'll get a confirmation with a request ID and a target date.\n"
    "• If it needs approval, the approver gets it automatically.\n"
    "• You'll hear back when it's done.\n"
    "Requests sent by DM or email will be pointed back here — if it isn't in the intake, we can't track it. Thank you! — People Ops"
)
AUTO_REPLY = (
    "Thanks for reaching out to People Ops. HR change requests are handled through our request form so nothing gets lost: <IntakeFormUrl>. "
    "For anything else, a member of the team will reply soon."
)
