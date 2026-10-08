"""Single source of truth for the HR Change + Orchestration build kit (v3).

v3: the one list keeps only the core service-desk columns; everything request-specific lives in the
Description and the sub-step Checklist. Change types are general HR request types. Microsoft Forms is
the front door, so the list is HR-only.

Everything the generator writes (reference workbook, tracker, Word blueprint, list CSV) comes from here.
"""
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))

LIST_NAME = "HR Change Requests"
REF_FILE = "HR Change Reference.xlsx"
FORM_NAME = "HR Request"
CR = "decodeUriComponent('%0D')"
LF = "decodeUriComponent('%0A')"

# ------------------------------------------------------------------ reference tables (flows read these)
SETTINGS = {
    "RequestPrefix": ("HRC-", "Prefix for Request IDs (HRC-000123)."),
    "TimeZone": ("Eastern Standard Time", "Windows time zone name used for dates, SLAs and the daily digest."),
    "CutoffDay1": (10, "Payroll cutoff day of month for the 1st–15th pay period."),
    "CutoffDay2": (25, "Payroll cutoff day of month for the 16th–end pay period (28 or lower)."),
    "ReminderBusinessDays": (2, "Business days before each cutoff that managers get the extra-pay reminder."),
    "FormUrl": ("https://forms.office.com/r/<your-form-id>", "The front door (Microsoft Form link). Pinned in #lane-people-talent."),
    "ServiceAccountEmail": ("svc-peopleops@yourcompany.com", "The account the flows' SharePoint connection uses. Edits by this account are not counted as a human first response."),
    "AdminEmail": ("cj@yourcompany.com", "Who receives automation failure alerts."),
}

# Code, ChangeType (exact choice text), Group, DefaultPriority, SLA days, DueDateRule, SLATarget, ApprovalChain,
# DefaultAssigneeEmail, CloseLoopEmail, PostSummaryToPeopleOps, NotifyPayrollOnClose, Sensitive, InfoNeeded, PlanRef, Notes
CHANGE_TYPES = [
    ("ONB", "ONB — Onboarding / New Hire", "Employee lifecycle", "Time-sensitive", "", "Effective date", "Ready by start date", "", "", "Yes", "Yes", "No",
     "Start date (as Effective date), title, department, location, manager, FT/PT/PRN, equipment and system needs", "new", "Effective date = start date."),
    ("JOB", "JOB — Job / Status Change", "Employee lifecycle", "Routine", 3, "Business days", "2–3 business days", "", "", "Yes", "No", "No",
     "New title, department, location, FT/PT/PRN, new manager (if changing)", "B", "Planned chain (switch on at Sprint 3, after matrix sign-off): Manager > Regional"),
    ("COMP", "COMP — Compensation Change", "Pay & time", "Routine", 5, "Business days", "3–5 business days", "", "", "Yes", "No", "No",
     "New rate, effective pay period, reason (merit, promotion, market)", "C", "Planned chain (switch on at Sprint 3): Manager > Regional > HR"),
    ("PAY", "PAY — Payroll / Banking / Tax", "Pay & time", "Time-sensitive", 2, "Business days", "1–2 business days", "", "", "Yes", "No", "No",
     "What's changing (direct deposit, tax withholding, garnishment, pay discrepancy) and whether verification is ready. Never type account numbers — HR collects them securely", "D", "Manual verification by CJ or Johanna."),
    ("XPAY", "XPAY — Extra Shift / Additional Pay", "Pay & time", "Time-sensitive", 2, "Business days", "Before the payroll cutoff", "HR", "", "No", "No", "Yes",
     "First date worked (as Effective date), all date(s) worked, shift type (training, weekend clinic, coverage, other), hourly vs flat/day rate, rate, hours or days, reason", "I",
     "HR approval live from Sprint 2 (pilot). Payroll gets an FYI at close — verifier, not a gate."),
    ("PTO", "PTO — Time Off", "Pay & time", "Routine", 2, "Business days", "1–2 business days", "", "", "Yes", "No", "No",
     "Type of time off, dates, hours", "F", "Planned chain (switch on at Sprint 3): Manager"),
    ("LOA", "LOA — Leave of Absence / Accommodation", "Benefits & leave", "Time-sensitive", "", "None", "Per policy", "", "cj@yourcompany.com", "No", "No", "No",
     "Leave or accommodation type, start date (as Effective date), expected return, whether documentation is ready. No medical details here", "G",
     "Routes to CJ. Stays open until return is confirmed — set Next Action Due to the expected return. Sensitive."),
    ("BEN", "BEN — Benefits / Life Event", "Benefits & leave", "Routine", 3, "Business days", "2–3 business days", "", "", "Yes", "No", "No",
     "Event type, qualifying date (as Effective date), whether documentation is ready", "H", ""),
    ("DATA", "DATA — Employee Data Update", "Service & records", "Routine", 1, "Business days", "1 business day", "", "", "Yes", "No", "No",
     "Field(s) changing and the new value", "E", "Build and test this path first (Sprint 1)."),
    ("VER", "VER — Verification / Letter", "Service & records", "Routine", 2, "Business days", "1–2 business days", "", "", "Yes", "No", "No",
     "What's needed (employment / income verification, letter), who it's for, deadline, signed authorization for third parties", "new", ""),
    ("TRN", "TRN — Training / Credentialing", "Service & records", "Routine", 3, "Business days", "2–3 business days", "", "", "Yes", "No", "No",
     "Training, license or credential, due date", "new", ""),
    ("POL", "POL — Policy / General HR Question", "Service & records", "Routine", 2, "Business days", "1–2 business days", "", "", "No", "No", "No",
     "Your question, and any deadline", "new", "The 'Reply to the requester' step is the close-the-loop."),
    ("ER", "ER — Employee Relations Concern", "Employee relations", "Urgent", 1, "Business days", "Personal contact within 1 business day", "", "cj@yourcompany.com", "No", "No", "No",
     "A short summary only — HR will contact you to talk it through", "new",
     "Sensitive. Keep only a summary on the list; details go to the restricted case file."),
    ("REQ", "REQ — Hiring Requisition / Backfill", "Employee lifecycle", "Routine", 3, "Business days", "2–3 business days", "", "", "Yes", "No", "No",
     "Role, new vs backfill (who left), department, location, target start", "new", "Approval chain to be confirmed with the matrix."),
    ("OFF-V", "OFF-V — Offboarding (Voluntary)", "Employee lifecycle", "Time-sensitive", "", "Effective date", "Same day intake, executed by last day", "", "", "No", "Yes", "No",
     "Last day (as Effective date), reason (resignation, end of contract, mutual separation), equipment to retrieve, final pay type, resignation letter", "A1", ""),
    ("OFF-I", "OFF-I — Offboarding (Involuntary)", "Employee lifecycle", "Urgent", "", "Effective date", "Same day intake, executed by last day", "", "", "No", "Yes", "No",
     "Last day (as Effective date), reason (termination, reduction in force, non-renewal, performance), equipment to retrieve, final pay type, timing for access removal", "A2",
     "Sensitive: employee name withheld from Slack and the digest."),
    ("OTHER", "OTHER — Something else", "Service & records", "Routine", "", "None", "Needs manual scoping", "", "", "Yes", "No", "No",
     "Describe what you need", "Other", "No checklist on purpose: the router posts 'needs manual scoping'."),
]
SENSITIVE = {"OFF-I", "LOA", "ER"}
CHANGE_TYPE_COLS = ["Code", "ChangeType", "Group", "DefaultPriority", "SLABusinessDays", "DueDateRule", "SLATarget", "ApprovalChain", "DefaultAssigneeEmail",
                    "CloseLoopEmail", "PostSummaryToPeopleOps", "NotifyPayrollOnClose", "Sensitive", "InfoNeeded", "PlanRef", "Notes"]


def change_type_rows():
    rows = []
    for c in CHANGE_TYPES:
        r = list(c[:12]) + ["Yes" if c[0] in SENSITIVE else "No"] + list(c[13:])
        rows.append(r)
    return rows


with open(os.path.join(HERE, "checklists.json"), encoding="utf-8") as f:
    CHECKLISTS = json.load(f)
CHECKLIST_COLS = ["Code", "Step", "Task", "Details", "Owner", "Auto", "AfterApproval", "DueOnLastDay", "Active"]

APPROVERS = [
    ("Regional", "*", "regional.director@yourcompany.com", "Default Regional approver", "Yes", "Replace. Add one row per region (Region must match the list's Region choice)."),
    ("HR", "*", "cj@yourcompany.com", "HR approver", "Yes", "Used by COMP (stage 3) and XPAY (pilot)."),
    ("Payroll", "*", "payroll@yourcompany.com", "Payroll approver", "No", "Only if Payroll ever becomes an approval stage (decision D5). Inactive by default."),
]
APPROVER_COLS = ["Role", "Region", "ApproverEmail", "ApproverName", "Active", "Notes"]

# ------------------------------------------------------------------ the ONE list: core columns only
CT_CHOICES = [c[1] for c in CHANGE_TYPES]
FIELDS = [
    # internal, display, type, choices/default, source (inspiration core / plan universal / workflow), set by
    ("Title", "Request", "Single line of text", "", "Core", "Flow 1: '<Change Type> — <Employee Name>'"),
    ("RequestId", "Ticket Reference", "Single line of text", "", "Core", "Flow 1: HRC-000123"),
    ("ChangeType", "Category", "Choice", "; ".join(CT_CHOICES), "Core", "Form"),
    ("Priority", "Priority", "Choice", "Urgent; Time-sensitive; Routine", "Core", "Form, else the change type's default (Flow 1)"),
    ("Status", "Status", "Choice", "0-New; 1-Assigned; 2-Waiting; 3-Resolved  (default 0-New)", "Core", "HR (approvals flow sets 2-Waiting / result)"),
    ("AssignedTo", "Assigned To", "Person", "", "Core", "HR, or the change type's default assignee (Flow 1)"),
    ("RequesterName", "Requester Name", "Single line of text", "", "Core", "Flow 0 (form responder)"),
    ("RequesterEmail", "Requester Email", "Single line of text", "", "Core", "Flow 0 (form responder)"),
    ("EmployeeName", "Employee Name", "Single line of text", "", "Plan universal", "Form"),
    ("EmployeeId", "Employee ID", "Single line of text", "", "Plan universal", "Form"),
    ("ManagerEmail", "Manager Email", "Single line of text", "", "Plan universal", "Form (manager approvals route here)"),
    ("Region", "Region", "Choice", "East; Central; West  (replace with yours)", "Workflow", "Form (picks the Regional approver)"),
    ("EffectiveDate", "Effective Date", "Date only", "", "Plan universal", "Form (start date / last day / first date worked / leave start)"),
    ("Description", "Description", "Multiple lines of text (plain)", "", "Core", "Form — everything request-specific lives here"),
    ("ReceivedDate", "Received Date", "Date and time", "", "Core", "Flow 1 (= submission time)"),
    ("SLADueDate", "SLA Due Date", "Date only", "", "Plan", "Flow 1 (from the change type)"),
    ("FirstResponseDate", "First Human Response", "Date and time", "", "Core", "Flow 3 (first edit by a person that moves Status off 0-New)"),
    ("NextAction", "Next Action", "Single line of text", "", "Core", "Flow 2/3 (first open checklist step); HR can overwrite"),
    ("NextActionDue", "Next Action Due", "Date only", "", "Core", "Flow 1 (= SLA); HR updates"),
    ("WaitingOn", "Waiting On", "Choice", "Requester; Employee; Manager; Approver; Payroll; IT; Third party; Other", "Core (adapted)", "HR with 2-Waiting (approvals flow sets 'Approver')"),
    ("Checklist", "Checklist", "Multiple lines of text (plain)", "", "Workflow", "Flow 2 writes the sub-steps; HR ticks [ ] → [x]"),
    ("ChecklistProgress", "Checklist %", "Number (0–100)", "default 0", "Workflow", "Flow 3"),
    ("ApprovalStatus", "Approval Status", "Choice", "Not required; Pending; Approved; Rejected; Needs approver  (default Not required)", "Workflow", "Flow 5"),
    ("CurrentApprover", "Current Approver", "Single line of text", "", "Workflow", "Flow 5"),
    ("Outcome", "Outcome", "Choice", "Completed; Rejected; Withdrawn", "Workflow", "Flow 4 (Completed) / Flow 5 (Rejected) / HR (Withdrawn)"),
    ("ClosedDate", "Closed Date", "Date only", "", "Core", "Flow 4 (auto-stamp)"),
    ("InternalNotes", "Internal Notes", "Multiple lines of text (plain)", "", "Core", "HR; Flow 5 appends approval decisions"),
    ("AutomationNote", "Automation Note", "Single line of text", "", "Core (exceptions)", "Catch blocks"),
]
OMITTED_CORE = [
    ("Initial message identifier", "Only needed for email intake (duplicate detection). Add back if email intake is ever turned on."),
    ("Conversation identifier", "Same — email reply matching only."),
    ("Last inbound date", "Same — tracks employee email replies. With form intake, follow-ups happen in Slack or the item comments."),
]

CSV_SAMPLE = {
    "Title": "SAMPLE - delete after import", "RequestId": "HRC-000000", "ChangeType": CT_CHOICES[8], "Priority": "Routine",
    "Status": "0-New", "RequesterName": "Sample Requester", "RequesterEmail": "requester@yourcompany.com", "EmployeeName": "Sample Employee",
    "EmployeeId": "000000", "ManagerEmail": "manager@yourcompany.com", "Region": "East", "EffectiveDate": "2026-10-15",
    "Description": "Sample row - delete", "ReceivedDate": "2026-10-14 09:30", "SLADueDate": "2026-10-15", "FirstResponseDate": "2026-10-14 10:00",
    "NextAction": "01. Confirm field(s) changing and new value", "NextActionDue": "2026-10-15", "WaitingOn": "Requester",
    "Checklist": "[ ] 01. Sample step", "ChecklistProgress": "0", "ApprovalStatus": "Not required", "CurrentApprover": "none",
    "Outcome": "Completed", "ClosedDate": "2026-10-15", "InternalNotes": "none", "AutomationNote": "none",
}

FORM_QUESTIONS = [
    ("1", "What do you need?", "Choice (drop-down)", "Yes", "The Category values, exactly (see Change Types).", "ChangeType"),
    ("2", "Employee name", "Text", "Yes", "Who the request is about (can be you).", "EmployeeName"),
    ("3", "Employee ID", "Text", "Yes", "", "EmployeeId"),
    ("4", "Manager's email", "Text", "Yes", "Manager approvals go here.", "ManagerEmail"),
    ("5", "Region", "Choice", "Yes", "Same values as the list's Region column.", "Region"),
    ("6", "Effective date", "Date", "Yes", "Start date, last day, first date worked or leave start — whichever fits.", "EffectiveDate"),
    ("7", "How urgent is this?", "Choice", "No", "Urgent; Time-sensitive; Routine. Blank = the change type's default.", "Priority"),
    ("8", "Details", "Long text", "Yes",
     "Subtitle: 'Tell us everything we need (we'll email you the checklist of what's needed for your request type). Please don't include bank account numbers, SSNs or medical details — we'll collect those securely.'",
     "Description"),
]

STATUSES = [
    ("0-New", "Submitted, not yet picked up", "Column default at intake; approvals flow returns here if approved with nobody assigned", "Pick it up: set Assigned To, move to 1-Assigned (this stamps First Human Response)"),
    ("1-Assigned", "Has an owner; work in progress", "HR, or the approvals flow after approval when someone is assigned", "Work the Checklist; keep Next Action / Next Action Due current"),
    ("2-Waiting", "Owner acted; waiting on someone", "HR (set Waiting On), or the approvals flow (Waiting On = Approver)", "Chase whoever Waiting On names"),
    ("3-Resolved", "All steps complete; case closed", "HR, or the approvals flow on rejection (Outcome = Rejected)", "Flow 4 stamps Closed Date + ✅ comment + close-the-loop"),
]

VIEWS = [
    ("New / Unassigned", "Status = 0-New", "ReceivedDate ascending", "RequestId, Title, Priority, SLADueDate, ReceivedDate", "Nobody owns these yet (Tag, You're It)."),
    ("My Active", "AssignedTo = [Me] AND Status ≠ 3-Resolved", "NextActionDue ascending", "RequestId, Title, Status, NextAction, NextActionDue, ChecklistProgress", "Your workbench."),
    ("Needs Attention Today", "Status ≠ 3-Resolved AND NextActionDue ≤ [Today]", "Priority, then NextActionDue", "RequestId, Title, AssignedTo, NextAction, NextActionDue", "Every active case has a next action and a date."),
    ("Waiting", "Status = 2-Waiting", "Group by WaitingOn", "RequestId, Title, AssignedTo, WaitingOn, CurrentApprover, Modified", "Who we're waiting on."),
    ("Awaiting Approval", "ApprovalStatus = Pending", "ReceivedDate ascending", "RequestId, Title, CurrentApprover", "Approvals in flight."),
    ("All Open by SLA", "Status ≠ 3-Resolved", "SLADueDate ascending; group by ChangeType", "RequestId, Title, Priority, Status, AssignedTo, SLADueDate, ChecklistProgress", "Dashboard-lite (Phase 3)."),
    ("Recently Closed", "Status = 3-Resolved AND ClosedDate ≥ [Today]-14", "ClosedDate descending", "RequestId, Title, Outcome, ClosedDate, AssignedTo", "Review completed work."),
    ("Exceptions", "AutomationNote is not empty", "Modified descending", "RequestId, Title, AutomationNote, Modified", "Anything an automation could not finish."),
]

# ------------------------------------------------------------------ decisions, inspiration review, plan traceability
DECISIONS = [
    ("D1", "System of record", "ONE SharePoint list, HR Change Requests, replaces the HR Change Slack List.", "Your direction.", "Decided"),
    ("D2", "List shape", "The list keeps only the core service-desk columns (from the inspiration) plus the plan's universal fields. Everything request-specific goes in Description and the sub-step Checklist.",
     "One list that stays simple for any HR request. No per-type columns to maintain.", "Decided"),
    ("D3", "Front door", "A Microsoft Form (org-only, records the responder) pinned in #lane-people-talent. The list is HR-only; HR can still add items directly for walk-ins or calls.",
     "Requesters never need access to the HR queue (inspiration: don't expose the queue), so Internal Notes stay internal. A Slack form can't write to SharePoint without a bot token.", "Confirm (CJ)"),
    ("D4", "Change types", "16 general HR request types (+ Other) in five groups. Your A1–I types map into them (PlanRef column).",
     "Covers the whole HR team's work, not only HR changes. Add or retire types in Excel.", "Confirm (CJ)"),
    ("D5", "XPAY — Payroll role", "Payroll is a verifier, not a gate: a one-way Slack FYI when People Ops closes an XPAY case.",
     "The plan's Flow 3 section says this; the master reference's 'Payroll final approval' line conflicts. To make Payroll a gate, add 'Payroll' to XPAY's ApprovalChain.", "Confirm (CJ)"),
    ("D6", "Reference data", "SLAs, approval chains, checklists, approvers, settings and 'info needed' prompts live in one Excel workbook the flows read.",
     "Managed like a reference sheet. Contract: never rename a table or a column header.", "Decided"),
    ("D7", "Checklist / sub-steps", "Flow 2 posts the steps to the item's comment thread and writes them to the Checklist field. HR ticks [ ] → [x]; Flow 3 keeps Checklist % and Next Action current.",
     "The plan's Phase 1 thread checklist plus the Phase 2 progress goal, without a second list.", "Decided"),
    ("D8", "Plan Flow 3 (Payroll Notification)", "Merged into Flow 4 Close-out: it runs once per closure, and posts the Payroll FYI when the change type says NotifyPayrollOnClose = Yes.",
     "Same behaviour, one less flow, and no extra tracking column.", "Decided"),
    ("D9", "First human response", "Stamped automatically the first time a person (not the service account) moves a request off 0-New.",
     "The inspiration's first-response date without manual entry. Needs the flows' SharePoint connection on a service account.", "Decided"),
    ("D10", "Approvals", "Power Automate Approvals (email, Teams, app), sequential, from ApprovalChain. Decisions go to the comment thread and Internal Notes. In-Slack buttons wait for the Phase 3 bot token.", "Standard connector, no IT app approval.", "Decided"),
    ("D11", "Slack connection", "Power Automate Slack connector (OAuth) under a People Ops service account; posts go to channels.", "Answers the plan's open question: no bot token needed.", "Decided"),
    ("D12", "Sensitive types", "OFF-I, LOA and ER withhold the employee name from Slack and the digest. ER keeps only a summary on the list; details live in the restricted case file.", "Limit sensitive content (inspiration: restricted lane).", "Decided"),
    ("D13", "Approval rollout", "XPAY's HR approval goes live in Sprint 2 (pilot). JOB, COMP, PTO (and REQ) chains switch on in Sprint 3 by filling ApprovalChain after matrix sign-off.", "Matches the plan dependency 'Approval matrix signed off'.", "Decided"),
    ("D14", "SLA calendar", "Business days skip weekends, not holidays.", "Holiday table is on the Run backlog.", "Accepted"),
    ("D15", "Sensitive data in Description", "The form tells people not to type bank numbers, SSNs or medical details. HR collects them securely.", "Description is free text and visible to the HR team.", "Decided"),
    ("D16", "Plan typo", "'deparHRent' → 'department' (a find/replace of 'tm' → 'HR'). Check the source doc for other damaged words.", "", "FYI"),
]

INSPIRATION = [
    ("Core ticket columns (reference, requester, category, priority, status, owner, received, first response, next action + due, closed, internal notes)", "Adopt", "Your direction — a lean, general queue.", "They are the list (List Fields tab). Plan fields kept: Employee Name/ID, Manager, Effective Date, SLA Due Date."),
    ("Message / conversation identifiers, last inbound date", "Defer", "Only meaningful with email intake.", "Add back if email intake is turned on (the earlier hr-ticketing-triage solution has the matching logic)."),
    ("Email-first intake", "Reject", "Breaks 'One front door' and 'Not in the intake, does not exist'.", "Adapt: HR mailbox auto-reply points to the form (Sprint 2)."),
    ("Don't give employees access to the whole HR queue", "Adopt", "Internal notes and other people's requests must stay private.", "Microsoft Forms front door; HR-only list (D3)."),
    ("One owner + Next action + Next action date on every active case", "Adopt", "'In progress' alone doesn't say what happens next.", "Flow 3 keeps Next Action = first open sub-step; HR can override."),
    ("Richer waiting statuses", "Adapt", "Keep Tag, You're It 0–3 for continuity.", "Waiting On field used with 2-Waiting."),
    ("First human response date", "Adopt", "Shows pick-up speed separately from completion.", "Stamped automatically (D9)."),
    ("Action-based workbench views", "Adopt", "Fast daily triage.", "8 views, incl. Needs Attention Today and Exceptions."),
    ("Daily digest instead of per-update emails", "Adopt", "Less noise; one place to see risk.", "Flow 6, sensitive names withheld."),
    ("Restricted lane for sensitive casework", "Adapt", "Most requests are routine.", "ER type with summary-only rule + restricted case file; Sensitive flag hides names."),
    ("Ticket Messages list; copying email bodies", "Reject", "One-list rule; the comment thread is the log.", ""),
    ("Send HR responses from list fields", "Reject", "Complexity for little gain.", "Close-the-loop email at resolution; replies in Slack/Outlook."),
    ("Auto-close after inactivity", "Reject", "LOA must stay open; closure is a deliberate HR act.", "Digest surfaces stale items."),
    ("Exception path — never lose a request", "Adopt", "A failed step must stay visible.", "Try/Catch in every flow → admin email + Automation Note + Exceptions view."),
    ("Loop guards", "Adopt", "Flows that update their own trigger item can loop.", "Trigger conditions (Flow 4) and compare-before-write (Flow 3)."),
    ("Don't promise resolution dates in acknowledgements", "Adopt", "Sets the right expectation.", "Confirmation email says 'target' and lists the info needed."),
    ("Pilot scenario list", "Adapt", "Email-specific cases don't apply.", "22-case Test Pack for HR request paths, approvals, permissions, failures."),
    ("Test access with a non-HR account", "Adopt", "Filtered views are not security.", "Test T21."),
]

TRACE = [
    ("Guiding principle — One front door", "Microsoft Form pinned in #lane-people-talent; HR mailbox auto-reply points to it", "S0, S2"),
    ("Guiding principle — Not in the intake, does not exist", "Every case is a list item; views and digest only show list items", "S2"),
    ("Guiding principle — Slack native first", "Updated (D3): Microsoft-native for the record, Slack for visibility", "S0"),
    ("Guiding principle — The list is the system of record", "ONE SharePoint list with core columns (D1, D2)", "S0"),
    ("Guiding principle — Templates do the thinking", "Checklists table → Flow 2 posts the sub-steps", "S1"),
    ("Guiding principle — The thread is the work log", "Item comments: intake, checklist, approvals, resolution", "S1"),
    ("Guiding principle — Offboarding is a path", "OFF-V / OFF-I are change types on the same list and engine", "S2"),
    ("Layer 1 — Intake + confirmation to submitter", "Flow 0 (form → item) + Flow 1 (confirmation email with 'info needed', People Ops ping)", "S1"),
    ("Layer 2 / Phase 1 — Thread-based checklist", "Flow 2 checklist comment", "S1"),
    ("Layer 2 / Phase 2 — True subtasks + progress", "Checklist field [ ]/[x] + Flow 3 Checklist % and Next Action", "S1"),
    ("Master template items never completed", "Checklists table in the reference workbook", "S0"),
    ("Intake fields — universal", "Form questions 1–6 + list core columns", "S0"),
    ("Intake fields — conditional by change type", "Description + the change type's InfoNeeded prompt (form subtitle, confirmation email, intake comment)", "S0, S1"),
    ("Change types A1–I", "Mapped into 16 general HR types (PlanRef column): A1→OFF-V, A2→OFF-I, B→JOB, C→COMP, D→PAY, E→DATA, F→PTO, G→LOA, H→BEN, I→XPAY", "S0"),
    ("List fields — Request ID … Completion date", "Ticket Reference, Category, Status, Assigned To, SLA Due Date, Closed Date, Internal Notes (Jira link)…", "S0"),
    ("Status labels 0–3", "Status choice (Statuses tab)", "S0"),
    ("SLA targets by change type", "Change Types table → Flow 1 SLA Due Date", "S1"),
    ("Completion auto-comment (✅ Resolved on [date] by [assignee] + stamp date)", "Flow 4 Close-out", "S1"),
    ("Flow 1 — Intake Notification", "Flow 1 (Flow 0 feeds it from the form)", "S1"),
    ("Flow 2 — Checklist Router (paths + Other catch-all)", "Flow 2 (one generic path reads the table; OTHER → 'needs manual scoping')", "S1"),
    ("Flow 3 — Payroll Notification (Template I, not a gate)", "Flow 4 Close-out, NotifyPayrollOnClose (D8)", "S2"),
    ("Template I — HR approval → Paycor earning → resolve", "XPAY: Flow 5 (HR stage) + checklist + Payroll FYI", "S2"),
    ("Template I — right pay period; flag after cutoff", "Flow 1 pay period + after-cutoff flag in the intake comment and Slack", "S2"),
    ("Template G — route to CJ, stays open", "LOA default assignee CJ; nothing auto-resolves; Next Action Due = expected return", "S2"),
    ("Approval routing per approval matrix", "Flow 5 + Approvers table; switched on after sign-off", "S3"),
    ("Nice to add — cutoff reminder", "Flow 6 cutoff reminder", "S4"),
    ("Phase 3 — Dashboard view", "Views + daily digest; Power BI later", "S4, S6"),
    ("Phase 3 — Manager self-service form", "Already delivered by the Microsoft Form front door; scope can widen", "S2"),
    ("Phase 3 — Bot token / in-Slack approvals; Paycor iPaaS", "Run backlog", "S6"),
    ("Build order — E first, then I, then G", "DATA end to end (S1), then XPAY pilot, then LOA (S2)", "S1, S2"),
    ("Open question — Slack app vs Power Automate connector", "Answered: connector (D11)", "S0"),
]

DEPENDENCIES = [
    ("HR Change Reference.xlsx stored on the People Ops site", "CJ", "Ready (this kit)", "S0", "Flows read it."),
    ("People Ops service account (SharePoint + Slack connections)", "CJ / IT", "To confirm", "S0", "Needed for the first-response stamp and for Slack posts."),
    ("Change types and form wording signed off", "CJ / Johanna", "To confirm", "S0", "Gate 0."),
    ("Approval matrix signed off by stakeholders", "CJ", "To confirm", "S3", "Switches on JOB, COMP, PTO, REQ approvals."),
    ("Jira connector access + HR Jira project", "IT", "To confirm", "S5", "Plan open question."),
    ("Shared offboarding email distro exists and is current", "CJ / Johanna", "To confirm", "S5", "Plan open question."),
    ("Word termination letter template", "CJ", "To build", "S5", "'Populate a Microsoft Word template' is a premium action."),
    ("Monday backfill request form URL", "Talent", "To confirm", "S5", ""),
    ("Payroll Monday board — column/item", "Sydney / Sandra", "To confirm", "S5", "Plan open question."),
    ("Paycor bot token / iPaaS", "IT / Paycor CSM", "Pending", "S6", "Run phase."),
]

# ------------------------------------------------------------------ roadmap
SPRINTS = [
    ("S0", "Foundations", 1, "Crawl", "Reference workbook, the one list, the form, views, connections.", "Plan approved; this kit downloaded.",
     "G0 — Ready to build", ["D3, D4, D5 confirmed by CJ", "Change types, InfoNeeded wording and checklists reviewed with Johanna",
                             "Reference workbook uploaded and Settings filled", "List built (core columns, choices, views); form built and not yet shared",
                             "T21 passed (a non-HR colleague can submit the form but can't open the list)"]),
    ("S1", "Crawl — DATA path end to end", 2, "Crawl", "Flow 0 Form Intake, Flow 1 Intake Notification, Flow 2 Checklist Router, Flow 3 Case Updates, Flow 4 Close-out. Proven on DATA.",
     "G0 passed.", "G1 — Plumbing proven", ["T01–T06 pass", "One live DATA request by Johanna closed end to end",
                                            "A failure alert received from a Catch block", "Flows owned by the service account, not a personal one"]),
    ("S2", "Crawl complete — XPAY pilot + front door live", 2, "Crawl", "Flow 5 Approval Routing (HR stage for XPAY), Payroll FYI, LOA/offboarding checks, front door launch.",
     "G1 passed.", "G2 — Front door live", ["T07–T13 pass", "First real XPAY case approved, earning entered, closed, Payroll FYI posted",
                                             "Form pinned in #lane-people-talent; announcement sent", "HR mailbox auto-reply points to the form"]),
    ("S3", "Walk — approval chains + UAT", 2, "Walk", "JOB, COMP, PTO (and REQ) chains switched on from Excel; UAT across every type.",
     "G2 passed; approval matrix signed off.", "G3 — Approvals live", ["Matrix signed off; Approvers table filled per region", "T14–T17 and T22 pass", "UAT with Johanna on every change type"]),
    ("S4", "Walk — daily digest, cutoff reminder, hardening", 2, "Walk", "Flow 6; exception drill; two weeks of steady operation.",
     "G3 passed.", "G4 — Operational", ["T18–T20 pass", "Two weeks of digests reviewed", "No request handled outside the list in those two weeks",
                                        "First-response and SLA baseline reviewed (views)"]),
    ("S5", "Walk — downstream actions (dependency-driven)", 2, "Walk", "Jira, offboarding distro, Word letters, Monday board — each when its dependency is confirmed.",
     "G4 passed; dependency confirmed.", "G5 — per item", ["Dependency confirmed in writing", "Built and tested on a test request", "Checklist step marked automated (Auto column)"]),
    ("S6", "Run — full audit + integrations", 0, "Run", "Dashboard, in-Slack approvals, Paycor iPaaS, holiday-aware SLA.",
     "G5 items as available.", "Ongoing", ["Each item ships with its own tests and a rollback note"]),
]

# ------------------------------------------------------------------ build steps
STEPS = []


def step(id, area, title, where="", do=(), paste=(), check="", owner="Johanna"):
    STEPS.append(dict(id=id, sprint=id.split("-")[0], area=area, title=title, where=where, do=list(do), paste=list(paste), check=check, owner=owner))


OPENING = [
    ("Settings row  (Excel Online (Business) → List rows present in a table)", f"Location: People Ops site · Document Library: Documents · File: HR Change/{REF_FILE} · Table: tblSettings"),
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
DISPLAY_NAME = ("Display Name (Compose)", "if(equals(outputs('Config')?['Sensitive'], 'Yes'), '(confidential — see list)', triggerBody()?['EmployeeName'])")

# ================= S0
step("S0-01", "Decisions", "Confirm the open design decisions", "Decisions tab",
     ["Review D3 (Microsoft Form front door, HR-only list), D4 (general change types), D5 (XPAY: Payroll verifier, not gate).", "Mark each Confirmed or Changed."],
     check="D3, D4, D5 no longer say 'Confirm'.", owner="CJ")
step("S0-02", "Excel", "Store the reference workbook", "People Ops SharePoint site → Documents",
     [f"Create a folder 'HR Change' and upload {REF_FILE}. Folder → Manage access: HR owners edit; nobody else needs access.",
      "Read the About sheet: never rename a table (tblSettings, tblChangeTypes, tblChecklist, tblApprovers) or a column header."],
     check="The file opens in Excel for the web; Table Design shows the four table names.", owner="CJ")
step("S0-03", "Excel", "Fill in Settings", f"{REF_FILE} → Settings",
     ["Set TimeZone, CutoffDay1/2 (real payroll cutoffs), ReminderBusinessDays, ServiceAccountEmail, AdminEmail. FormUrl comes in S0-10."],
     check="No '<' placeholders left except FormUrl.")
step("S0-04", "Excel", "Review Change Types with Johanna", f"{REF_FILE} → Change Types",
     ["Check every type's Group, DefaultPriority, SLA, DueDateRule and InfoNeeded wording (it goes into the form subtitle and the confirmation email).",
      "Retire a type by deleting its row and its choice; add one by adding a row (Code = the text before ' — ').",
      "Leave ApprovalChain blank for JOB, COMP, PTO, REQ until Sprint 3 (D13). XPAY = 'HR' (pilot).",
      "Set DefaultAssigneeEmail where a type always goes to the same person (LOA and ER → CJ are pre-filled)."],
     check="Every row has Code, ChangeType, DueDateRule; Yes/No columns contain only Yes or No.")
step("S0-05", "Excel", "Review the sub-step checklists", f"{REF_FILE} → Checklists",
     ["Walk each type's steps. Keep Step as two digits (01, 02 …). Auto: 'approval' / 'resolution' steps tick themselves. AfterApproval = Yes shows 🔒."],
     check="Johanna signs off each type's steps.")
step("S0-06", "Excel", "Draft the Approvers table", f"{REF_FILE} → Approvers",
     ["Rows for Regional / HR / Payroll by Region ('*' = any region). Region values must match the list's Region choices."],
     check="An active HR row with Region '*' exists (needed for the XPAY pilot).", owner="CJ")
step("S0-07", "SharePoint", "Create the list from CSV", "Site → + New → List → From CSV → HR-Change-Requests.csv",
     ["Keep the headers exactly (no spaces = the internal names the flows use). Set types per the List Fields tab.",
      f"Name it '{LIST_NAME}' → Create. Delete the SAMPLE row.",
      "Add the Person column AssignedTo (type the name without spaces; rename the display name later if you like).",
      "Paste the full choice lists into ChangeType, Priority, Status, Region, WaitingOn, ApprovalStatus, Outcome. Defaults: Status 0-New, ApprovalStatus Not required.",
      "Rename display names to the List Fields tab's 'Display name' (e.g. ChangeType → Category, RequestId → Ticket Reference)."],
     check="List settings → Columns: internal names have no spaces (hover a column: Field=EmployeeName).")
step("S0-08", "SharePoint", "Make the list HR-only", "List settings → Permissions for this list",
     ["Stop inheriting permissions. Keep HR team (Edit or above) and the service account; remove everyone else.",
      "Nobody else needs access: requesters use the form and get emails."], check="Test T21.", owner="CJ")
step("S0-09", "SharePoint", "Create the workbench views", "List → All Items → Create new view",
     ["Create the 8 views on the Views tab."], check="'Needs Attention Today' and 'Exceptions' exist.")
step("S0-10", "Forms", f"Build the '{FORM_NAME}' form", "forms.office.com → New Form",
     ["Settings: 'Only people in my organization can respond' + 'Record name'.",
      "Add the 8 questions from the Form Questions tab, in order. Question 1's choices = the Category values exactly (copy from Change Types → ChangeType).",
      "Details subtitle: paste the wording from the Form Questions tab. Optionally list each type's InfoNeeded under it.",
      "Copy the form link (Collect responses → Copy link) into Settings → FormUrl. Don't share it yet."],
     check="Preview: all 8 questions; question 1 shows all 17 categories.")
step("S0-11", "Power Automate", "Create the connections", "make.powerautomate.com → Data → Connections",
     ["Sign in as the People Ops service account for SharePoint, Excel Online (Business), Office 365 Outlook, Office 365 Users, Microsoft Forms, Approvals and Slack (D11).",
      "Make sure that account owns the form (or is a form co-owner) so Flow 0 can read responses."],
     check="Seven connections show 'Connected' under the service account.", owner="CJ")

# ================= S1
step("S1-01", "Flow 0", "Create 'HR Change - 0 Form Intake'", "My flows → + New flow → Automated cloud flow",
     [f"Trigger: Microsoft Forms — When a new response is submitted · Form: {FORM_NAME}.",
      "Add Scope 'Try'. Inside: Microsoft Forms → Get response details 'Response' (Form + Response Id from the trigger); Office 365 Users → Get user profile (V2) 'Requester' (User = Responders' Email).",
      f"SharePoint → Create item 'Create request' in {LIST_NAME} with the fields below (pick each answer from 'Response' in Dynamic content).",
      "Below Try: the Catch pattern (no 'Flag the item' — there is no item yet)."],
     paste=[("Title", "Dynamic content 'What do you need?', then type ' — ', then 'Employee name' (Flow 1 rewrites it anyway)"),
            ("ChangeType Value", "answer: What do you need?"), ("EmployeeName / EmployeeId / ManagerEmail", "answers 2, 3, 4"),
            ("Region Value", "answer 5"), ("EffectiveDate", "answer 6"),
            ("Priority Value (fx) — blank when unanswered; inside fx, click answer 7 in Dynamic content where it says <answer 7>", "if(empty(<answer 7>), null, <answer 7>)"),
            ("Description", "answer 8"),
            ("RequesterEmail", "Responders' Email (from Response)"), ("RequesterName", "Display Name (from Requester)")],
     check="Submit the form once: a new list item appears with these fields.")
step("S1-02", "Flow 1", "Create 'HR Change - 1 Intake Notification'", "Automated cloud flow",
     [f"Trigger: SharePoint — When an item is created · List: {LIST_NAME}. (Runs for form requests AND items HR adds by hand.)",
      "Add Scope 'Try' and the Standard opening block inside it (Appendix A)."], paste=OPENING,
     check="Run once with a DATA request: Config shows the DATA row.")
step("S1-03", "Flow 1", "SLA due date", "Inside Try, after Request ID", ["Four Compose actions."],
     paste=[("Submitted Local", "convertFromUtc(triggerBody()?['Created'], outputs('Settings')?['TimeZone'])"),
            ("SLA Start  (weekend submissions start Monday)", "addDays(startOfDay(outputs('Submitted_Local')), if(equals(dayOfWeek(outputs('Submitted_Local')), 6), 2, if(equals(dayOfWeek(outputs('Submitted_Local')), 0), 1, 0)))"),
            ("SLA Days", "if(empty(outputs('Config')?['SLABusinessDays']), 0, int(outputs('Config')?['SLABusinessDays']))"),
            ("SLA Due Date", "if(equals(outputs('Config')?['DueDateRule'], 'Effective date'), formatDateTime(triggerBody()?['EffectiveDate'], 'yyyy-MM-dd'), if(equals(outputs('SLA_Days'), 0), null, formatDateTime(addDays(outputs('SLA_Start'), add(add(mul(div(outputs('SLA_Days'), 5), 7), mod(outputs('SLA_Days'), 5)), if(greater(add(dayOfWeek(outputs('SLA_Start')), mod(outputs('SLA_Days'), 5)), 5), 2, 0))), 'yyyy-MM-dd')))")],
     check="DATA submitted Monday → SLA Due Date = Tuesday.")
step("S1-04", "Flow 1", "Pay period + cutoff flag (XPAY)", "Inside Try", ["Three Compose actions. Effective Date = first date worked for XPAY."],
     paste=[("Pay Period", "concat(formatDateTime(triggerBody()?['EffectiveDate'], 'yyyy-MM'), if(lessOrEquals(dayOfMonth(triggerBody()?['EffectiveDate']), 15), ' (1-15)', ' (16-EOM)'))"),
            ("Cutoff Date", "addDays(startOfMonth(triggerBody()?['EffectiveDate']), sub(int(if(lessOrEquals(dayOfMonth(triggerBody()?['EffectiveDate']), 15), outputs('Settings')?['CutoffDay1'], outputs('Settings')?['CutoffDay2'])), 1))"),
            ("After Cutoff", "and(equals(outputs('Code'), 'XPAY'), greater(ticks(outputs('Submitted_Local')), ticks(addDays(outputs('Cutoff_Date'), 1))))")],
     check="For DATA: After Cutoff = false.")
step("S1-05", "Flow 1", "Write the core fields back", "Inside Try → SharePoint → Update item",
     [f"List: {LIST_NAME} · Id: ID (trigger). Required columns it asks for: pick the same field from the trigger.",
      "Don't set Status or ApprovalStatus here — column defaults cover intake; Flow 5 owns approvals."],
     paste=[("Title (fx)", "concat(triggerBody()?['ChangeType']?['Value'], ' — ', triggerBody()?['EmployeeName'])"),
            ("RequestId (fx)", "outputs('Request_ID')"),
            ("Priority Value (fx)", "coalesce(triggerBody()?['Priority']?['Value'], outputs('Config')?['DefaultPriority'], 'Routine')"),
            ("ReceivedDate (fx)", "triggerBody()?['Created']"),
            ("SLADueDate (fx) / NextActionDue (fx)", "outputs('SLA_Due_Date')   ·   outputs('SLA_Due_Date')"),
            ("RequesterEmail (fx)  — keeps form value; falls back to whoever created the item", "coalesce(triggerBody()?['RequesterEmail'], triggerBody()?['Author']?['Email'])"),
            ("RequesterName (fx)", "coalesce(triggerBody()?['RequesterName'], triggerBody()?['Author']?['DisplayName'])")],
     check="Item shows Ticket Reference HRC-00000N, Priority, Received Date, SLA Due Date.")
step("S1-06", "Flow 1", "Default assignee", "Inside Try → Condition 'Default assignee?'",
     ["Left (fx) and(empty(triggerBody()?['AssignedTo']), not(empty(coalesce(outputs('Config')?['DefaultAssigneeEmail'], '')))) · is equal to · (fx) true",
      "True → Update item 'Set default assignee' (required columns from trigger)."],
     paste=[("AssignedTo Claims (fx)", "outputs('Config')?['DefaultAssigneeEmail']")], check="An LOA request gets CJ as Assigned To.")
step("S1-07", "Flow 1", "Intake comment", "Inside Try — Comment pattern (Appendix B) named 'Intake'", ["Add 'Display Name' first (Slack uses it too)."],
     paste=[DISPLAY_NAME,
            ("Intake text (Compose)", "📥 @{outputs('Request_ID')} received from @{coalesce(triggerBody()?['RequesterName'], triggerBody()?['Author']?['DisplayName'])}. Priority: @{coalesce(triggerBody()?['Priority']?['Value'], outputs('Config')?['DefaultPriority'])}. SLA target: @{coalesce(outputs('SLA_Due_Date'), outputs('Config')?['SLATarget'])}.\nInfo needed for this type: @{outputs('Config')?['InfoNeeded']}@{if(equals(outputs('Code'), 'XPAY'), concat(decodeUriComponent('%0A'), 'Pay period: ', outputs('Pay_Period'), if(outputs('After_Cutoff'), ' ⚠️ submitted after the payroll cutoff', '')), '')}"),
            ("Post intake — Body (fx)", "setProperty(json('{}'), 'text', take(outputs('Intake_text'), 1900))")],
     check="Comment in the item's Comments pane.")
step("S1-08", "Flow 1", "Ping the People Ops room", "Inside Try → Slack → Post message (V2)", ["Channel: People Ops room."],
     paste=[("Message text", ":inbox_tray: *New HR request <@{triggerBody()?['{Link}']}|@{outputs('Request_ID')}>* — @{coalesce(triggerBody()?['Priority']?['Value'], outputs('Config')?['DefaultPriority'])}\n@{triggerBody()?['ChangeType']?['Value']} — @{outputs('Display_Name')}\nEffective @{formatDateTime(triggerBody()?['EffectiveDate'], 'MMM d, yyyy')} • SLA target @{coalesce(outputs('SLA_Due_Date'), outputs('Config')?['SLATarget'])}@{if(outputs('After_Cutoff'), ' • :warning: after payroll cutoff', '')}")],
     check="Posts; for OFF-I / LOA / ER the name shows '(confidential — see list)'.")
step("S1-09", "Flow 1", "Confirmation to the requester", "Inside Try → Office 365 Outlook → Send an email (V2)",
     ["Lists what HR needs for this request type, without promising a date."],
     paste=[("To (fx)", "coalesce(triggerBody()?['RequesterEmail'], triggerBody()?['Author']?['Email'])"),
            ("Subject", "We've got your request: @{outputs('Request_ID')}"),
            ("Body", "Hi @{coalesce(triggerBody()?['RequesterName'], triggerBody()?['Author']?['DisplayName'])},\n\nThanks for using the HR front door. Your request @{outputs('Request_ID')} (@{triggerBody()?['ChangeType']?['Value']}) is in the People Ops queue.\n\nTo finish it we'll need: @{outputs('Config')?['InfoNeeded']}. If anything's missing from what you sent, just reply to this email.\n\nTarget: @{coalesce(outputs('SLA_Due_Date'), outputs('Config')?['SLATarget'])}. This is an automatic confirmation — a team member will pick it up and keep you posted.\n\nPeople Ops")],
     check="Email arrives with the InfoNeeded line.")
step("S1-10", "Flow 1", "Catch block + test", "Below Try", ["Catch pattern (Appendix C). Run T01 and T02."], paste=CATCH, check="T01, T02 = Pass.")
step("S1-11", "Flow 2", "Create 'HR Change - 2 Checklist Router'", "Automated cloud flow",
     [f"Trigger: When an item is created · List: {LIST_NAME}. Separate from Flow 1 on purpose: one failing can't stop the other.",
      "Scope 'Try' + Standard opening block."], paste=OPENING, check="Saved.")
step("S1-12", "Flow 2", "Read this type's sub-steps", "Inside Try", ["Excel List rows 'Checklist rows' (tblChecklist); Filter array 'Active steps'."],
     paste=[("Checklist rows — Filter Query / Order By", "Code eq '@{outputs('Code')}'   ·   Step"),
            ("Active steps — From (fx) / advanced", "outputs('Checklist_rows')?['body/value']   ·   @equals(item()?['Active'], 'Yes')")],
     check="For DATA: 5 rows.")
step("S1-13", "Flow 2", "Write and post the checklist", "Inside Try → Condition 'Has checklist'",
     ["Left (fx) length(body('Active_steps')) · is greater than · 0. True branch: two Selects (Map in text mode), one Compose, Update item, Comment pattern 'Checklist comment'."],
     paste=[("Checklist lines — From / Map", "body('Active_steps')   ·   concat('[ ] ', item()?['Step'], '. ', item()?['Task'], if(equals(item()?['AfterApproval'], 'Yes'), ' 🔒', ''), if(equals(item()?['DueOnLastDay'], 'Yes'), concat(' 📅 ', formatDateTime(triggerBody()?['EffectiveDate'], 'MMM d')), ''), if(empty(item()?['Auto']), '', concat(' ⚙ ', item()?['Auto'])))"),
            ("Comment lines — From / Map", "body('Active_steps')   ·   concat('☐ ', item()?['Step'], '. ', item()?['Task'], if(empty(item()?['Details']), '', concat(' — ', item()?['Details'])))"),
            ("Checklist text (Compose)", f"join(body('Checklist_lines'), {LF})"),
            ("Save checklist (Update item) — Checklist / ChecklistProgress / NextAction", "outputs('Checklist_text')   ·   0   ·   concat(first(body('Active_steps'))?['Step'], '. ', first(body('Active_steps'))?['Task'])"),
            ("Checklist comment text (Compose)", "📋 @{outputs('Code')} sub-steps — @{length(body('Active_steps'))}. Tick them in the Checklist field by changing [ ] to [x].\n@{join(body('Comment_lines'), decodeUriComponent('%0A'))}"),
            ("Post checklist comment — Body (fx)", "setProperty(json('{}'), 'text', take(outputs('Checklist_comment_text'), 1900))")],
     check="DATA item: 5 '[ ]' lines; Next Action = '01. …'; checklist comment posted.")
step("S1-14", "Flow 2", "Catch-all for OTHER", "Condition 'Has checklist' → False", ["Comment pattern 'No template' + Slack post. Then Catch block below Try. Run T03, T04."],
     paste=[("No template text (Compose)", "⚠️ No sub-steps for @{outputs('Code')} — needs manual scoping. Assign it and write the steps into the Checklist field ([ ] one per line)."),
            ("Slack message", ":warning: <@{triggerBody()?['{Link}']}|@{outputs('Request_ID')}> (@{triggerBody()?['ChangeType']?['Value']}) needs manual scoping.")] + CATCH,
     check="T03, T04 = Pass.")
step("S1-15", "Flow 3", "Create 'HR Change - 3 Case Updates'", "Automated cloud flow",
     [f"Trigger: When an item is created or modified · List: {LIST_NAME}. Trigger Settings → Concurrency control ON, degree 1.",
      "Purpose: stamps First Human Response, keeps Checklist % and Next Action in step with the ticks. Compare-before-write ends its own loop after one extra run.",
      "Scope 'Try': Settings row + Settings, then the composes below."],
     paste=[("First response needed (Compose)", "and(empty(triggerBody()?['FirstResponseDate']), not(equals(triggerBody()?['Status']?['Value'], '0-New')), not(equals(toLower(triggerBody()?['Editor']?['Email']), toLower(outputs('Settings')?['ServiceAccountEmail']))))"),
            ("Lines (Compose)", f"split(replace(coalesce(triggerBody()?['Checklist'], ''), {CR}, ''), {LF})"),
            ("Open steps / Done steps (Filter arrays on Lines)", "@startsWith(trim(item()), '[ ]')   ·   @startsWith(toLower(trim(item())), '[x]')"),
            ("Progress (Compose)", "if(equals(add(length(body('Open_steps')), length(body('Done_steps'))), 0), 0, div(mul(length(body('Done_steps')), 100), add(length(body('Open_steps')), length(body('Done_steps')))))"),
            ("Progress changed (Compose)", "not(equals(float(outputs('Progress')), float(coalesce(triggerBody()?['ChecklistProgress'], 0))))"),
            ("Checklist next (Compose)", "if(empty(body('Open_steps')), '', take(trim(substring(trim(first(body('Open_steps'))), 3)), 250))"),
            ("Next action (Compose) — HR's own text sticks until the next tick", "if(or(outputs('Progress_changed'), empty(triggerBody()?['NextAction'])), outputs('Checklist_next'), triggerBody()?['NextAction'])")],
     check="Saved.")
step("S1-16", "Flow 3", "Write only when something changed", "Inside Try → Condition 'Changed'",
     ["Left (fx) below · is equal to · (fx) true → Update item (required columns from trigger) → Condition 'Just finished' → Comment pattern 'Ready' + Slack. Then Catch."],
     paste=[("Changed — Left", "or(outputs('First_response_needed'), outputs('Progress_changed'), not(equals(outputs('Next_action'), coalesce(triggerBody()?['NextAction'], ''))))"),
            ("Update — FirstResponseDate (fx)", "if(outputs('First_response_needed'), utcNow(), triggerBody()?['FirstResponseDate'])"),
            ("Update — ChecklistProgress / NextAction", "outputs('Progress')   ·   outputs('Next_action')"),
            ("Just finished — Left (· is equal to · true)", "and(outputs('Progress_changed'), equals(outputs('Progress'), 100), not(equals(triggerBody()?['Status']?['Value'], '3-Resolved')))"),
            ("Ready text (Compose)", "☑️ All sub-steps are done — set Status to 3-Resolved to close (LOA: only once the return is confirmed)."),
            ("Slack message", ":ballot_box_with_check: All steps done on <@{triggerBody()?['{Link}']}|@{triggerBody()?['RequestId']}> — @{coalesce(triggerBody()?['AssignedTo']?['DisplayName'], 'owner')}, ready to resolve.")] + CATCH,
     check="T17 passes.")
step("S1-17", "Flow 4", "Create 'HR Change - 4 Close-out'", "Automated cloud flow",
     [f"Trigger: When an item is created or modified · List: {LIST_NAME}. Trigger conditions (2 lines): fires once per closure — the flow stamps Closed Date, so its own update can't re-fire it.",
      "Scope 'Try' + opening block (skip Request ID)."],
     paste=[("Trigger condition 1", "@equals(triggerBody()?['Status']?['Value'], '3-Resolved')"), ("Trigger condition 2", "@empty(triggerBody()?['ClosedDate'])")] + OPENING[:5],
     check="Saved.")
step("S1-18", "Flow 4", "Outcome, resolver and ticked checklist", "Inside Try",
     ["Composes, one Select, one Filter array."],
     paste=[("Outcome / Resolver", "coalesce(triggerBody()?['Outcome']?['Value'], 'Completed')   ·   coalesce(triggerBody()?['AssignedTo']?['DisplayName'], triggerBody()?['Editor']?['DisplayName'])"),
            ("Today", "convertFromUtc(utcNow(), outputs('Settings')?['TimeZone'])"),
            ("Checklist before", f"split(replace(coalesce(triggerBody()?['Checklist'], ''), {CR}, ''), {LF})"),
            ("Checklist after — From / Map", "outputs('Checklist_before')   ·   if(and(startsWith(item(), '[ ]'), or(contains(item(), '⚙ resolution'), not(equals(outputs('Outcome'), 'Completed')))), concat(if(equals(outputs('Outcome'), 'Completed'), '[x]', '[–]'), substring(item(), 3)), item())"),
            ("Still open — From / advanced", "body('Checklist_after')   ·   @startsWith(item(), '[ ]')")],
     check="On a DATA close: steps 03–05 become [x].")
step("S1-19", "Flow 4", "Stamp Closed Date and post ✅", "Inside Try", ["Update item 'Stamp close' (required columns from trigger), then Comment pattern 'Resolved'."],
     paste=[("ClosedDate / Outcome Value", "formatDateTime(outputs('Today'), 'yyyy-MM-dd')   ·   outputs('Outcome')"),
            ("Checklist (fx)", f"if(empty(triggerBody()?['Checklist']), '', join(body('Checklist_after'), {LF}))"),
            ("NextAction", "''"),
            ("Resolved text (Compose)", "✅ Resolved on @{formatDateTime(outputs('Today'), 'MMM d, yyyy')} by @{outputs('Resolver')}@{if(equals(outputs('Outcome'), 'Completed'), '', concat(' (Outcome: ', outputs('Outcome'), ')'))}@{if(greater(length(body('Still_open')), 0), concat(' — note: ', string(length(body('Still_open'))), ' sub-step(s) were still open'), '')}")],
     check="Comment reads '✅ Resolved on <date> by <assignee>'.")
step("S1-20", "Flow 4", "Close-the-loop + People Ops summary", "Inside Try",
     ["Condition 'Close the loop?' → Send an email (V2). Condition 'Summary?' → Slack People Ops room."],
     paste=[("Close the loop? — Left (· is equal to · true)", "or(equals(outputs('Config')?['CloseLoopEmail'], 'Yes'), not(equals(outputs('Outcome'), 'Completed')))"),
            ("Email To / Subject", "coalesce(triggerBody()?['RequesterEmail'], triggerBody()?['Author']?['Email'])   ·   Update on your HR request @{triggerBody()?['RequestId']}"),
            ("Email Body", "Hi @{triggerBody()?['RequesterName']},\n\n@{if(equals(outputs('Outcome'), 'Completed'), concat('Good news — your request ', triggerBody()?['RequestId'], ' (', triggerBody()?['ChangeType']?['Value'], ') is complete.'), concat('Your request ', triggerBody()?['RequestId'], ' (', triggerBody()?['ChangeType']?['Value'], ') was closed as ', outputs('Outcome'), '. Your HR contact will follow up with the details.'))}\n\nQuestions? Reply to this email or ask in #lane-people-talent.\n\nPeople Ops"),
            ("Summary? — Left", "equals(outputs('Config')?['PostSummaryToPeopleOps'], 'Yes')"),
            ("Slack summary", ":white_check_mark: *<@{triggerBody()?['{Link}']}|@{triggerBody()?['RequestId']}>* @{outputs('Code')} closed by @{outputs('Resolver')} (@{outputs('Outcome')}).@{if(greater(length(body('Still_open')), 0), concat(' :warning: ', string(length(body('Still_open'))), ' step(s) still open.'), '')}")],
     check="T05 passes.")
step("S1-21", "Flow 4", "Payroll FYI (plan Flow 3) + Catch", "Inside Try → Condition 'Payroll FYI?'",
     ["Left (fx) and(equals(outputs('Config')?['NotifyPayrollOnClose'], 'Yes'), equals(outputs('Outcome'), 'Completed')) · is equal to · true → Slack post to the Payroll room. One-way — never waits for Payroll (D5).",
      "Below Try: Catch pattern. (The Payroll part is tested in Sprint 2, T10.)"],
     paste=[("Slack message (Payroll room)", ":moneybag: *Extra pay submitted in Paycor* — <@{triggerBody()?['{Link}']}|@{triggerBody()?['RequestId']}>\nEmployee: @{triggerBody()?['EmployeeName']} (@{triggerBody()?['EmployeeId']}) • First date worked: @{formatDateTime(triggerBody()?['EffectiveDate'], 'MMM d, yyyy')} • Pay period: @{concat(formatDateTime(triggerBody()?['EffectiveDate'], 'yyyy-MM'), if(lessOrEquals(dayOfMonth(triggerBody()?['EffectiveDate']), 15), ' (1-15)', ' (16-EOM)'))}\nDetails: @{take(coalesce(triggerBody()?['Description'], ''), 600)}\n_FYI — please verify the earning in Paycor. No response needed._")] + CATCH,
     check="Saved.")
step("S1-22", "Test", "Run the Sprint 1 tests", "Test Pack", ["T05, T06 (and re-run T01–T04 if anything changed)."], check="Gate G1 checklist complete.")

# ================= S2
step("S2-01", "Flow 5", "Create 'HR Change - 5 Approval Routing'", "Automated cloud flow",
     [f"Trigger: When an item is created · List: {LIST_NAME}.",
      "BEFORE Try (variables can't live in a scope): Initialize variable 'Outcome' (String) = Approved; 'History' (String) = empty.",
      "Scope 'Try' + opening block, then Filter array 'Stages' and Condition 'Needs approval' (Left fx length(body('Stages')) · is greater than · 0). Everything below goes in its True branch."],
     paste=OPENING + [("Stages — From (fx) / advanced", "split(replace(coalesce(outputs('Config')?['ApprovalChain'], ''), ' ', ''), '>')   ·   @not(empty(item()))")],
     check="XPAY → Stages = [\"HR\"]; DATA → empty, nothing happens.")
step("S2-02", "Flow 5", "Mark waiting and loop the stages", "Needs approval → True",
     ["Excel List rows 'Approvers rows' (tblApprovers). Update item 'Mark pending'. Apply to each 'Each stage' over body('Stages') — ⚠️ Settings → Concurrency ON, degree 1.",
      "Inside the loop: Condition 'Still approved' (variables('Outcome') is equal to Approved); S2-03 and S2-04 go in its True branch."],
     paste=[("Mark pending — ApprovalStatus / Status / WaitingOn", "Pending   ·   2-Waiting   ·   Approver")], check="Item shows 2-Waiting / Pending / Approver.")
step("S2-03", "Flow 5", "Find the approver", "Still approved → True", ["Filter arrays 'This region' / 'Any region' (From fx outputs('Approvers_rows')?['body/value']), Compose 'Approver'."],
     paste=[("This region — advanced", "@and(equals(item()?['Role'], items('Each_stage')), equals(item()?['Active'], 'Yes'), equals(toLower(trim(item()?['Region'])), toLower(trim(coalesce(triggerBody()?['Region']?['Value'], '')))))"),
            ("Any region — advanced", "@and(equals(item()?['Role'], items('Each_stage')), equals(item()?['Active'], 'Yes'), equals(trim(item()?['Region']), '*'))"),
            ("Approver (Compose)", "if(equals(items('Each_stage'), 'Manager'), coalesce(triggerBody()?['ManagerEmail'], ''), coalesce(first(body('This_region'))?['ApproverEmail'], first(body('Any_region'))?['ApproverEmail'], ''))")],
     check="XPAY: Approver = the HR row's email.")
step("S2-04", "Flow 5", "Ask, record, stop on reject", "Condition 'Approver found' (Left fx empty(outputs('Approver')) · is equal to · false)",
     ["True: Update item 'Show approver'; Approvals → Start and wait for an approval 'Approval' (Approve/Reject – First to respond); Append to string variable 'History'; Comment pattern 'Decision'; Condition 'Rejected?' (outputs('Approval')?['body/outcome'] is equal to Reject) → Set variable Outcome = Rejected.",
      "False: Set variable Outcome = NoApprover; Slack alert to the People Ops room."],
     paste=[("Show approver — CurrentApprover", "@{items('Each_stage')}: @{outputs('Approver')}"),
            ("Approval — Title / Assigned to", "@{items('Each_stage')} approval: @{triggerBody()?['ChangeType']?['Value']} — @{triggerBody()?['EmployeeName']} [@{outputs('Request_ID')}]   ·   outputs('Approver')"),
            ("Approval — Details", "**Employee:** @{triggerBody()?['EmployeeName']} (@{triggerBody()?['EmployeeId']})\n**Request:** @{triggerBody()?['ChangeType']?['Value']} • effective @{formatDateTime(triggerBody()?['EffectiveDate'], 'MMM d, yyyy')}\n**Requested by:** @{triggerBody()?['RequesterName']}\n\n@{triggerBody()?['Description']}"),
            ("Approval — Item link / description", "triggerBody()?['{Link}']   ·   Open @{outputs('Request_ID')}"),
            ("History — value (end with Enter)", "@{formatDateTime(convertFromUtc(utcNow(), outputs('Settings')?['TimeZone']), 'yyyy-MM-dd')} @{items('Each_stage')} @{if(equals(outputs('Approval')?['body/outcome'], 'Approve'), 'approved', 'rejected')} by @{first(outputs('Approval')?['body/responses'])?['responder']?['displayName']}@{if(empty(first(outputs('Approval')?['body/responses'])?['comments']), '', concat(' — ', first(outputs('Approval')?['body/responses'])?['comments']))}"),
            ("Decision text (Compose)", "@{if(equals(outputs('Approval')?['body/outcome'], 'Approve'), '👍', '👎')} @{items('Each_stage')} @{if(equals(outputs('Approval')?['body/outcome'], 'Approve'), 'approved', 'rejected')} by @{first(outputs('Approval')?['body/responses'])?['responder']?['displayName']}"),
            ("No approver — Slack message", ":rotating_light: <@{triggerBody()?['{Link}']}|@{outputs('Request_ID')}> needs a *@{items('Each_stage')}* approver but none was found (Manager email empty, or no active Approvers row for this role/region).")],
     check="Approval email arrives; decision comment appears.")
step("S2-05", "Flow 5", "Save the result after the loop", "Needs approval → True, AFTER the loop",
     ["SharePoint Get item 'Request now'. Select 'Approval ticks'. Update item 'Save approval result' (required columns from Request now). Comment pattern 'Result'. Then Catch below Try."],
     paste=[("Approval ticks — From / Map", f"split(replace(coalesce(outputs('Request_now')?['body/Checklist'], ''), {CR}, ''), {LF})   ·   if(and(startsWith(item(), '[ ]'), contains(item(), '⚙ approval')), concat('[x]', substring(item(), 3)), item())"),
            ("ApprovalStatus Value", "if(equals(variables('Outcome'), 'Approved'), 'Approved', if(equals(variables('Outcome'), 'Rejected'), 'Rejected', 'Needs approver'))"),
            ("Status Value", "if(equals(variables('Outcome'), 'Rejected'), '3-Resolved', if(equals(variables('Outcome'), 'NoApprover'), '2-Waiting', if(empty(outputs('Request_now')?['body/AssignedTo']), '0-New', '1-Assigned')))"),
            ("Outcome Value / WaitingOn Value", "if(equals(variables('Outcome'), 'Rejected'), 'Rejected', outputs('Request_now')?['body/Outcome']?['Value'])   ·   if(equals(variables('Outcome'), 'NoApprover'), 'Approver', null)"),
            ("CurrentApprover / InternalNotes", f"''   ·   concat(coalesce(outputs('Request_now')?['body/InternalNotes'], ''), {LF}, variables('History'))"),
            ("Checklist", f"if(or(equals(variables('Outcome'), 'NoApprover'), empty(outputs('Request_now')?['body/Checklist'])), outputs('Request_now')?['body/Checklist'], join(body('Approval_ticks'), {LF}))"),
            ("Result text (Compose)", "@{if(equals(variables('Outcome'), 'Approved'), concat('🔓 Approved (', join(body('Stages'), ' → '), '). 🔒 steps can start now.'), if(equals(variables('Outcome'), 'Rejected'), '⛔ Rejected — closed as Rejected; the requester will be told.', '⚠️ Approval stopped: no approver found. Fix the manager email or the Approvers table, approve offline and log it in Internal Notes.'))}")] + CATCH,
     check="T07: after approving, Approval Status = Approved and step 01 is [x].")
step("S2-06", "Test", "XPAY pilot tests", "Test Pack", ["Run T07, T08, T09, T10."], check="All pass.")
step("S2-07", "Test", "LOA, offboarding and ER paths", "Test Pack", ["Run T11, T12, T13."], check="All pass.")
step("S2-08", "Launch", "Open the front door", "Slack + Outlook",
     ["Pin the FormUrl in #lane-people-talent with the announcement (Appendix D).", "HR mailbox: automatic reply pointing to the form.",
      "Tell managers: 'Not in the intake, does not exist.'"], check="Pinned; auto-reply on; announcement sent.", owner="CJ")
step("S2-09", "Test", "First real XPAY case", "Live",
     ["Next real extra-shift request end to end: HR approval → Paycor single event earning → close → Payroll FYI."], check="Gate G2 checklist complete.")

# ================= S3
step("S3-01", "Decisions", "Approval matrix sign-off", "Stakeholders", ["Written sign-off on who approves JOB, COMP, PTO, REQ by region."], check="Dependency Confirmed.", owner="CJ")
step("S3-02", "Excel", "Switch on the approval chains", f"{REF_FILE} → Change Types + Approvers",
     ["ApprovalChain: JOB = 'Manager > Regional', COMP = 'Manager > Regional > HR', PTO = 'Manager' (REQ as signed off).", "Add Approvers rows per region."],
     check="Next JOB request starts a Manager approval (no flow edit).", owner="CJ")
step("S3-03", "Test", "Approval and config tests", "Test Pack", ["Run T14, T15, T16, T17, T22."], check="All pass.")
step("S3-04", "Test", "UAT with Johanna", "Every change type", ["Submit one of each type through the form; work and close each. Log issues in Notes."], check="Gate G3 checklist complete.")

# ================= S4
step("S4-01", "Flow 6", "Create 'HR Change - 6 Daily Digest + Cutoff Reminder'", "Scheduled cloud flow — every 1 day at 8:00 (your time zone)",
     ["Try: Settings row + Settings; Compose 'Today' and 'Today key'; Excel List rows 'Change types'; Filter 'Sensitive types'; Select 'Sensitive codes'; SharePoint Get items 'Open requests'."],
     paste=[("Today / Today key", "convertFromUtc(utcNow(), outputs('Settings')?['TimeZone'])   ·   formatDateTime(outputs('Today'), 'yyyyMMdd')"),
            ("Sensitive types — From / advanced", "outputs('Change_types')?['body/value']   ·   @equals(item()?['Sensitive'], 'Yes')"),
            ("Sensitive codes — From / Map (text mode)", "body('Sensitive_types')   ·   item()?['Code']"),
            ("Open requests — Filter Query / Order By / Top Count", "Status ne '3-Resolved'   ·   SLADueDate asc   ·   5000")],
     check="Run manually: Open requests returns your test items.")
step("S4-02", "Flow 6", "Digest sections", "Inside Try",
     ["Five Filter arrays named exactly as below (From fx outputs('Open_requests')?['body/value']), and one Select per filter named '<filter> lines' (e.g. 'SLA overdue lines', From fx body('SLA_overdue')) with the same Map. Then Condition 'Weekday with news' → Slack People Ops room."],
     paste=[("Unassigned", "@equals(item()?['Status']?['Value'], '0-New')"),
            ("SLA overdue", "@and(not(empty(item()?['SLADueDate'])), less(formatDateTime(item()?['SLADueDate'], 'yyyyMMdd'), outputs('Today_key')))"),
            ("Action due", "@and(not(empty(item()?['NextActionDue'])), lessOrEquals(formatDateTime(item()?['NextActionDue'], 'yyyyMMdd'), outputs('Today_key')))"),
            ("Waiting", "@equals(item()?['Status']?['Value'], '2-Waiting')"),
            ("Exceptions", "@not(empty(item()?['AutomationNote']))"),
            ("Line Map (text mode)", "concat('• <', item()?['{Link}'], '|', item()?['RequestId'], '> ', first(split(item()?['ChangeType']?['Value'], ' ')), ' — ', if(contains(body('Sensitive_codes'), first(split(item()?['ChangeType']?['Value'], ' '))), '(confidential)', item()?['EmployeeName']), ' — ', coalesce(item()?['AssignedTo']?['DisplayName'], 'unassigned'), if(empty(item()?['WaitingOn']?['Value']), '', concat(' — waiting on ', item()?['WaitingOn']?['Value'])), if(empty(item()?['NextAction']), '', concat(' — next: ', item()?['NextAction'])))"),
            ("Weekday with news — Left (· is equal to · true)", "and(greater(dayOfWeek(outputs('Today')), 0), less(dayOfWeek(outputs('Today')), 6), greater(add(add(add(length(body('Unassigned')), length(body('SLA_overdue'))), add(length(body('Action_due')), length(body('Exceptions')))), length(body('Waiting'))), 0))"),
            ("Slack message", ":bar_chart: *HR requests — daily check* (@{length(outputs('Open_requests')?['body/value'])} open)\n@{if(empty(body('Unassigned_lines')), '', concat('*:white_circle: 0-New, not picked up*', decodeUriComponent('%0A'), join(body('Unassigned_lines'), decodeUriComponent('%0A')), decodeUriComponent('%0A')))}@{if(empty(body('SLA_overdue_lines')), '', concat('*:red_circle: SLA overdue*', decodeUriComponent('%0A'), join(body('SLA_overdue_lines'), decodeUriComponent('%0A')), decodeUriComponent('%0A')))}@{if(empty(body('Action_due_lines')), '', concat('*:large_yellow_circle: Next action due*', decodeUriComponent('%0A'), join(body('Action_due_lines'), decodeUriComponent('%0A')), decodeUriComponent('%0A')))}@{if(empty(body('Waiting_lines')), '', concat('*:hourglass: Waiting*', decodeUriComponent('%0A'), join(body('Waiting_lines'), decodeUriComponent('%0A')), decodeUriComponent('%0A')))}@{if(empty(body('Exceptions_lines')), '', concat('*:warning: Automation exceptions*', decodeUriComponent('%0A'), join(body('Exceptions_lines'), decodeUriComponent('%0A'))))}")],
     check="T18 passes.")
step("S4-03", "Flow 6", "Payroll cutoff reminder", "Inside Try",
     ["Compose 'Cutoff days'; Select 'This month cutoffs' / 'Next month cutoffs' (Map text mode); Filter 'Remind today'; Condition (length > 0) → Slack managers channel. Then Catch (email only)."],
     paste=[("Cutoff days (Compose)", "createArray(int(outputs('Settings')?['CutoffDay1']), int(outputs('Settings')?['CutoffDay2']))"),
            ("This month cutoffs — Map", "formatDateTime(addDays(startOfMonth(outputs('Today')), sub(item(), 1)), 'yyyy-MM-dd')"),
            ("Next month cutoffs — Map", "formatDateTime(addDays(startOfMonth(addToTime(outputs('Today'), 1, 'Month')), sub(item(), 1)), 'yyyy-MM-dd')"),
            ("Remind today — From (fx)", "union(body('This_month_cutoffs'), body('Next_month_cutoffs'))"),
            ("Remind today — advanced", "@equals(formatDateTime(outputs('Today'), 'yyyy-MM-dd'), formatDateTime(addDays(addDays(item(), mul(int(outputs('Settings')?['ReminderBusinessDays']), -1)), if(equals(dayOfWeek(addDays(item(), mul(int(outputs('Settings')?['ReminderBusinessDays']), -1))), 6), -1, if(equals(dayOfWeek(addDays(item(), mul(int(outputs('Settings')?['ReminderBusinessDays']), -1))), 0), -2, 0))), 'yyyy-MM-dd'))"),
            ("Slack message (managers)", ":alarm_clock: *Payroll cutoff is @{formatDateTime(first(body('Remind_today')), 'dddd, MMM d')}.* Any extra shifts or additional pay to report for this pay period? Submit them as *XPAY — Extra Shift / Additional Pay*: <@{outputs('Settings')?['FormUrl']}|HR request form>")],
     check="T19 passes.")
step("S4-04", "Test", "Exception drill", "Flow 1", ["Run T20, then undo the change."], check="T20 = Pass.")
step("S4-05", "Operate", "Two weeks of steady operation", "Views + digest",
     ["Use the digest every morning; redirect DM/email requests to the form.", "End of week 2: review first-response times and SLA (All Open by SLA / Recently Closed)."],
     check="Gate G4 checklist complete.", owner="CJ")

# ================= S5
step("S5-01", "Downstream", "Jira ticket for offboarding", "New flow or a branch in Flow 1",
     ["When IT confirms: on OFF-V / OFF-I create → Jira 'Create a new issue' (access + equipment) → add the issue link to Internal Notes → comment.",
      "Set Auto = 'jira' on the 'Create Jira ticket' steps and tick them in the same flow."], check="A test OFF-V creates a Jira issue.", owner="CJ")
step("S5-02", "Downstream", "Offboarding distro email", "Flow 1 branch",
     ["When the distro address is confirmed: OFF-V → email the distro at intake. OFF-I stays manual (timing-sensitive)."], check="Test OFF-V emails the distro.", owner="CJ")
step("S5-03", "Downstream", "Word letters", "Word Online (Business) → Populate a Microsoft Word template (premium)",
     ["Termination letter (OFF-I), comp change (COMP) and job change (JOB) confirmations. Save to the HR file; attach to the item."], check="Generated letter opens correctly.", owner="CJ")
step("S5-04", "Downstream", "Monday.com payroll board + backfill link", "Monday connector",
     ["When confirmed: post OFF-V / OFF-I final pay and COMP changes to the payroll board; send the backfill link on OFF-V (or raise a REQ request)."], check="Item appears on the board.", owner="CJ")

# ================= S6
step("S6-01", "Run", "Dashboard", "List views → Power BI (optional)", ["Open by age, first response and SLA compliance, request-type breakdown."], owner="CJ")
step("S6-02", "Run", "In-Slack approvals", "Bot token", ["Replace email approvals with Slack buttons once the bot token is approved."], owner="CJ")
step("S6-03", "Run", "Paycor write-back", "Paycor iPaaS", ["Auto-write approved changes instead of manual entry."], owner="CJ")
step("S6-04", "Run", "Holiday-aware SLA", "Reference workbook", ["Add tblHolidays and skip those dates in Flow 1."], owner="CJ")
step("S6-05", "Run", "Email intake (optional)", "hr-ticketing-triage solution", ["If the team ever wants email intake back, add the message/conversation identifier columns and reuse the triage flow."], owner="CJ")

# ------------------------------------------------------------------ test pack
TESTS = [
    ("T01", "S1", "DATA — basic intake", "Submit the form: DATA, weekday.", "Item created by Flow 0; Request = 'DATA — Employee Data Update — <name>'; HRC-00000N; Priority Routine; Received Date; SLA = next business day; 📥 comment with 'Info needed'; Slack ping; confirmation email listing what's needed."),
    ("T02", "S1", "Weekend SLA", "Submit DATA on Saturday (or a Friday-evening test).", "Saturday submission → SLA = Tuesday."),
    ("T03", "S1", "DATA sub-steps", "Look at the T01 item.", "Checklist has 5 '[ ] 0n.' lines; 03–05 end with '⚙ resolution'; Next Action = '01. …'; 📋 comment lists the steps."),
    ("T04", "S1", "OTHER — catch-all", "Submit OTHER.", "No checklist; '⚠️ … needs manual scoping' comment; Slack warning."),
    ("T05", "S1", "Pick up, tick, close", "Assign yourself and set 1-Assigned; tick 01–02; set 3-Resolved.", "First Human Response stamped on pick-up; Checklist % 40; Next Action moves to 03; on close: Closed Date today, 03–05 [x], '✅ Resolved on <date> by <you>', close-the-loop email."),
    ("T06", "S1", "Live DATA", "Johanna submits and works a real request.", "Same as T01–T05 with no manual fixes."),
    ("T07", "S2", "XPAY — HR approval", "Submit XPAY (hourly $100, first date worked before this period's cutoff).", "2-Waiting / Pending / Approver; HR approval email; approve → Approved, step 01 [x], 🔓 comment, decision line in Internal Notes; Status 0-New (unassigned) or 1-Assigned."),
    ("T08", "S2", "XPAY — after cutoff", "Submit XPAY with first date worked in a period whose cutoff has passed.", "Intake comment and Slack show the pay period and ⚠️ after cutoff."),
    ("T09", "S2", "XPAY — rejected", "Submit XPAY; reject it.", "3-Resolved; Outcome Rejected; open steps [–]; requester email says it was closed as Rejected; NO Payroll FYI."),
    ("T10", "S2", "XPAY — Payroll FYI", "Close the T07 request.", "One Payroll room message with the details; editing the closed item again doesn't re-post."),
    ("T11", "S2", "LOA", "Submit LOA.", "SLA blank ('Per policy'); Assigned To = CJ; Slack shows '(confidential — see list)'; ticking every step doesn't close it."),
    ("T12", "S2", "OFF-V", "Submit OFF-V with Effective date = next Friday.", "SLA Due Date = that Friday; 16 steps; 'Terminate in Paycor' line shows 📅 date; Priority Time-sensitive."),
    ("T13", "S2", "OFF-I and ER (sensitive)", "Submit OFF-I and ER.", "Names withheld in Slack; Priority Urgent; ER routed to CJ with summary-only steps."),
    ("T14", "S3", "JOB — two stages", "Submit JOB (Manager email = you; Region with a Regional row).", "Manager approval first; Regional only after it; both lines in Internal Notes."),
    ("T15", "S3", "COMP — reject at stage 2", "Approve Manager, reject Regional.", "HR stage never sent; Rejected path as T09."),
    ("T16", "S3", "No approver", "Submit JOB with a Region that has no Regional row and the '*' row set Active = No.", "Approval Status = Needs approver; 2-Waiting; Slack alert. Restore the '*' row."),
    ("T17", "S3", "Next Action override", "Type your own Next Action, then tick one more step.", "Your text stays until the tick, then Next Action moves to the next open step; at 100% '☑️ … ready to resolve'; Status unchanged."),
    ("T18", "S4", "Daily digest", "Run Flow 6 with open test items (overdue, 0-New, waiting, an exception).", "One Slack post with the right sections; OFF-I / LOA / ER names show '(confidential)'."),
    ("T19", "S4", "Cutoff reminder", "Set CutoffDay1 to a date ReminderBusinessDays business days ahead; run Flow 6.", "Managers channel reminder with the form link. Restore the setting."),
    ("T20", "S4", "Exception drill", "Point Flow 1's Slack step at a channel the service account can't post to; submit DATA.", "Admin email with run link; Automation Note; item in Exceptions view and digest. Undo."),
    ("T21", "S0", "Permissions", "A non-HR colleague submits the form, then tries the list URL.", "Form works; list says they don't have access."),
    ("T22", "S3", "Config without flow edits", "Change DATA's SLABusinessDays to 2 in Excel; submit DATA.", "SLA = 2 business days. Restore to 1."),
]

ANNOUNCEMENT = (
    "📣 *New: one front door for HR requests*\n"
    "From today, every HR request — onboarding, job or pay changes, time off, leave, benefits, letters, questions, offboarding, extra shifts — "
    "goes through this form: <FormUrl>\n"
    "• You'll get a confirmation with a ticket reference, a target date and what we need from you.\n"
    "• If it needs approval, the approver gets it automatically.\n"
    "• You'll hear back when it's done.\n"
    "Requests sent by DM or email will be pointed back here — if it isn't in the intake, we can't track it. Thank you! — People Ops"
)
AUTO_REPLY = (
    "Thanks for reaching out to People Ops. HR requests are handled through our request form so nothing gets lost: <FormUrl>. "
    "For anything else, a member of the team will reply soon."
)
