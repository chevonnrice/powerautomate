# HR Change + Orchestration: manual build, one step at a time

Use this instead of importing the solution .zip. You create each piece yourself in the Power Automate designer,
test it, and only then move on. Every step says exactly what to add, what to rename it to, and what to paste.

| Step | What you build | Test it with | Time |
|---|---|---|---|
| [0 Lists](00-lists.md) | 5 SharePoint lists from the CSV files in [`csv/`](csv) | Open the New form | 30 min |
| [1 Intake](01-intake-flow.md) | Request ID, title, SLA due date, comment, Slack ping, confirmation email | Submit an **E** request | 30 min |
| [2 Checklist router](02-checklist-router.md) | Copies the template steps into HR Change Tasks | Same **E** request → 5 tasks appear | 20 min |
| [3 Resolution](03-resolution-flow.md) | ✅ comment, Completion date, close-the-loop, Payroll FYI | Set the **E** request to 3-Resolved | 25 min |
| [4 Approvals](04-approvals-flow.md) | Manager → Regional → HR chain from HR Change Types | **I** request (HR approval), then **C** (3 stages + a rejection) | 45 min |
| [5 Task progress](05-task-progress-flow.md) | Tasks Done / Progress % on the request | Tick tasks on any request | 15 min |
| [6 Daily digest](06-daily-digest-flow.md) | SLA digest + payroll cutoff reminder | *Test → Manually* | 25 min |

Steps 1–3 give you the whole Phase 1 "crawl" loop: intake → checklist → resolve. Approvals, progress and the digest add on top.

## Conventions used in every step

**Where to build.** Power Automate → **My flows → + New flow → Automated cloud flow**, give it the name shown, and
pick the trigger listed. (To move the flows into a solution later, build them inside **Solutions → New solution** instead.)

**Rename every action** to the name in the step (click the action → click its title). The expressions refer to
actions by name, with spaces replaced by `_`. For example, an action renamed `Request ID` is read with `outputs('Request_ID')`.

**Pasting an expression.** Click in the field → click **fx** (the *Insert expression* button), paste the expression, and
click **Add**. Paste only the expression, without a leading `@`.

**Text with expressions inside.** When a field shows text like `HRC request @{outputs('Request_ID')} received`, type the
plain text yourself, and wherever you see `@{ ... }` insert the part inside the braces with **fx** at that spot.

**Filter Query fields** (SharePoint *Get items*) are text. Type them exactly, including the single quotes, and use **fx**
for the `@{ ... }` parts.

**SharePoint "Update item" asks for required columns.** For any required column the action asks for that the step
doesn't mention (Employee Name, Employee ID, Change Type, Effective Date, Manager), pick the **same field from the
trigger** in Dynamic content. That just writes back the value it already has.

**Settings compose.** Each flow starts with a Compose called `Settings` holding a little JSON (time zone, prefix,
channels). Change the values there, never in the expressions.

**Test loop.** Save → **Test → Manually** → then create or edit the list item. SharePoint triggers check about once a
minute, so give it 1–2 minutes. Open the run to see each action's inputs and outputs. If something fails, the red action
shows the exact error. Paste it to me.
