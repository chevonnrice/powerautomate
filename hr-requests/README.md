# HR Requests v1 (ship today)

The simple version of the HR Change + Orchestration plan: built entirely in the browser, no scripts, about 5 hours.

* **Two ways in, one queue:** a Microsoft Form, or an email to the HR mailbox. Auto-replies and no-reply senders are ignored;
  replies (to our acknowledgement, confirmation or close-out emails) join the existing request.
* **One SharePoint list with 13 columns:** Title, Category, Status, Priority, AssignedTo, RequesterEmail, EmployeeName,
  Description, DueDate, NextAction, Checklist, ClosedDate, ConversationId. Built-in ID is the ticket reference (HR-123);
  Comments are the work log.
* **One Excel table** (`tblTypes`): per request type, the SLA, default priority and owner, sensitive flag, close-the-loop,
  payroll FYI, "info we need", and the checklist (one step per line).
* **Four flows:** Setup (checklist, due date, owner, comment, Slack, confirmation), Close-out (closed date, ✅ comment,
  close-the-loop, payroll FYI), Email intake, Form intake. Optional Copilot filter for email if you have AI Builder credits.

## Files (`v1/`)

| File | Use it for |
|---|---|
| [`HR Requests v1 - Build Guide.docx`](v1/HR%20Requests%20v1%20-%20Build%20Guide.docx) | The step-by-step build (Parts A–C), tests, roadmap and plan coverage |
| [`HR Request Types.xlsx`](v1/HR%20Request%20Types.xlsx) | The table the flows read. Upload to the People Ops site |
| [`HR-Requests.csv`](v1/HR-Requests.csv) | Creates the list (*+ New → List → From CSV*) |

## Roadmap

| Release | What |
|---|---|
| v1 (today) | Everything above |
| v1.1 | Daily digest flow, first-response date, checklist %, pay-period flag, Copilot filter if licensed |
| v1.2 | Manager / Regional approvals once the approval matrix is signed off |
| v2 | Jira, Word letters, Monday payroll board, Paycor write-back, in-Slack approvals |

## Earlier designs

`archive/` keeps the five-list build (`five-lists`, packed solution only) and the full sprint kit (`sprint-kit`) for reference.
`v1/source/build_v1.py` regenerates the three files above. You don't need to run it to build the solution.
