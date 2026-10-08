# HR Change + Orchestration (v3: one form, one core-column list)

The HR Change + Orchestration project plan (Project 10k, Steps 5, 5b, 6), rebuilt on Microsoft 365:

* **One Microsoft Form** is the front door, pinned in `#lane-people-talent`.
* **One SharePoint list**, *HR Change Requests*, is the system of record. It keeps only core service-desk columns
  (ticket reference, category, priority, status, owner, requester, received, first response, next action + due, closed, internal notes).
  Everything specific to a request lives in its **Description** and its **sub-step Checklist**.
* **16 general HR request types** in five groups (onboarding, job and pay changes, time off, leave, benefits, data updates,
  letters, training, policy questions, employee relations, requisitions, offboarding, extra pay), with the original A1–I mapped in.
* **One Excel workbook of reference tables** holds settings, SLAs, approval chains, the sub-step checklists for every request type, and approvers.
  The flows read it, so you change behaviour by editing Excel, not flows.
* **Seven small Power Automate flows**: 0 Form Intake, 1 Intake Notification, 2 Checklist Router, 3 Case Updates, 4 Close-out
  (completion auto-comment + the plan's Payroll FYI), 5 Approval Routing, 6 Daily Digest + payroll cutoff reminder.
* **Slack for visibility**: People Ops room, Payroll room and managers channel posts, with the front door pinned in `#lane-people-talent`.

## The kit (`kit/`)

| File | Use it for |
|---|---|
| [`HR Change Blueprint.docx`](kit/HR%20Change%20Blueprint.docx) | **The blueprint to follow**: roadmap, sprints and gates, then every build step with what to click, what to name it and what to paste |
| [`HR Change Build Blueprint.xlsx`](kit/HR%20Change%20Build%20Blueprint.xlsx) | Tracker: start date → sprint dates, a status for each of the 60 steps, gate checklists, 22-case test pack, decisions, inspiration review, plan traceability, dependencies, launch comms |
| [`HR Change Reference.xlsx`](kit/HR%20Change%20Reference.xlsx) | The reference tables the flows read (`tblSettings`, `tblChangeTypes`, `tblChecklist`, `tblApprovers`), plus list fields, form questions, statuses and views for reference. Store it on the People Ops site |
| [`HR-Change-Requests.csv`](kit/HR-Change-Requests.csv) | Creates the one list with its core columns (*+ New → List → From CSV*) |

## Roadmap at a glance

| Sprint | Focus | Gate |
|---|---|---|
| S0 (1 wk) | Reference workbook, the list, permissions, form, views, connections | G0: Ready to build |
| S1 (2 wks) | Crawl: Flows 0–4. **DATA** (was E) end to end | G1: Plumbing proven |
| S2 (2 wks) | Crawl complete: **XPAY** (was I) pilot with HR approval and Payroll FYI, LOA / offboarding / ER, front door launch | G2: Front door live |
| S3 (2 wks) | Walk: JOB / COMP / PTO / REQ approvals switched on from Excel after matrix sign-off, UAT | G3: Approvals live |
| S4 (2 wks) | Walk: daily digest, cutoff reminder, exception drill, two weeks of operation | G4: Operational |
| S5 | Walk: Jira, offboarding distro, Word letters, Monday board, each behind its dependency | G5: per item |
| S6 | Run: dashboard, in-Slack approvals, Paycor iPaaS, holiday-aware SLA | ongoing |

## Regenerating the kit

All four files come from one data file, so they never disagree:

```bash
python kit/source/build_kit.py
```

Edit `kit/source/blueprint_data.py` (steps, tests, decisions, tables) and re-run.

## Earlier versions

`archive/v1-five-lists/` holds the first build: a five-list design with an importable solution. It's superseded by this one-list
design and kept for reference only.
