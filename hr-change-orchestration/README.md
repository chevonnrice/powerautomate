# HR Change + Orchestration (v2: one SharePoint list)

The HR Change + Orchestration project plan (Project 10k, Steps 5, 5b, 6), rebuilt on Microsoft 365:

* **One SharePoint list**, *HR Change Requests*, is the system of record (replaces the HR Change Slack List).
* **One Excel workbook of reference tables** holds settings, SLAs, approval chains, the A1–I checklists and approvers.
  The flows read it, so you change behaviour by editing Excel, not flows.
* **Seven small Power Automate flows**: the plan's Flow 1 Intake Notification, Flow 2 Checklist Router, Flow 3 Payroll
  Notification and the Completion auto-comment, plus Approval Routing, Checklist Progress and a Daily Digest with the payroll cutoff reminder.
* **Slack for visibility**: People Ops room, Payroll room and managers channel posts, with the front door pinned in `#lane-people-talent`.

## The kit (`kit/`)

| File | Use it for |
|---|---|
| [`HR Change Blueprint.docx`](kit/HR%20Change%20Blueprint.docx) | **The blueprint to follow**: roadmap, sprints and gates, then every build step with what to click, what to name it and what to paste |
| [`HR Change Build Blueprint.xlsx`](kit/HR%20Change%20Build%20Blueprint.xlsx) | Tracker: start date → sprint dates, a status for each of the 69 steps, gate checklists, 22-case test pack, decisions, inspiration review, plan traceability, dependencies, launch comms |
| [`HR Change Reference.xlsx`](kit/HR%20Change%20Reference.xlsx) | The reference tables the flows read (`tblSettings`, `tblChangeTypes`, `tblChecklist`, `tblApprovers`), plus list fields, statuses and views for reference. Store it on the People Ops site |
| [`HR-Change-Requests.csv`](kit/HR-Change-Requests.csv) | Creates the one list (*+ New → List → From CSV*) |

## Roadmap at a glance

| Sprint | Focus | Gate |
|---|---|---|
| S0 (1 wk) | Reference workbook, the list, permissions, form, views, connections | G0: Ready to build |
| S1 (2 wks) | Crawl: Flow 1, 2, 4. Path **E** end to end | G1: Plumbing proven |
| S2 (2 wks) | Crawl complete: Template **I** pilot (HR approval, pay period/cutoff, Payroll FYI), G/A1/A2, front door launch | G2: Front door live |
| S3 (2 wks) | Walk: B/C/F approvals switched on from Excel after matrix sign-off, checklist progress, UAT | G3: Approvals live |
| S4 (2 wks) | Walk: daily digest, cutoff reminder, exception drill, two weeks of operation | G4: Operational |
| S5 | Walk: Jira, offboarding distro, Word letters, Monday board, each behind its dependency | G5: per item |
| S6 | Run: manager self-service form, dashboard, in-Slack approvals, Paycor iPaaS | ongoing |

## Regenerating the kit

All four files come from one data file, so they never disagree:

```bash
python kit/source/build_kit.py
```

Edit `kit/source/blueprint_data.py` (steps, tests, decisions, tables) and re-run.

## Earlier versions

`archive/v1-five-lists/` holds the first build: a five-list design with an importable solution. It's superseded by this one-list
design and kept for reference only.
