# Step 4: Approvals flow ("HR Change - 4 Approvals")

**What it does:** reads the approval chain for the change type from **HR Change Types** (ApprovalStage1 → 2 → 3) and
runs the approvals **one after another**, stopping at the first rejection. Seeded chains:

| Type | Chain |
|---|---|
| B — Job / Status | Manager → Regional |
| C — Comp | Manager → Regional → HR |
| F — Time / PTO | Manager |
| I — Extra Shift | HR |

* **Manager** = the request's **Manager** person field. **Regional / HR / Payroll** = the **HR Approval Matrix** row
  for that Role whose Region matches the request's Region, falling back to the Region `*` row.
* While waiting: Status `2-Waiting`, ApprovalStatus `Pending`, CurrentApprover `Regional: name@…`. Each decision is posted
  as a comment and added to *ApprovalHistory*.
* **Approved** → Status `1-Assigned` (or `0-New` if nobody owns it yet). "Route approval" tasks tick themselves, and
  "On approval: …" tasks go from *Blocked* to *Not started*.
* **Rejected** → Status `3-Resolved`, Outcome *Rejected*, open gated tasks *Cancelled*. Step 3 then emails the requester.
* Approvers answer from the approval email, Teams or the Power Automate app. The details are in the approval itself,
  so they don't need access to the list.

**Create:** Automated cloud flow → `HR Change - 4 Approvals` → **SharePoint — When an item is created** → **HR Change Requests**

## Actions

### 1. Two variables (Variables → Initialize variable)
* `Outcome`: Type **String**, Value `Approved`
* `History`: Type **String**, Value *(empty)*

### 2. `Settings`, `Request ID`, `Get change type`, `Config`: same as Step 2 (#1–#4)

### 3. `Stages` (Data Operation → Filter array)
* From: fx `createArray(outputs('Config')?['ApprovalStage1'], outputs('Config')?['ApprovalStage2'], outputs('Config')?['ApprovalStage3'])`
* *Edit in advanced mode*: `@not(empty(item()))`

### 4. `Any approvals?` (Condition)
* Left: fx `length(body('Stages'))` · is greater than · `0`
* **False** branch: leave empty (no approvals needed. ApprovalStatus stays *Not required*).

Everything below goes in the **True** branch.

### 5. `Each stage` (Apply to each) over fx `body('Stages')`
**⚠️ Required:** on the loop → **Settings** → **Concurrency control: On**, **Degree of parallelism: 1**. Without this, all
stages would ask at once.

Inside the loop, add one **Condition** `Still approved`: Left fx `variables('Outcome')` · is equal to · `Approved`.
Then put 5a–5f in its **True** branch:

**5a. `Get matrix rows`** (Get items) → **HR Approval Matrix**
* Filter Query: `Role eq '@{items('Each_stage')}' and Active eq 1`

**5b. `Matrix this region`** (Filter array). From fx `outputs('Get_matrix_rows')?['body/value']`, advanced:
`@equals(toLower(trim(coalesce(item()?['Region'], ''))), toLower(trim(coalesce(triggerBody()?['Region']?['Value'], ''))))`

**5c. `Matrix any region`** (Filter array). Same From, advanced: `@equals(trim(coalesce(item()?['Region'], '')), '*')`

**5d. `Approver`** (Compose)
```
if(equals(items('Each_stage'), 'Manager'), coalesce(triggerBody()?['Manager']?['Email'], ''), coalesce(first(body('Matrix_this_region'))?['ApproverEmail'], first(body('Matrix_any_region'))?['ApproverEmail'], ''))
```

**5e. `Approver found?`** (Condition): Left fx `empty(outputs('Approver'))` · is equal to · fx `false`

**True branch:**

1. **`Mark waiting`** (Update item) → HR Change Requests, Id = trigger ID, required columns from the trigger:
   Status Value `2-Waiting` · ApprovalStatus Value `Pending` · CurrentApprover `@{items('Each_stage')}: @{outputs('Approver')}`
2. **`Approval`** (Approvals → **Start and wait for an approval**)
   * Approval type: **Approve/Reject - First to respond**
   * Title: `@{items('Each_stage')} approval: @{triggerBody()?['ChangeType']?['Value']} — @{triggerBody()?['EmployeeName']} [@{outputs('Request_ID')}]`
   * Assigned to: fx `outputs('Approver')`
   * Details (Markdown works):
     ```
     **Employee:** @{triggerBody()?['EmployeeName']} (@{triggerBody()?['EmployeeId']})
     **Change type:** @{triggerBody()?['ChangeType']?['Value']}
     **Effective date:** @{formatDateTime(triggerBody()?['EffectiveDate'], 'MMM d, yyyy')}
     **New title / dept / location:** @{triggerBody()?['NewTitle']} @{triggerBody()?['NewDepartment']} @{triggerBody()?['NewLocation']}
     **New rate:** @{triggerBody()?['NewRate']} **Pay period:** @{triggerBody()?['EffectivePayPeriod']}
     **Extra shift:** @{triggerBody()?['DatesWorked']} @{triggerBody()?['ShiftType']?['Value']} @{triggerBody()?['RateType']?['Value']} @{triggerBody()?['RateAmount']}
     **Notes:** @{triggerBody()?['Notes']}
     ```
   * Item link: fx `triggerBody()?['{Link}']` · Item link description: `Open @{outputs('Request_ID')}`
3. **`Add to history`** (Variables → Append to string variable) → `History`, Value:
   ```
   @{items('Each_stage')}: @{if(equals(outputs('Approval')?['body/outcome'], 'Approve'), 'Approved', 'Rejected')} by @{first(outputs('Approval')?['body/responses'])?['responder']?['displayName']} — @{first(outputs('Approval')?['body/responses'])?['comments']}
   ```
   (press Enter at the end of the value so each stage gets its own line)
4. **`Post decision comment`** (Send an HTTP request to SharePoint, set up like Step 1 #15). Body:
   ```
   { "text": "@{items('Each_stage')} @{if(equals(outputs('Approval')?['body/outcome'], 'Approve'), 'approved 👍', 'rejected 👎')} by @{first(outputs('Approval')?['body/responses'])?['responder']?['displayName']}" }
   ```
5. **`Rejected?`** (Condition): Left fx `outputs('Approval')?['body/outcome']` · is equal to · `Reject`
   → **True**: **Set variable** `Outcome` = `Rejected`

**False branch** (no approver found):
1. **Set variable** `Outcome` = `NoApprover`
2. Slack **Post message (V2)** → People Ops room:
   `:rotating_light: @{outputs('Request_ID')} needs a *@{items('Each_stage')}* approver but none was found (Manager empty, or no HR Approval Matrix row).`

### 6. After the loop, still in the True branch of `Any approvals?`

**6a. `Get request now`** (SharePoint → Get item) → HR Change Requests, Id = trigger ID

**6b. `Save approval result`** (Update item) → HR Change Requests, Id = trigger ID, required columns from the trigger:

| Field | Value |
|---|---|
| ApprovalStatus Value | fx `if(equals(variables('Outcome'), 'Approved'), 'Approved', if(equals(variables('Outcome'), 'Rejected'), 'Rejected', 'Needs approver'))` |
| Status Value | fx `if(equals(variables('Outcome'), 'Rejected'), '3-Resolved', if(equals(variables('Outcome'), 'NoApprover'), '2-Waiting', if(empty(outputs('Get_request_now')?['body/Assignee']), '0-New', '1-Assigned')))` |
| Outcome Value | fx `if(equals(variables('Outcome'), 'Rejected'), 'Rejected', outputs('Get_request_now')?['body/Outcome']?['Value'])` |
| CurrentApprover | fx `''` |
| ApprovalHistory | fx `variables('History')` |

**6c. `Decision made?`** (Condition): Left fx `variables('Outcome')` · is not equal to · `NoApprover` → **True**:
* **`Get gated tasks`** (Get items) → **HR Change Tasks**, Filter Query:
  `RequestId eq '@{outputs('Request_ID')}' and (TaskStatus eq 'Blocked - awaiting approval' or Automation eq 'Approval')`
* **`Each gated task`** (Apply to each over fx `outputs('Get_gated_tasks')?['body/value']`) → **Update item** → HR Change Tasks:
  Id fx `items('Each_gated_task')?['ID']` · Title fx `items('Each_gated_task')?['Title']` · TaskStatus Value:
  ```
  if(equals(items('Each_gated_task')?['Automation'], 'Approval'), 'Done', if(equals(variables('Outcome'), 'Approved'), 'Not started', 'Cancelled'))
  ```

## Test (in your build-guide order)

1. **I — Extra Shift** (the live pilot): make sure the HR row in HR Approval Matrix has your email, then submit an I request
   (e.g. the DVA case: Hourly, $100, first date worked Aug 15).
   * [ ] Status `2-Waiting`, ApprovalStatus `Pending`, an approval email arrives
   * [ ] Approve → ApprovalStatus `Approved`, the "HR approval" task is *Done*, "Earning setup" goes from *Blocked* to *Not started*
   * [ ] Resolve it → Step 3 posts the **Payroll FYI** in the Payroll room
2. **C — Comp**: set Manager to yourself and point the Regional/HR rows at yourself (or test colleagues). You'll get 3
   approvals one after another. Approve the first, **reject** the second:
   * [ ] The third approval never arrives · Status `3-Resolved`, Outcome `Rejected` · gated tasks *Cancelled* · rejection email to the requester

✅ Then: [Step 5: Task progress](05-task-progress-flow.md)
