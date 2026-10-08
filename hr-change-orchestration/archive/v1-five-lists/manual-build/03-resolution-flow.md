# Step 3: Resolution flow ("HR Change - 3 Resolution")

**What it does:** the moment someone sets **Status = 3-Resolved**, it:
* stamps **Completion Date** with today's date and posts `✅ Resolved on [date] by [assignee]` (your completion auto-comment)
* ticks the checklist steps the flow covers (*Close-the-loop*, *Mark list item complete*, *Log completion*). For Rejected/Withdrawn, it cancels whatever is still open
* sends the close-the-loop email (if the change type says so, and always for Rejected)
* for **I — Extra Shift**, posts the one-way FYI to the Payroll room ("Flow 3 — Payroll Notification" from your plan; not a gate)

**Create:** Automated cloud flow → `HR Change - 3 Resolution` → **SharePoint — When an item is created or modified** → list **HR Change Requests**

**Trigger conditions** (trigger → **Settings** → *Trigger conditions* → add both, each on its own line):
```
@equals(triggerBody()?['Status']?['Value'], '3-Resolved')
```
```
@empty(triggerBody()?['CompletionDate'])
```
These make it fire once per resolution. The flow fills Completion Date, so its own update doesn't re-trigger it.
To re-send, clear Completion Date and save.

## Actions

### 1. `Settings` (Compose)
```json
{ "timeZone": "Eastern Standard Time" }
```

### 2. `Get change type` + 3. `Config`: exactly as in Step 2 (#3, #4)

### 4. `Outcome` (Compose)
```
coalesce(triggerBody()?['Outcome']?['Value'], 'Completed')
```

### 5. `Resolver` (Compose)
```
coalesce(triggerBody()?['Assignee']?['DisplayName'], triggerBody()?['Editor']?['DisplayName'])
```

### 6. `Stamp completion` (SharePoint → Update item) → HR Change Requests, Id = trigger ID, required columns from the trigger
* CompletionDate: fx `formatDateTime(convertFromUtc(utcNow(), outputs('Settings')?['timeZone']), 'yyyy-MM-dd')`
* Outcome Value: fx `outputs('Outcome')`

### 7. `Get open tasks` (SharePoint → Get items) → **HR Change Tasks**
* Filter Query: `RequestId eq '@{triggerBody()?['RequestId']}' and TaskStatus ne 'Done' and TaskStatus ne 'Cancelled' and TaskStatus ne 'N/A'`
* Top Count: `500`

### 8. `Each open task` (Apply to each) over fx `outputs('Get_open_tasks')?['body/value']`
Inside: **Condition** `Closes with case`:
* Left: fx `or(not(equals(outputs('Outcome'), 'Completed')), equals(items('Each_open_task')?['Automation'], 'Resolution'))` · is equal to · fx `true`
* **True** → **Update item** `Close task` → **HR Change Tasks**, Id fx `items('Each_open_task')?['ID']`, Title fx `items('Each_open_task')?['Title']`,
  TaskStatus Value fx `if(equals(outputs('Outcome'), 'Completed'), 'Done', 'Cancelled')`

### 9. `Still open` (Data Operation → Filter array), after the loop
* From: fx `outputs('Get_open_tasks')?['body/value']`
* *Edit in advanced mode*: `@and(equals(outputs('Outcome'), 'Completed'), not(equals(item()?['Automation'], 'Resolution')))`

### 10. `Post resolved comment` (Send an HTTP request to SharePoint): same setup as Step 1 #15. Body:
```
{ "text": "✅ Resolved on @{formatDateTime(convertFromUtc(utcNow(), outputs('Settings')?['timeZone']), 'MMM d, yyyy')} by @{outputs('Resolver')}@{if(equals(outputs('Outcome'), 'Completed'), '', concat(' (Outcome: ', outputs('Outcome'), ')'))}@{if(greater(length(body('Still_open')), 0), concat(' — note: ', string(length(body('Still_open'))), ' checklist task(s) were still open'), '')}" }
```

### 11. `Close the loop?` (Condition)
* Left: fx `or(equals(outputs('Config')?['CloseLoopEmail'], true), not(equals(outputs('Outcome'), 'Completed')))` · is equal to · fx `true`
* **True** → **Send an email (V2)** `Close-the-loop email`
  * To: fx `coalesce(triggerBody()?['RequestedBy']?['Email'], triggerBody()?['Author']?['Email'])`
  * Subject: `@{outputs('Outcome')}: @{triggerBody()?['Title']} [@{triggerBody()?['RequestId']}]`
  * Body: `Your HR change request @{triggerBody()?['RequestId']} (@{triggerBody()?['ChangeType']?['Value']} for @{triggerBody()?['EmployeeName']}) is @{toLower(outputs('Outcome'))}.`

### 12. `Intake summary?` (Condition)
* Left: fx `outputs('Config')?['PostSummaryToIntake']` · is equal to · fx `true`
* **True** → Slack **Post message (V2)** → People Ops room:
  `:white_check_mark: *@{triggerBody()?['RequestId']}* @{triggerBody()?['ChangeType']?['Value']} — @{triggerBody()?['EmployeeName']} resolved by @{outputs('Resolver')} (@{outputs('Outcome')}).`

### 13. `Payroll FYI?` (Condition)
* Left: fx `and(equals(outputs('Config')?['NotifyPayrollOnResolve'], true), equals(outputs('Outcome'), 'Completed'))` · is equal to · fx `true`
* **True** → Slack **Post message (V2)** → **Payroll room**:
```
:moneybag: *Extra shift earning submitted in Paycor* — @{triggerBody()?['RequestId']}
Employee: @{triggerBody()?['EmployeeName']} (@{triggerBody()?['EmployeeId']})
Dates worked: @{coalesce(triggerBody()?['DatesWorked'], '-')} • Shift: @{coalesce(triggerBody()?['ShiftType']?['Value'], '-')}
Rate: @{coalesce(triggerBody()?['RateType']?['Value'], '-')} @{coalesce(string(triggerBody()?['RateAmount']), '')} • Hours/days: @{coalesce(string(triggerBody()?['HoursOrDays']), '-')}
Pay period: @{coalesce(triggerBody()?['PayPeriod'], '-')}@{if(equals(triggerBody()?['AfterCutoff'], true), ' :warning: submitted after cutoff', '')}
_FYI only — please verify the earning in Paycor. No response needed._
```

## Test: the full Phase 1 loop on the **E** request

1. Tick a couple of its tasks *Done* in HR Change Tasks (optional).
2. On the request, set **Assignee** to yourself and **Status = 3-Resolved** → Save.
3. Within ~1–2 minutes:
   * [ ] Completion Date = today, Outcome = Completed
   * [ ] The "Log completion / Close-the-loop / Mark complete" tasks are *Done*
   * [ ] Comment `✅ Resolved on <today> by <you>` (plus a note if the Paycor steps were still open)
   * [ ] Close-the-loop email to the submitter

✅ Phase 1 (crawl) works end to end. Then: [Step 4: Approvals](04-approvals-flow.md)
